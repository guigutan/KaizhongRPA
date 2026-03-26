using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SAPFEWSELib;
using static System.Runtime.CompilerServices.RuntimeHelpers;


namespace KaizhongRPA
{
    public class Rpa005 : InvokeCenter
    {
        public RpaInfo GetThisRpaInfo()
        {
            RpaInfo rpaInfo = new RpaInfo();
            rpaInfo.RpaClassName = this.GetType().Name;
            rpaInfo.RpaName = "工艺路线导入流程";
            rpaInfo.DefaultRunTime1 = "****-**-** **:**:**";
            rpaInfo.DefaultRunTime2 = "****-**-** **:**:**";
            rpaInfo.DefaultStatus = "有效";
            rpaInfo.DefaultPathStype = "相对路径";
            rpaInfo.DefaultConfigPath = @"config\RpaGroup\工艺路线导入.xlsx";
            return rpaInfo;
        }

        #region 头部声明
        PublicClass publicClass = new PublicClass();
        public IWebDriver MyDriver;
        public WebDriverWait wait60;
        public WebDriverWait wait30;
        public WebDriverWait wait10;
        public WebDriverWait wait3;
        public string OaDownDir = $@"{MyPath.Documents}\{typeof(MyPath).Namespace}\WL015\OaDown\";
        public string SapResTempDir = $@"{MyPath.Documents}\{typeof(MyPath).Namespace}\WL015\SapResTemp\";
        public string SapResRequestidDir = $@"{MyPath.Documents}\{typeof(MyPath).Namespace}\WL015\SapResRequestid\";

      

        public DataTable dt_config;
        #endregion

        public async Task RpaMain(CancellationToken token, RpaInfo rpaInfo)
        {
            string WechatKey = "";          
            string OutFolderSap = "";
            try
            {
                if (!Directory.Exists(OaDownDir)) { Directory.CreateDirectory(OaDownDir); }
                if (!Directory.Exists(SapResTempDir)) { Directory.CreateDirectory(SapResTempDir); }
                if (!Directory.Exists(SapResRequestidDir)) { Directory.CreateDirectory(SapResRequestidDir); }


                dt_config = await publicClass.ExcelToDataTable(token, rpaInfo.DefaultConfigPath);
                if (!(dt_config != null && dt_config.Rows.Count > 0)) { return; }
                WechatKey = await publicClass.GetLibValue(token, dt_config, "WechatKey");
                string OutFolder = await publicClass.GetLibValue(token, dt_config, "OutFolder");
                OutFolderSap = OutFolder + DateTime.Now.ToString("yyyy-MM-dd")+@"\";
                if (!Directory.Exists(OutFolderSap)) { Directory.CreateDirectory(OutFolderSap); }

            }
            catch (Exception ex)
            {
                await publicClass.NoteLog(token, ex, dt_config);
                return;
            }
           

            bool isExistPending = true;
            while (isExistPending)
            {
                isExistPending = false;
                string temp_Requestid = "";
                string requestnamespan = "";
                string ItemNO = "";


                bool temp_AllS = false;
                string temp_WechatPostFilePath = "";
                try
                {                 

                    await publicClass.DisableScreen(token);

                    #region 浏览器的Options设置
                    ChromeOptions Options = new ChromeOptions();
                    Options.BinaryLocation = MyWeb.Chrome;
                    Options.EnableDownloads = true;
                    string Argument = await publicClass.GetLibValue(token, dt_config, "OaLoginUrl");
                    Options.AddArgument("--test-type"); //禁用沙盒
                    Options.AddArgument("--no-sandbox");//禁用沙盒
                    Options.AddExcludedArgument("enable-automation");   //去除状态栏【正在受到自动软件的控制】的提示 
                    Options.AddArgument("--start-maximized");           //窗口最大化             
                    Options.AddArgument("disable-blink-features=AutomationControlled");             //覆盖window.navigator.webdriver的值。以免跳出登陆的滑块验证
                    Options.AddArgument($"--unsafely-treat-insecure-origin-as-secure={Argument}");  //已阻止不安全的下载                
                    Options.AddUserProfilePreference("download.default_directory", OaDownDir);      //设置下载目录
                    Options.AddUserProfilePreference("download.prompt_for_download", false);        //禁用下载提示
                    Options.AddUserProfilePreference("download.directory_upgrade", true);           //确保使用指定目录
                    Options.AddUserProfilePreference("credentials_enable_service", false);          //禁止弹出密码保存对话框
                    Options.AddUserProfilePreference("profile.password_manager_enabled", false);    //禁止弹出密码保存对话框
                    #endregion

                    //网页处理（SAP处理包含在网页处理中）         
                    using (MyDriver = new ChromeDriver(MyWeb.ChromeDriver, Options))
                    {
                        wait60 = new WebDriverWait(MyDriver, TimeSpan.FromSeconds(60));
                        wait30 = new WebDriverWait(MyDriver, TimeSpan.FromSeconds(30));
                        wait10 = new WebDriverWait(MyDriver, TimeSpan.FromSeconds(10));
                        wait3 = new WebDriverWait(MyDriver, TimeSpan.FromSeconds(1));

                        //----------------------------------------------------------------------------------------

                        await LoginOA(token, dt_config);                //01-登录OA
                        await Nav_liucheng(token);                      //02-选择顶部菜单《流程》
                        await Li_daiban(token);                         //03-选择左侧菜单《待办事宜》
                        await SwitchTo_mainFrame(token);                //04-切换IFrame
                        isExistPending = await Li_WL015(token);         //05-是否存在待办列表>>>点击左侧WL015类流程菜单
                        if (!isExistPending) { continue; }


                        //----------------------------------------------------------------------------------------
                        await publicClass.DisableScreen(token);

                        await SwitchTo_myFrame(token);                  //01-切换IFrame  
                        await SwitchTo_tabframe(token, @"WFSearchResult.jsp");//02-切换IFrame
                        await ClickItem(token);                         //03-点开一行待办项
                        await SwitchTo_WindowMRNF(token);               //04-切换新窗口


                        //----------------------------------------------------------------------------------------
                        /*获取 requestnamespan*/
                        requestnamespan = await GetRequestmarkSpan(token);
                        ItemNO = await GetItemNO(token);
                     

                        /*获取 requestid*/
                        string url = MyDriver.Url.ToLower();
                        string requestid = url.Contains("requestid") ? url.Substring(url.IndexOf("requestid=") + "requestid=".Length) : "";
                        temp_Requestid = requestid = requestid != "" ? requestid.Substring(0, requestid.IndexOf("&")) : "";
                        if (requestid == "") { throw new Exception("从网址中获取requestid失败"); }
                        //----------------------------------------------------------------------------------------

                        await PostWillDo(requestid, requestnamespan, ItemNO, WechatKey);//存在待办列表>>>>>>POST 执行前的预告


                        //----------------------------------------------------------------------------------------
                       
                        await ClearDir(token, OaDownDir);               //06-下载前清空下载目录
                        await SelectDownload(token);                    //07-点击下载附件
                        string filePath = await IsFinishDown(token);    //08-判断是否已经下载完成（向上抛异常）
                                              

                       
                        string SapResFile = SapResRequestidDir + requestid + ".xlsx"; //待优化。限定附件为.xlsx忽略.xls
                        temp_WechatPostFilePath = OutFolderSap + requestid + ".xlsx";
                        //----------------------------------------------------------------------------------------



                        if (!File.Exists(SapResFile))
                        {
                            //OA下载的Excel工艺路线 导入SAP ，导入的结果为：SapResFile
                            try
                            {
                                await publicClass.GotoSapHome(token, dt_config);//01-登录SAP
                                await EnterTransaction(token, "ZPPB002");       //02-进入ZPPB002事务
                                await InputExcle(token, filePath);              //03-工艺路线期初批导(*********)
                                string errorMsg = await InputErrorMsg(token);     //04-检查是否导入失败
                                if (errorMsg != "")
                                {
                                    await publicClass.ExitSap(token);
                                    await OA_Transfer(token, "", errorMsg);
                                    throw new Exception($"{errorMsg}");//失败，提交转办后，跳出当前迭代
                                }

                                await publicClass.DisableScreen(token);

                                await ClearDir(token, SapResTempDir);           //01-导出前清空目录
                                string outputRes = await OutputExcleRes(token); //02-导出SAP结果excel
                                await publicClass.CloseExcle(token);            //关闭打开的excel文件 // await KillExcel(token);
                              

                                File.Copy(outputRes, SapResFile, true);         //待优化。如果outputRes不是.xlsx格式，与SapResFile扩展名不一致，复制强制改后缀，则后续文件读取失败
                                await Task.Delay(3000, token); token.ThrowIfCancellationRequested();
                                File.Copy(outputRes, temp_WechatPostFilePath, true); //复制一份给输出
                                await Task.Delay(3000, token); token.ThrowIfCancellationRequested();
                              
                            }
                            catch (Exception ex) { await publicClass.NoteLog(token, ex, dt_config); continue; }
                            finally { await publicClass.ExitSap(token); }                            
                        }

                        if (!File.Exists(SapResFile)) { await OA_Transfer(token, "", "ZPPB002工艺路线导入SAP时发生异常"); continue; }



                        //----------------------------------------------------------------------------------------
                        //----------------------------------------------------------------------------------------
                        

                        bool AllS = temp_AllS= await Check_SE(token, SapResFile);//判断状态(S\E)
                        if (AllS)
                        {
                            try { await OA_Approve(token, SapResFile); }//异常时 ，转办                           
                            catch (Exception ex) { await OA_Transfer(token, SapResFile, $"全S状态即将批准但创建生成版本时异常{ex.Message}"); }
                        }
                        else {
                            await OA_Transfer(token, SapResFile, ""); //非全S，转办
                        }
                       

                        //-----结束，关闭网页重来-----------------------------------------------------------------------------------
                        //待优化：重来时跳过异常的OA流程，点击下一个OA流程(记录requestid)


                        if (isExistPending) { await PostResInfo(temp_Requestid, requestnamespan,ItemNO, temp_AllS, temp_WechatPostFilePath, WechatKey,""); } //POST 执行后的结果  成功

                    }
                }
                catch (Exception ex)
                {
                    if (isExistPending) { await PostResInfo(temp_Requestid, requestnamespan, ItemNO, temp_AllS, temp_WechatPostFilePath, WechatKey,$"{ex.Message}"); } //POST 执行后的结果  异常

                    await publicClass.NoteLog(token, ex, dt_config);
                }
                finally
                {
                    await publicClass.ExitSap(token);
                }


            }

        }

       




        public async Task OA_Approve(CancellationToken token, string SapResFile)
        {
            try
            { 
                //----创建生成版本------------------------------------------------------------------------------------
                await publicClass.GotoSapHome(token, dt_config);//01-登录SAP
                await EnterTransaction(token, "ZPPB072");       //02-进入ZPPB072事务

                string text = await GetItemNOText(token, SapResFile);   //获取所有物料号                                                                                              
                await PublicClass.SetClipboardTextAsync(text);          //复制到剪贴板

                await EnterFromClipboard(token);            //02-点击多项选择，粘贴物料号
                await JobAndF8(token);                      //03-勾选JOB，并执行
                await ClickSave(token);                     //04-点击保存
                await publicClass.ExitSap(token);           //05-退出SAP

                //----OA批准------------------------------------------------------------------------------------
                MyDriver.SwitchTo().DefaultContent();
                await ClickOaHandle(token, "批准");     //01-点击批准
                await ResAndSwitchToWindow(token);      //02-提交结果处理并切换window
            }
            catch (Exception ex) { throw ex; }

        }
        public async Task OA_Transfer(CancellationToken token, string SapResFile, string ErrorMsg)
        {
            try
            {    
                //----OA转办附上表格------------------------------------------------------------------------------------
                MyDriver.SwitchTo().DefaultContent();
                await ClickOaHandle(token, "转办");     //05-点击转办
                await SwitchTo_DialogFrame(token, @"workflow/request/Remark.jsp");          //06-切换IFrame
                await SwitchTo_tabframe(token, @"workflow/request/RemarkFrame.jsp"); //07-切换IFrame
                await ClickIcoUser(token);                                                  //08-点击用户图标按钮


                MyDriver.SwitchTo().DefaultContent();
                await SwitchTo_DialogFrame(token, @"systeminfo/BrowserMain.jsp");               //01-切换IFrame
                await SwitchToIFrame_ByID(token, "main", @"hrm/resource/ResourceBrowser.jsp");  //02-切换IFrame
                await SwitchToIFrame_ByID(token, "frame1", @"hrm/resource/Select.jsp");         //03-切换IFrame                
                await EnterUserName(token, dt_config);                                             //04-输入转办人
                await SelectUserName(token, dt_config);                                            //05-点击选中转办人


                MyDriver.SwitchTo().DefaultContent();
                await SwitchTo_DialogFrame(token, @"workflow/request/Remark.jsp");                          //01-切换IFrame
                await SwitchToIFrame_ByID(token, "tabcontentframe", @"workflow/request/RemarkFrame.jsp");   //02-切换IFrame
                if (SapResFile != "")
                {
                    await ClickUploadICO(token);                                    //03-点击上传附件图标
                    await publicClass.WinApi_ChooseFile32770(token, SapResFile);    //04-选择附件
                    await ClickSelectFile(token);                                   //05-点选刚刚上传的附件
                }
                if (ErrorMsg != "")
                {
                    //把异常信息写到签字出
                    await SwitchToIFrame_ByID(token, "ueditor_0", @"");   //02-切换IFrame
                    IWebElement viewP = MyDriver.FindElement(By.XPath($"//body[@class='view']/p"));
                    viewP.SendKeys($"{ErrorMsg}");
                }               

                MyDriver.SwitchTo().DefaultContent();                               //01-切换IFrame
                await SwitchTo_DialogFrame(token, @"workflow/request/Remark.jsp");  //02-切换IFrame
                await SubTransfer(token);                                           //03-提交转办
                await ResAndSwitchToWindow(token);                                  //04-提交结果处理并切换window
            }
            catch (Exception ex) { throw ex; }


        }


        #region  Switch 相关
        //originalWindow
        private async Task SwitchToWindow(CancellationToken token, List<string> originalWindow, int timeout = 60)
        {
            try
            {
                //等待新窗口打开
                var windowHandles = MyDriver.WindowHandles;
                for (int t = 0; t < timeout; t++)
                {
                    if (windowHandles.Count > originalWindow.Count) { break; }
                    windowHandles = MyDriver.WindowHandles;
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    if (t + 1 == timeout) { throw new Exception("未能等待到打开的新窗口"); }
                }

                //切换新窗口
                bool res = false;
                foreach (string windowHandle in windowHandles)
                {
                    if (!originalWindow.Contains(windowHandle))
                    {
                        MyDriver.SwitchTo().Window(windowHandle);
                        res = true;
                        break;
                    }
                }
                if (!res) { Console.WriteLine("无法切换新窗口"); throw new Exception("无法切换新窗口"); }


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private async Task SwitchToWindow(CancellationToken token, string url, int timeout = 60)
        {
            try
            {
                Console.WriteLine($"进入SwitchToWindow{url}");
                bool isSwitch = false;
                await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                var windowHandles = MyDriver.WindowHandles;
                foreach (var wh in windowHandles)
                {
                    MyDriver.SwitchTo().Window(wh); //先切换
                    if (MyDriver.Url.ToUpper().Contains($"{url}".ToUpper())) //再判断
                    {
                        isSwitch = true;
                        return;
                    }
                }
                if (isSwitch) { throw new Exception($"未能切换窗口{url}"); }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private async Task SwitchToDefaultContent(CancellationToken token)
        {
            try
            {
                await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                MyDriver.SwitchTo().DefaultContent();
            }
            catch (Exception ex) { throw ex; }
        }

        private async Task SwitchToIframe(CancellationToken token, By by, int timeout = 60)
        {
            try
            {
                var iframes = MyDriver.FindElements(by);//切换窗口或者有页面变化改变DOM的>>>取消Displayed，使用FindElements带s的，可以等待，且异常不抛出  
                for (int t = 0; t < timeout; t++)
                {
                    if (iframes.Count > 0) { MyDriver.SwitchTo().Frame(iframes[0]); return; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    iframes = MyDriver.FindElements(by);
                    if (t + 1 == timeout) { throw new Exception($"未能切换iframe>>{by}"); }
                }
            }
            catch (Exception ex) { throw ex; }
        }

        private async Task SwitchToIframe(CancellationToken token, By by, string documentUrl, int timeout = 60)
        {
            try
            {
                var iframes = MyDriver.FindElements(by); //切换窗口或者有页面变化改变DOM的>>>取消Displayed，使用FindElements带s的，可以等待，且异常不抛出      
                for (int t = 0; t < timeout; t++)
                {
                    foreach (var iframe in iframes)
                    {
                        string url = (string)((IJavaScriptExecutor)MyDriver).ExecuteScript("return arguments[0].contentDocument.URL;", iframe); //#document
                        if (url.ToUpper().Contains(documentUrl.ToUpper()))
                        {
                            MyDriver.SwitchTo().Frame(iframe);
                            return;
                        }
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    iframes = MyDriver.FindElements(by);
                    if (t + 1 == timeout) { throw new Exception($"未能切换iframe>>{by}"); }
                }
            }
            catch (Exception ex) { throw ex; }
        }


        #endregion


        #region OA相关       
        //--------------------------------------------------------------------------------------------------
        public async Task LoginOA(CancellationToken token, DataTable dt_config)
        {
            try
            {
                string OaLoginUrl = await publicClass.GetLibValue(token, dt_config, "OaLoginUrl");
                string OaUser = await publicClass.GetLibValue(token, dt_config, "OaUser");
                string OaPassWorld = await publicClass.GetLibValue(token, dt_config, "OaPassWorld");
                if (OaLoginUrl == "" || OaUser == "" || OaPassWorld == "") { throw new Exception($"配置表中的OA相关信息为空字符"); }               
                MyDriver.Navigate().GoToUrl(OaLoginUrl);
                // await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                wait60.Until(d => d.FindElement(By.Id("loginid")).Displayed);
                MyDriver.Manage().Window.Maximize();
                 await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                MyDriver.FindElement(By.Id("loginid")).SendKeys(OaUser);
                MyDriver.FindElement(By.Id("userpassword")).SendKeys(OaPassWorld);
                MyDriver.FindElement(By.Id("login")).Click();
                 await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                wait60.Until((d) => MyDriver.FindElement(By.Id("lftop")));
               
                //处理登录后可能存在的提示框
                var btns = wait3.Until(d => d.FindElements(By.CssSelector(".zd_btn_cancle.btn_submit")));
                for (int wi = 0; wi < 10; wi++)
                {
                    btns = wait3.Until(d => d.FindElements(By.CssSelector(".zd_btn_cancle.btn_submit")));                   
                    if (btns.Count > 0) { btns[0].Click();break; }
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                }
            }
            catch (Exception ex) { throw ex; }
        }
        private async Task Nav_liucheng(CancellationToken token)
        {
            try
            {
                var divs = wait60.Until(d => d.FindElements(By.XPath("//div[@class='slideItemText' and text()='流程']")));
                foreach (IWebElement div in divs)
                {
                     await Task.Delay(200, token);token.ThrowIfCancellationRequested(); 
                    IWebElement parent = div.FindElement(By.XPath(".."));                    
                    if (parent.GetDomAttribute("title") == "我的流程")
                    {
                        parent.Click();
                         await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                        break;
                    }
                }              
            }
            catch (Exception ex)
            {               
                throw ex;
            }
        }
        private async Task Li_daiban(CancellationToken token)
        {
            try
            {
                IWebElement drillmenu = wait60.Until((d) => MyDriver.FindElement(By.Id("drillmenu")));
                var liCss2s = drillmenu.FindElements(By.CssSelector("li.liCss2"));
                foreach (var li in liCss2s)
                {
                     await Task.Delay(200, token);token.ThrowIfCancellationRequested(); 
                    var a = li.FindElement(By.TagName("a"));
                    if (a.GetDomAttribute("href").Contains("RequestView.jsp"))
                    {
                        a.Click();
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private async Task SwitchTo_mainFrame(CancellationToken token)
        {
            try
            {               
                for (int wi = 0; wi <= 30; wi++)
                {                   
                    IWebElement mainFrame = wait60.Until((d) => MyDriver.FindElement(By.Id("mainFrame")));
                    string iframeUrl = (string)((IJavaScriptExecutor)MyDriver).ExecuteScript("return arguments[0].contentDocument.URL;", mainFrame);
                    bool b = iframeUrl.Contains(@"workflow/request/RequestView.jsp");
                    if (b) { MyDriver.SwitchTo().Frame(mainFrame); break; }
                    if (!b && wi == 30) { throw new Exception("未找到对应的iframe，《RequestType.jsp》"); }
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private async Task<bool> Li_WL015(CancellationToken token)
        {
            bool result = false;
            try
            {
                var myas = wait60.Until(d => d.FindElements(By.XPath("//a[@class='level1 e8menu']")));
                for (int wi = 0; wi < 30; wi++)
                {
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                    myas = wait60.Until(d => d.FindElements(By.XPath("//a[@class='level1 e8menu']")));
                    if (myas.Count > 0) { break; }
                }
                if (!(myas.Count > 0)) { throw new Exception("未找到//a[@class='level1 e8menu']"); }

                foreach (var mya in myas)
                {
                    token.ThrowIfCancellationRequested(); await Task.Delay(10, token);
                    if (mya.GetDomAttribute("title").Contains("WL015-SAP工艺路线导入"))
                    {
                        result = true;
                        mya.Click();
                        await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                        break;
                    }
                }
            }
            catch
            {
                result = false;               
            }
            return result;
        }

        private async Task SwitchTo_myFrame(CancellationToken token)
        {
            try
            {               
                for (int wi = 0; wi <= 30; wi++)
                {                    
                    IWebElement myFrame = wait60.Until((d) => d.FindElement(By.Id("myFrame")));
                    string iframeUrl = (string)((IJavaScriptExecutor)MyDriver).ExecuteScript("return arguments[0].contentDocument.URL;", myFrame);
                    bool b = iframeUrl.Contains(@"workflow/search/wfTabFrame.jsp");
                    if (b) { MyDriver.SwitchTo().Frame(myFrame); break; }
                    if (!b && wi == 30) { throw new Exception("未找到对应的iframe，《wfTabFrame.jsp》"); }
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private async Task SwitchTo_tabframe(CancellationToken token,string urlLike)
        {
            try
            {
                bool isfinded = false;               
                for (int wi = 0; wi <= 30; wi++)
                {
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                    var iframes = wait60.Until(d => d.FindElements(By.XPath("//iframe[@name='tabcontentframe']")));
                    if (iframes.Count <= 0) { continue; }
                    foreach (var iframe in iframes)
                    {
                        string url = (string)((IJavaScriptExecutor)MyDriver).ExecuteScript("return arguments[0].contentDocument.URL;", iframe); //#document
                        if (url.ToUpper().Contains(urlLike.ToUpper())) { isfinded = true; MyDriver.SwitchTo().Frame(iframe); return; }
                    }                    
                }
                if (!isfinded) { throw new Exception($"未找到对应的iframe，《{urlLike}》"); }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private async Task ClickItem(CancellationToken token,int tryAgain=0)
        {
            try
            {
                string xpathToFind = "//form[@id='weaver']//div[@id='_xTable']//div[@class='table']//table[@class='ListStyle']//tbody//tr[@style='vertical-align: middle;']";
                var  trs = wait60.Until(d => d.FindElements(By.XPath(xpathToFind)));
                for (int wi = 0; wi < 30; wi++)
                {
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                    trs = wait60.Until(d => d.FindElements(By.XPath(xpathToFind)));
                    if (trs.Count > 0) { break; }
                }
                if (!(trs.Count > 0)) {throw new Exception($"未找到{xpathToFind}"); }


                foreach (IWebElement tr in trs)
                {
                    var tds = wait60.Until(d => tr.FindElements(By.XPath(".//td")));
                    foreach (var td in tds)
                    {
                        if (td.GetDomAttribute("title") != null && td.GetDomAttribute("title").ToString().Contains("WL015-SAP工艺路线"))
                        {
                            IWebElement a = wait60.Until(d => td.FindElement(By.XPath(".//a"))); //改为用a点击
                            a.Click();
                             await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                            var windowHandles = MyDriver.WindowHandles;
                            if (windowHandles.Count != 2 && tryAgain < 5) { await ClickItem(token, tryAgain + 1); }
                            return;
                        }
                    }                   
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}\r\n{ex.StackTrace}");
                throw ex;//没有可点击的项 退出方式直接为抛异常方式
            }
        }
        private async Task SwitchTo_WindowMRNF(CancellationToken token)
        {
            try
            {
                bool findWinows = false;
                for (int wi = 0; wi <= 60; wi++)
                {                    
                    var windowHandles = MyDriver.WindowHandles;                                     
                    foreach (var wh in windowHandles)
                    {
                        MyDriver.SwitchTo().Window(wh);                        
                        if (MyDriver.Url.ToUpper().Contains(@"ManageRequestNoForm.jsp".ToUpper()))
                        {
                            findWinows = true;
                            return;
                        }                        
                    }
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                }

                if (!findWinows) { throw new Exception("未找到弹出的新窗口"); }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private async Task SwitchTo_bodyiframe(CancellationToken token)
        {
            try
            {
                for (int wi = 0; wi <= 30; wi++)
                {
                    IWebElement myFrame = wait60.Until((d) => d.FindElement(By.Id("bodyiframe")));
                    string iframeUrl = (string)((IJavaScriptExecutor)MyDriver).ExecuteScript("return arguments[0].contentDocument.URL;", myFrame);
                    bool b = iframeUrl.Contains(@"ManageRequestNoFormIframe.jsp");
                    if (b) { MyDriver.SwitchTo().Frame(myFrame); break; }
                    if (!b && wi == 30) { throw new Exception("未找到对应的iframe，《ManageRequestNoFormIframe.jsp》"); }
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private async Task <string> GetRequestmarkSpan(CancellationToken token,int timeout=60)
        {
            string requestnamespan = "";
            try
            {
                MyDriver.SwitchTo().DefaultContent();              
                await SwitchToIframe(token, By.XPath("//iframe[@id='bodyiframe']"));


                var spans = MyDriver.FindElements(By.XPath("//span[@id='requestnamespan']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (spans.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    MyDriver.SwitchTo().DefaultContent();
                    await SwitchToIframe(token, By.XPath("//iframe[@id='bodyiframe']"));
                    spans = MyDriver.FindElements(By.XPath("//span[@id='requestnamespan']"));
                    if (t + 1 >= timeout) { throw new System.Exception($"未能找到元素//span[@id='requestnamespan']"); }
                }

                requestnamespan = spans[0].Text;

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return requestnamespan;
        }

        private async Task<string> GetItemNO(CancellationToken token, int timeout = 60)
        {
            string ItemNO = "";
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.XPath("//iframe[@id='bodyiframe']"));


                var spans = MyDriver.FindElements(By.XPath("//span[@id='field101821span']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (spans.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    MyDriver.SwitchTo().DefaultContent();
                    await SwitchToIframe(token, By.XPath("//iframe[@id='bodyiframe']"));
                    spans = MyDriver.FindElements(By.XPath("//span[@id='field101821span']"));
                    if (t + 1 >= timeout) { throw new System.Exception($"未能找到元素//span[@id='field101821span']"); }
                }

                ItemNO = spans[0].Text;

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return ItemNO;
        }






        private async Task ClearDir(CancellationToken token,string dirPath, int count = 0)
        {
            try
            {                
                DirectoryInfo di = new DirectoryInfo(dirPath);
                if (di.Exists)
                {
                    FileInfo[] files = di.GetFiles();
                    if (files.Length > 0)
                    {
                        foreach (FileInfo file in files)
                        {                            
                            file.Delete();
                             await Task.Delay(500, token);token.ThrowIfCancellationRequested(); 
                        }
                    }
                }

                DirectoryInfo di2 = new DirectoryInfo(dirPath);
                if (di2.Exists&& di.GetFiles().Length>0) 
                { 
                    await ClearDir(token, dirPath,count + 1); 
                    if (count > 5) { throw new Exception($"我需要删除{dirPath}下所有文件，尝试多次失败，请帮我删除。"); } 
                }              
            }

            catch (Exception ex){ throw ex; }
        }

        private async Task SelectDownload(CancellationToken token)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.XPath("//iframe[@id='bodyiframe']"));

                var myas = wait60.Until(d => d.FindElements(By.XPath("//span[@id='selectDownload']//nobr//a")));
                for (int wi = 0; wi <= 30; wi++)
                {
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested();

                    MyDriver.SwitchTo().DefaultContent();
                    await SwitchToIframe(token, By.XPath("//iframe[@id='bodyiframe']"));
                    myas = wait60.Until(d => d.FindElements(By.XPath("//span[@id='selectDownload']//nobr//a")));
                    if (myas.Count() > 0) { break; }
                }
                if (!(myas.Count > 0)) { throw new Exception("未找到//span[@id='selectDownload']//nobr//a"); }


                myas[0].Click();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
              

        private async Task<string> IsFinishDown(CancellationToken token)
        {
            try
            {              
                string result = "";
                string[] files = Directory.GetFiles(OaDownDir, "*.*", SearchOption.AllDirectories);              
                for (int wi = 0; wi <= 30; wi++)
                {                   
                    if (files.Length == 1&& Path.GetExtension(files[0]).ToLower().Contains("xls"))//前提：下载保证清空OaDownDir文件夹里的文件
                    { 
                        result = files[0]; 
                        break; 
                    } 
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    files = Directory.GetFiles(OaDownDir, "*.*", SearchOption.AllDirectories);                   
                } 
                if (result == "") { throw new Exception("超过60秒附件仍未下载完成！"); }
                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
               
        private async Task ClickOaHandle(CancellationToken token,string oaHandleMenu)
        {
            try
            {
                string xpathToFind = "//div[@id='toolbarmenudiv']//input[@value='"+ oaHandleMenu + "']";
                var buttons = wait60.Until(d => d.FindElements(By.XPath(xpathToFind)));
                for (int wi = 0; wi < 30; wi++)
                {
                    if (buttons.Count > 0) 
                    {
                        buttons[0].Click();
                        break;
                    }
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                    buttons = wait60.Until(d => d.FindElements(By.XPath(xpathToFind))); 
                }              
            }
            catch (Exception ex) { throw ex; }
        }

        private async Task SwitchTo_DialogFrame(CancellationToken token,string urlLike)
        {
            try
            {
                bool isfinded=false;
                var iframes = wait60.Until(d => d.FindElements(By.XPath("//iframe")));
                for (int wi = 0; wi <= 30; wi++)
                {
                    if(iframes.Count>0)
                    {
                        foreach (var iframe in iframes)
                        {
                            string url = (string)((IJavaScriptExecutor)MyDriver).ExecuteScript("return arguments[0].contentDocument.URL;", iframe);
                            if (url.ToUpper().Contains(urlLike.ToUpper())) { isfinded = true; MyDriver.SwitchTo().Frame(iframe); return; }
                        }
                    }
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                    iframes = wait60.Until(d => d.FindElements(By.XPath("//iframe")));
                }
                if (!isfinded) { throw new Exception($"未找到对应的iframe，《{urlLike}》"); }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private async Task ClickIcoUser(CancellationToken token)
        {
            try
            {
                string xpathToFind = "//button[@id='field5_browserbtn']";
                IWebElement field5_browserbtn = wait60.Until(d => d.FindElement(By.XPath(xpathToFind)));
                field5_browserbtn.Click();
                 await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
            }
            catch (Exception ex) { throw ex; }
        }
        private async Task EnterUserName(CancellationToken token,DataTable dt_config)
        {
            try
            {
                string TransferUser = await publicClass.GetLibValue(token, dt_config, "TransferUser");
                if (TransferUser == "") { throw new Exception("配置文件中的[转办人]为空"); }
                string xpathToFind = "//input[@id='flowTitle']";
                IWebElement input = wait60.Until(d => d.FindElement(By.XPath(xpathToFind)));
                input.SendKeys(TransferUser);
                 await Task.Delay(100, token);token.ThrowIfCancellationRequested(); 
                //input.SendKeys(OpenQA.Selenium.Keys.Tab);
                 await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
            }
            catch (Exception ex) { throw ex; }  
        }
        private async Task SelectUserName(CancellationToken token, DataTable dt_config)
        {
            try
            {
                bool result = false;
                string TransferUser = await publicClass.GetLibValue(token, dt_config, "TransferUser");
                if (TransferUser == "") { throw new Exception("配置文件中的[转办人]为空"); }
                string xpathToFind = "//div[@id='e8_box_source_quick']//div[@id='src_box_middle']//table[@class='e8_box_source']//tr";
                var trs = wait60.Until(d => d.FindElements(By.XPath(xpathToFind)));
                for (int wi = 0; wi < 30; wi++) 
                { 
                    if (trs.Count > 0) { break; } 
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested();  
                    trs = wait60.Until(d => d.FindElements(By.XPath(xpathToFind)));
                }

                foreach (IWebElement tr in trs)
                {
                    var tds = wait60.Until(d => tr.FindElements(By.XPath(".//td[@id='lastname']")));
                    foreach (var td in tds)
                    {
                        if (td.Text == TransferUser) { td.Click(); result = true; return; }
                    }
                }
                 await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 

                if (!result) { throw new Exception("单击选择姓名时错误"); }
            }
            catch (Exception ex) { throw ex; }
        }

        private async Task SwitchToIFrame_ByID(CancellationToken token, string iframeID, string urlLike)
        {
            try
            {
                bool isfinded = false;
                string xpathToFind = "//iframe[@id='" + iframeID + "']";
                var iframes = wait60.Until(d => d.FindElements(By.XPath(xpathToFind)));
                for (int wi = 0; wi <= 30; wi++)
                {
                    if (iframes.Count > 0)
                    {
                        foreach (var iframe in iframes)
                        {
                            string url = (string)((IJavaScriptExecutor)MyDriver).ExecuteScript("return arguments[0].contentDocument.URL;", iframe);
                            if (url.ToUpper().Contains(urlLike.ToUpper())) { isfinded = true; MyDriver.SwitchTo().Frame(iframe); return; }
                        }
                    }
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                    iframes = wait60.Until(d => d.FindElements(By.XPath(xpathToFind)));
                }

                if (!isfinded) { throw new Exception($"未找到元素：{xpathToFind}"); }
            }
            catch (Exception ex)
            {
                throw ex;
            }           
        }

        private async Task ClickUploadICO(CancellationToken token)
        {
            try
            {
                bool isfinded = false;
                string xpathToFind = "//div[@class='e8fileupload']";
                var divs = wait60.Until(d => d.FindElements(By.XPath(xpathToFind)));
                for (int wi = 0; wi <= 30; wi++)
                {
                    if (divs.Count > 0)
                    {
                        foreach (var div in divs)
                        {
                            if (div.GetDomAttribute("title").Contains("附件上传"))
                            {
                                div.Click();
                                isfinded = true;
                                 await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                                return;
                            }
                        }
                    }
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                    divs = wait60.Until(d => d.FindElements(By.XPath(xpathToFind)));
                }

                if (!isfinded) { throw new Exception($"未找到元素：{xpathToFind}"); }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private async Task ClickSelectFile(CancellationToken token)
        {
            try
            {
                await Task.Delay(3000, token); token.ThrowIfCancellationRequested(); //首次等待3秒，等附件上传完毕
                string xpathToFind = "//div[@id='_filecontentblock']//ul//li[@class='cg_item']";
                var lis = wait60.Until(d => d.FindElements(By.XPath(xpathToFind)));
                for (int wi = 0; wi <= 30; wi++)
                {
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    if (lis.Count > 0)
                    {
                        foreach (var li in lis)
                        {
                            await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                            li.Click();                           
                        }
                        return;
                    }
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                    lis = wait60.Until(d => d.FindElements(By.XPath(xpathToFind)));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
        }

        private async Task SubTransfer(CancellationToken token)
        {
            try 
            {
                string xpathToFind = "//div[@id='tabcontentframe_box']//input[@value='提交']";
                var inputs = wait60.Until(d => d.FindElements(By.XPath(xpathToFind)));
                for (int wi=0;wi<30;wi++)
                {
                    if (inputs.Count > 0) 
                    {
                        inputs[0].Click();
                        break;
                    }
                     await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                    inputs = wait60.Until(d => d.FindElements(By.XPath(xpathToFind)));
                }
                //转办成功提示框
                try
                {
                    IAlert alert = MyDriver.SwitchTo().Alert();
                    string text = alert.Text;
                    alert.Dismiss();
                }
                catch { }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private async Task ResAndSwitchToWindow(CancellationToken token)
        {
            try
            {
                IAlert alert = MyDriver.SwitchTo().Alert();
                for (int wi = 0; wi < 30; wi++)
                {
                    alert = MyDriver.SwitchTo().Alert();
                    string text = alert.Text;
                    alert.Dismiss();
                    if (text.Contains("成功")) { break; }
                     await Task.Delay(100, token);token.ThrowIfCancellationRequested(); 
                }               
            }
            catch { }


            var windowHandles1 = MyDriver.WindowHandles;
            foreach (var wh in windowHandles1)
            {
                MyDriver.SwitchTo().Window(wh);
                if (MyDriver.Url.ToUpper().Contains(@"ManageRequestNoForm.jsp".ToUpper())) //main.jsp
                {
                    MyDriver.SwitchTo().Window(wh).Close();
                    break;
                }
                 await Task.Delay(100, token);token.ThrowIfCancellationRequested(); 
            }

            var windowHandles2 = MyDriver.WindowHandles;
            foreach (var wh in windowHandles2)
            {
                MyDriver.SwitchTo().Window(wh);
                if (MyDriver.Url.ToUpper().Contains(@"main.jsp".ToUpper()))
                {
                    MyDriver.SwitchTo().Window(wh);
                    break;
                }
                 await Task.Delay(100, token);token.ThrowIfCancellationRequested(); 
            }
        }


        //--------------------------------------------------------------------------------------------------
        #endregion


        #region SAP相关   
        //--------------------------------------------------------------------------------------------------
        public async Task EnterTransaction(CancellationToken token,string transactionCode)
        {
            try
            {
                (MySap.Session.FindById("wnd[0]/tbar[0]/okcd") as GuiOkCodeField).Text = transactionCode;//输入事务码
                 await Task.Delay(100, token);token.ThrowIfCancellationRequested(); 
                (MySap.Session.FindById("wnd[0]") as GuiFrameWindow).SendVKey(0);//回车
                 await Task.Delay(500, token);token.ThrowIfCancellationRequested(); 
            }
            catch(Exception ex) { throw ex; }
        }
        private async Task InputExcle(CancellationToken token, string filePath)
        {
            try
            {               
                //(MySap.Session.FindById("wnd[0]/usr/ctxtP_WERKS") as GuiCTextField).Text = site; //输入站点  //这个非必填 取消 2025.01.08
                 await Task.Delay(100, token);token.ThrowIfCancellationRequested(); 
                (MySap.Session.FindById("wnd[0]/usr/radP_R1") as GuiRadioButton).Selected= true;//选中工序
                await Task.Delay(100, token);token.ThrowIfCancellationRequested();       
                
                var waitDia = publicClass.DialogExcleOpen(token);//异步处理>>>损坏，Micrisoft Excel文件格式和扩展名不匹配，是否信任仍要打开它
                                                                 
                (MySap.Session.FindById("wnd[0]/usr/ctxtP_FILE") as GuiCTextField).Text = filePath;//选择路径 
                await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                (MySap.Session.FindById("wnd[0]/tbar[1]/btn[8]") as GuiButton).Press();//执行
                await Task.Delay(500, token);token.ThrowIfCancellationRequested(); 
                Console.WriteLine("GuiButton Press() 结束");
                await waitDia;
                Console.WriteLine(" await waitDia;结束");
                
            }
            catch (Exception ex) { throw ex; }
        }

       

        private async Task<string> InputErrorMsg(CancellationToken token)
        {
            string result = "";

            //不能处理 Excel 文件
            try
            {
                 await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                if (MySap.Session.FindById("wnd[0]/sbar/pane[0]") != null)
                {                   
                    result =  (MySap.Session.FindById("wnd[0]/sbar/pane[0]") as GuiStatusPane).Text.Trim(); //失败消息
                }
                if (result != "") { result="ZPPB002工艺路线导入SAP时发生>>" + result; }
            }
            catch { result = ""; }

            //运行时错误
            try
            {
                await Task.Delay(3000, token); token.ThrowIfCancellationRequested();
                if (MySap.Session.FindById("wnd[0]/mbar/menu[0]") != null)
                {                  
                    string title = (MySap.Session.FindById("wnd[0]/mbar/menu[0]") as GuiMenu).Text;
                    if (title.Contains("运行时错误")) 
                    {
                        result = "ZPPB002工艺路线导入SAP时发生>>运行时错误\r\n";
                        GuiComponentCollection allColl = (MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).Children;
                        int count = 0;
                        foreach (GuiComponent myGuiComponent in allColl)
                        {
                            if (myGuiComponent is GuiLabel thisGui)
                            {
                                count += 1;
                                string str= thisGui.Text;
                                result += str == "" ? ">>" : str;
                                if (count >= 50) { break; }
                            }
                        }
                    }
                }               
            }
            catch { result = ""; }

            return result;
        }

        private async Task<string> GetItemNOText(CancellationToken token, string filePath)
        {
            string text = "";
            try
            {
                DataTable dtFile = await publicClass.ExcelToDataTable(token, filePath);

                if (dtFile == null) { throw new Exception("sap工艺导入后的返回结果为空"); }
                bool MesStyleName = false;
                for (int c = 0; c < dtFile.Columns.Count; c++)
                {
                    if (dtFile.Columns[c].ColumnName == "物料号") { MesStyleName = true; break; }
                    token.ThrowIfCancellationRequested(); await Task.Delay(10, token);
                }
                if (!MesStyleName) { throw new Exception("sap工艺导入后的返回结果不存在[物料号]这样的列"); }


                for (int i = 0; i < dtFile.Rows.Count; i++)
                {
                    string ItemNO = dtFile.Rows[i]["物料号"].ToString().Replace("\r\n","");
                    if (ItemNO != "")
                    {
                        text += ItemNO;
                        if (i + 1 != dtFile.Rows.Count) { text += "\r\n"; }
                    }
                }
            }
            catch (Exception ex) { throw ex; }

            return text;
        }

        private async Task<bool> Check_SE(CancellationToken token,string filePath)
        {
            bool isAll_S = false;
            try
            {
                DataTable dtFile = await publicClass.ExcelToDataTable(token, filePath);
                if (dtFile== null) { throw new Exception("sap工艺导入后的返回结果为空"); }
                bool MesStyleName=false;
                for (int c = 0; c < dtFile.Columns.Count; c++)
                {
                    if (dtFile.Columns[c].ColumnName == "消息类型") { MesStyleName = true;break; }
                    token.ThrowIfCancellationRequested(); await Task.Delay(10, token);
                }
                if (!MesStyleName) { throw new Exception("sap工艺导入后的返回结果不存在[消息类型]这样的列"); }

                isAll_S = true;
                for (int i = 0; i < dtFile.Rows.Count; i++)
                {
                    string mesStyle = dtFile.Rows[i]["消息类型"].ToString();
                    if (mesStyle.Trim() != "" && !mesStyle.ToUpper().Contains("S")) { isAll_S = false; break; }
                    token.ThrowIfCancellationRequested(); await Task.Delay(10, token);
                }
            }
            catch (Exception ex) { throw ex; }

            return isAll_S;
        }

        private async Task<bool> Check_SE0(CancellationToken token)
        {
            bool isAll_S = true;
            try
            {
                 await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                if (MySap.Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell") == null) { throw new Exception($"上一步导入Excel未成功！"); }
                 (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(46);//打印预览46 Ctrl+Shift+F10
                 await Task.Delay(500, token);token.ThrowIfCancellationRequested(); 


                string idstr = "";
                GuiComponentCollection GuiLabels = (MySap.Session.FindById("wnd[0]/usr") as GuiUserArea).Children;
                foreach (GuiComponent GuiLabel in GuiLabels)
                {
                    token.ThrowIfCancellationRequested(); await Task.Delay(0, token);
                    if (GuiLabel is GuiLabel lbl) { if (lbl.Text == "消息类型") { idstr = lbl.Id; break; } }
                }
                if (idstr == "") { throw new Exception($"未能找到<消息类型>的ID"); } //    /app/con[0]/ses[0]/wnd[0]/usr/lbl[1,3]
                string cNum = idstr.Substring(idstr.IndexOf("wnd[0]/usr/lbl[") + "wnd[0]/usr/lbl[".Length);//       1,3]
                cNum = cNum.Substring(0, cNum.IndexOf(","));//1


                bool isThrow = true;
                //后续考虑翻译
                //继续使用上面的GuiLabels               
                foreach (GuiComponent GuiLabel in GuiLabels)
                {
                    token.ThrowIfCancellationRequested(); await Task.Delay(0, token);
                    if (GuiLabel is GuiLabel lbl)
                    {
                        bool b1 = lbl.Id.Contains($"/wnd[0]/usr/lbl[{cNum},");
                        bool b2 = lbl.Text.Trim() != "";
                        bool b3 = !lbl.Text.Trim().Contains("消息类型");
                        if (b1 && b2 && b3) { isThrow = false; }
                        if (b1 && b2 && b3 && lbl.Text.Trim().ToUpper() != "S")
                        {
                            isAll_S = false;
                            break;
                        }
                    }
                }



                if (isThrow) { throw new Exception($"无信息行"); }
            }
            catch (Exception ex) { throw ex; }

            return isAll_S;
        }

     

        private async Task<string> OutputExcleRes(CancellationToken token)
        {
            string filePath = "";
            (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(46);//打印预览Ctrl+Shift+F10 46
             await Task.Delay(500, token);token.ThrowIfCancellationRequested(); 
            (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(43);//43 Ctrl+Shift+F7 电子表格...  
             await Task.Delay(500, token);token.ThrowIfCancellationRequested(); 
            (MySap.Session.FindById("wnd[1]/usr/radRB_OTHERS") as GuiRadioButton).Selected = true;//从所有可用格式中选择
            token.ThrowIfCancellationRequested(); await Task.Delay(10, token);
            (MySap.Session.FindById("wnd[1]/tbar[0]/btn[0]") as GuiButton).Press();//继续(Enter) 【用户登录后需设置为对话模式】
             await Task.Delay(500, token);token.ThrowIfCancellationRequested(); 
            try 
            {
                (MySap.Session.FindById("wnd[1]/usr/ctxtDY_PATH") as GuiCTextField).Text = SapResTempDir;//目录
                token.ThrowIfCancellationRequested(); await Task.Delay(10, token);
                string fileName = (MySap.Session.FindById("wnd[1]/usr/ctxtDY_FILENAME") as GuiCTextField).Text;//文件名称
                if (fileName == "") { fileName = "RPARes.XLSX"; (MySap.Session.FindById("wnd[1]/usr/ctxtDY_FILENAME") as GuiCTextField).Text = fileName; }
                (MySap.Session.FindById("wnd[1]/tbar[0]/btn[11]") as GuiButton).Press();//替换
                 await Task.Delay(500, token);token.ThrowIfCancellationRequested(); 
                filePath = SapResTempDir + fileName;
            }
            catch(Exception ex)
            {                
                ex.HelpLink += "导出电子表格时对话框异常，请确保当前登录用户设置了对话模式。\r\n";
                ex.HelpLink += "F1和F4设置为对话模式：登录后主界面>>>帮助>>>F1帮助>>>对话框模式>>>F4帮助>>>对话(模式)>>>✔确定\r\n";
                throw ex;
            }


            token.ThrowIfCancellationRequested(); await Task.Delay(0, token);
            return filePath;
        }
        private async Task KillExcel(CancellationToken token)
        {           
            try
            {
                 await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                Process[] pros = Process.GetProcesses();
                foreach (Process thisproc in pros)
                {
                    token.ThrowIfCancellationRequested(); await Task.Delay(0, token);
                    //WPS>>表格是：et.exe 文字是：wps.exe
                    if (thisproc.ProcessName.ToUpper() == "EXCEL".ToUpper() || thisproc.ProcessName.ToUpper() == "et".ToUpper())
                    {
                        thisproc.Kill();
                    }
                }
            }
            catch { }  
        }

        public async Task EnterFromClipboard(CancellationToken token)
        {
            try
            {
                (MySap.Session.FindById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH") as GuiButton).Press(); //点击多项选着                
                 await Task.Delay(500, token);token.ThrowIfCancellationRequested(); 
                (MySap.Session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA") as GuiTab).Select();//选择单值
                 await Task.Delay(100, token);token.ThrowIfCancellationRequested(); 
                (MySap.Session.FindById("wnd[1]/tbar[0]/btn[24]") as GuiButton).Press(); //自剪贴板上载
                 await Task.Delay(500, token);token.ThrowIfCancellationRequested(); 
                (MySap.Session.FindById("wnd[1]/tbar[0]/btn[0]") as GuiButton).Press(); //检查条目
                 await Task.Delay(100, token);token.ThrowIfCancellationRequested(); 
                (MySap.Session.FindById("wnd[1]/tbar[0]/btn[12]") as GuiButton).Press();//取消
                 await Task.Delay(100, token);token.ThrowIfCancellationRequested(); 
                (MySap.Session.FindById("wnd[2]/usr/btnSPOP-OPTION1") as GuiButton).Press(); //更改？是               
                 await Task.Delay(100, token);token.ThrowIfCancellationRequested(); 
            }
            catch (Exception ex) { throw ex; }
        }
        public async Task JobAndF8(CancellationToken token)
        {
            try
            {
                (MySap.Session.FindById("wnd[0]/usr/chkP_JOBS") as GuiCheckBox).Selected=true; //勾选JOB            
                 await Task.Delay(500, token);token.ThrowIfCancellationRequested(); 
                (MySap.Session.FindById("wnd[0]/tbar[1]/btn[8]") as GuiButton).Press();     //执行
                 await Task.Delay(100, token);token.ThrowIfCancellationRequested();                 
            }
            catch (Exception ex) { throw ex; }
        }
        public async Task ClickSave(CancellationToken token)
        {
             await Task.Delay(500, token);token.ThrowIfCancellationRequested(); 
            (MySap.Session.FindById("wnd[0]/tbar[0]/btn[11]") as GuiButton).Press();  //保存  
             await Task.Delay(100, token);token.ThrowIfCancellationRequested(); 
        }

        //--------------------------------------------------------------------------------------------------
        #endregion






        public async Task PostWillDo(string requestid, string requestnamespan, string ItemNO, string WechatKey)
        {
            if (WechatKey != "")
            {
                var requestBody = new
                {
                    msgtype = "markdown",
                    markdown = new
                    {
                        content = "SAP工艺路线导入流程有待办事项，即将开跑。请 <font color=\"warning\">" + "徐均任，褚雅庆" + "</font> 注意执行结果。\n" +
                              ">流程标题:<font color=\"comment\">" + requestnamespan + "</font>\n" +
                              ">物料号:<font color=\"comment\">" + ItemNO + "</font>\n" +                            
                              ">唯一标识:<font color=\"comment\">" + requestid + "</font>\n" +                            
                              ">执行结果:<font color=\"comment\">正在执行</font>"
                    }
                };
                string jsonContent = JsonConvert.SerializeObject(requestBody);
                await publicClass.WechatPost(WechatKey, jsonContent);
            }

        }

        public async Task PostResInfo(string requestid,string requestnamespan, string ItemNO, bool AllS, string WechatPostFilePath, string WechatKey,string ErrorMsg)
        {
            if (WechatKey != "")
            {
                string SE = AllS ? "全部为S状态" : "不全为S状态";
                string PUser = "徐均任，褚雅庆";
                string res = ErrorMsg == "" ? "成功" : ErrorMsg;
                if (res == "成功" && AllS) { res = "批准成功"; }
                if (res == "成功" && !AllS) { res = "转办成功"; }

                var requestBody = new
                {
                    msgtype = "markdown",                   
                    markdown = new
                    {
                        content = "SAP工艺路线导入流程执行完毕，请<font color=\"warning\">" + PUser + "</font>注意执行结果。\n" +
                              ">导入状态:<font color=\"comment\">" + SE + "</font>\n" +
                              ">流程标题:<font color=\"comment\">" + requestnamespan + "</font>\n" +
                              ">物料号:<font color=\"comment\">" + ItemNO + "</font>\n" +
                              ">唯一标识:<font color=\"comment\">" + requestid + "</font>\n" +
                              //">附件查阅:<font color=\"comment\">" + WechatPostFilePath + "</font>\n" +
                              ">执行结果:<font color=\"comment\">" + res + "</font>"
                    }
                };
                string jsonContent = JsonConvert.SerializeObject(requestBody);
                await publicClass.WechatPost(WechatKey, jsonContent);
            }
        }



    }
}
