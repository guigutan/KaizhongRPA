using Microsoft.Office.Interop.Outlook;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static Microsoft.IO.RecyclableMemoryStreamManager;
using static Org.BouncyCastle.Math.EC.ECCurve;


namespace KaizhongRPA
{
    internal class Rpa004 : InvokeCenter
    {
        public RpaInfo GetThisRpaInfo()
        {
            RpaInfo rpaInfo = new RpaInfo();
            rpaInfo.RpaClassName = this.GetType().Name;
            rpaInfo.RpaName = "预付待办删除流程";
            rpaInfo.DefaultRunTime1 = "****-**-** **:**:**";
            rpaInfo.DefaultRunTime2 = "****-**-** **:**:**";
            rpaInfo.DefaultStatus = "有效";
            rpaInfo.DefaultPathStype = "相对路径";
            rpaInfo.DefaultConfigPath = @"config\RpaGroup\预付待办删除.xlsx";
            return rpaInfo;
        }
        PublicClass publicClass = new PublicClass();
        public DataTable dt_config;
        public IWebDriver MyDriver;
        public WebDriverWait wait60;
        public async Task RpaMain(CancellationToken token, RpaInfo rpaInfo)
        {
            try
            {
                dt_config = await publicClass.ExcelToDataTable(token, rpaInfo.DefaultConfigPath);
                if (!(dt_config != null && dt_config.Rows.Count > 0)) { return; }


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
                                                                                                // Options.AddUserProfilePreference("download.default_directory", OaDownDir);      //设置下载目录
                Options.AddUserProfilePreference("download.prompt_for_download", false);        //禁用下载提示
                Options.AddUserProfilePreference("download.directory_upgrade", true);           //确保使用指定目录
                Options.AddUserProfilePreference("credentials_enable_service", false);          //禁止弹出密码保存对话框
                Options.AddUserProfilePreference("profile.password_manager_enabled", false);    //禁止弹出密码保存对话框
                #endregion

                using (MyDriver = new ChromeDriver(MyWeb.ChromeDriver, Options))
                {
                    wait60 = new WebDriverWait(MyDriver, TimeSpan.FromSeconds(60));

                    await LoginOA(token, dt_config);    //01-登录OA
                    await Nav_Liucheng(token);          //02-选择顶部菜单《流程》
                    await ClickDaiban(token);          //03-点击待办事宜

                    bool isExist = await ClickLibZF003(token); //是否存在待办
                    await publicClass.DisableScreen(token);
                    if (!isExist) { isExist = await ClickLibZF004(token); }

                    while (isExist)
                    {
                        await publicClass.DisableScreen(token);

                        await ClickItem(token);//点击打开待办流程项
                        if (await LoadedDelBox(token))
                        {
                            await ClickCornerMenu(token);//点击右上的菜单选项
                            await DelRequest(token);//删除流程
                        }
                        await SwitchMainWindow(token); //切换主窗口
                        await RefreshPage(token);//处理登录后可能存在的提示框

                        await Nav_Liucheng(token);          //02-选择顶部菜单《流程》
                        await ClickDaiban(token);          //03-点击待办事宜
                        isExist = await ClickLibZF003(token); //是否存在待办
                        if (!isExist) { isExist = await ClickLibZF004(token); }
                    }

                }
            }
            catch (System.Exception ex) { await publicClass.NoteLog(token, ex, dt_config); }
            finally
            {
                await publicClass.ExitSap(token);
            }

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
                    if (t + 1 == timeout) {  throw new System.Exception("未能等待到打开的新窗口"); }
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
                if (!res) { Console.WriteLine("无法切换新窗口");  throw new System.Exception("无法切换新窗口"); }


            }
            catch (System.Exception ex)
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
                if (isSwitch) {  throw new System.Exception($"未能切换窗口{url}"); }
            }
            catch (System.Exception ex)
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
            catch (System.Exception ex) { throw ex; }
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
                    if (t + 1 == timeout) {  throw new System.Exception($"未能切换iframe>>{by}"); }
                }
            }
            catch (System.Exception ex) { throw ex; }
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
                    if (t + 1 == timeout) {  throw new System.Exception($"未能切换iframe>>{by}"); }
                }
            }
            catch (System.Exception ex) { throw ex; }
        }


        #endregion


        public async Task LoginOA(CancellationToken token, DataTable dt_config)
        {
            try
            {
                string OaLoginUrl = await publicClass.GetLibValue(token, dt_config, "OaLoginUrl");
                string OaUser = await publicClass.GetLibValue(token, dt_config, "OaUser");
                string OaPassWorld = await publicClass.GetLibValue(token, dt_config, "OaPassWorld");
                if (OaLoginUrl == "" || OaUser == "" || OaPassWorld == "") {  throw new System.Exception($"配置表中的OA相关信息为空字符"); }

                //打开网页等待加载并最大化窗口
                MyDriver.Navigate().GoToUrl(OaLoginUrl);
                wait60.Until(d => d.FindElement(By.Id("loginid")).Displayed);
                MyDriver.Manage().Window.Maximize();

                //输入用户密码点击登录
                MyDriver.FindElement(By.Id("loginid")).SendKeys(OaUser);
                MyDriver.FindElement(By.Id("userpassword")).SendKeys(OaPassWorld);
                MyDriver.FindElement(By.Id("login")).Click();

                //等待加载
                wait60.Until((d) => d.FindElement(By.Id("lftop")).Displayed);


                //处理登录后可能存在的提示框
                try
                {
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                    wait60.Until((d) => d.FindElement(By.Id("lftop")).Displayed);
                    MyDriver.FindElement(By.CssSelector(".zd_btn_cancle.btn_submit")).Click();
                }
                catch { }
            }
            catch (System.Exception ex) { throw ex; }

        }
        private async Task Nav_Liucheng(CancellationToken token)
        {
            try
            {
                bool isClick = false;
                wait60.Until(d => d.FindElement(By.XPath("//div[@class='slideItemText' and text()='流程']")).Displayed);
                var divs = MyDriver.FindElements(By.XPath("//div[@class='slideItemText' and text()='流程']"));
                foreach (IWebElement div in divs)
                {
                    var parents = div.FindElements(By.XPath(".."));
                    foreach (IWebElement parent in parents)
                    {
                        if (parent.GetDomAttribute("title") == "我的流程") { parent.Click(); isClick = true; return; }
                    }
                    await Task.Delay(0, token); token.ThrowIfCancellationRequested();
                }
                if (!isClick) {  throw new System.Exception("单击<我的流程>异常"); }

            }
            catch (System.Exception ex)
            {
                throw ex;
            }
        }
        private async Task ClickDaiban(CancellationToken token)
        {
            try
            {

                bool isClick = false;
                wait60.Until(d => d.FindElement(By.Id("drillmenu")).Displayed);
                var drillmenu = MyDriver.FindElement(By.Id("drillmenu"));
                var liCss2s = drillmenu.FindElements(By.CssSelector("li.liCss2"));
                foreach (IWebElement li in liCss2s)
                {
                    var Mya = li.FindElements(By.TagName("a"));
                    foreach (IWebElement a in Mya)
                    {
                        if (a.GetDomAttribute("href").Contains("RequestView.jsp")) { isClick = true; a.Click(); return; }//待办
                        await Task.Delay(0, token); token.ThrowIfCancellationRequested();
                    }
                    await Task.Delay(0, token); token.ThrowIfCancellationRequested();
                }
                if (!isClick) {  throw new System.Exception("单击<待办事宜>异常"); }
            }
            catch (System.Exception ex) { throw ex; }
        }
        private async Task<bool> ClickLibZF003(CancellationToken token,int timeout=60)
        {
            bool res = false;
            try
            {                
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("mainFrame"), "workflow/request/RequestView.jsp");

                // 全部类型：ul  id=ztreeObj
                var uls = MyDriver.FindElements(By.XPath("//ul[@id='ztreeObj']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (uls.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    uls = MyDriver.FindElements(By.XPath("//ul[@id='ztreeObj']"));
                    if (t + 1 >= timeout) {  throw new System.Exception($"未能找到元素//ul[@id='ztreeObj']"); }
                }

                //区块含大分类小列表：li class=level0 e8_z_toplevel
                var lis = uls[0].FindElements(By.XPath(".//li[@class='level0 e8_z_toplevel']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (lis.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    lis = uls[0].FindElements(By.XPath(".//li[@class='level0 e8_z_toplevel']"));
                    if (t + 1 >= timeout) { throw new System.Exception($"未能找到元素.//li[@class='level0 e8_z_toplevel']"); }
                }

                //找出 “支付平台”的li区块
                int myz = -1;
                for (int z = 0; z < lis.Count; z++)
                {
                    var divs = lis[z].FindElements(By.XPath(".//div[@class='e8HoverZtreeDiv']"));
                    for (int t = 0; t < timeout; t++)
                    {
                        if (divs.Count > 0) { break; }
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                        divs = lis[z].FindElements(By.XPath(".//div[@class='e8HoverZtreeDiv']"));
                        if (t + 1 >= timeout) { throw new System.Exception($"未能找到元素.//div[@class='e8HoverZtreeDiv']"); }
                    }

                    var myas = divs[0].FindElements(By.XPath(".//a[@class='level0 e8menu']"));
                    for (int t = 0; t < timeout; t++)
                    {
                        if (divs.Count > 0) { if (myas[0].GetAttribute("title").Contains("支付平台")) { myz = z; break; } }
                        await Task.Delay(500, token); token.ThrowIfCancellationRequested();
                        divs[0].FindElements(By.XPath(".//a[@class='level0 e8menu']"));
                    }                  

                }
                if (myz == -1) { res = false; return res; }


                //“支付平台”的li区块
                IWebElement zhifu_li = lis[myz];
                //“支付平台”的li下的ul区块,只有一个ul
                var zhifu_uls = zhifu_li.FindElements(By.XPath(".//ul"));
                for (int t = 0; t < timeout; t++)
                {
                    if (zhifu_uls.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    zhifu_uls = zhifu_li.FindElements(By.XPath(".//ul"));
                    if (t + 1 >= timeout) { throw new System.Exception(@"未能找到元素。zhifu_uls"); }
                }
                IWebElement zhifu_ul = zhifu_uls[0];


                //支付平台的 分类列表：li   class=level1 e8_z_toplevel
                var ClassItems_lis = zhifu_ul.FindElements(By.XPath(".//li[@class='level1 e8_z_toplevel']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (ClassItems_lis.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    ClassItems_lis = zhifu_ul.FindElements(By.XPath(".//li[@class='level1 e8_z_toplevel']"));
                    if (t + 1 >= timeout) { throw new System.Exception(@"未能找到元素。ClassItems_lis"); }
                }

                //找出“ZF003-预付款申请流程”分类
                int myzf003 = -1;
                for (int z = 0; z < ClassItems_lis.Count; z++)
                {
                    var divs = ClassItems_lis[z].FindElements(By.XPath(".//div[@class='e8HoverZtreeDiv']"));
                    for (int t = 0; t < timeout; t++)
                    {
                        if (divs.Count > 0) { break; }
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                        divs = ClassItems_lis[z].FindElements(By.XPath(".//div[@class='e8HoverZtreeDiv']"));
                        if (t + 1 >= timeout) { throw new System.Exception($"未能找到元素.//div[@class='e8HoverZtreeDiv']"); }
                    }
                    var mya = divs[0].FindElement(By.XPath(".//a[@class='level1 e8menu']"));
                    if (mya.GetAttribute("title").Contains("ZF003-预付")) { 
                        myzf003 = z;
                        res = true;
                        mya.Click();//找到直接点击即可.(点击分类，不会有新窗口)
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                        break; 
                    }
                }
                if (myzf003 == -1) { res = false; return res; }
               
            }
            catch (System.Exception ex) { throw ex; }
            return res;
        }

        private async Task<bool> ClickLibZF004(CancellationToken token, int timeout = 60)
        {
            bool res = false;
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("mainFrame"), "workflow/request/RequestView.jsp");

                // 全部类型：ul  id=ztreeObj
                var uls = MyDriver.FindElements(By.XPath("//ul[@id='ztreeObj']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (uls.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    uls = MyDriver.FindElements(By.XPath("//ul[@id='ztreeObj']"));
                    if (t + 1 >= timeout) { throw new System.Exception($"未能找到元素//ul[@id='ztreeObj']"); }
                }

                //区块含大分类小列表：li class=level0 e8_z_toplevel
                var lis = uls[0].FindElements(By.XPath(".//li[@class='level0 e8_z_toplevel']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (lis.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    lis = uls[0].FindElements(By.XPath(".//li[@class='level0 e8_z_toplevel']"));
                    if (t + 1 >= timeout) { throw new System.Exception($"未能找到元素.//li[@class='level0 e8_z_toplevel']"); }
                }

                //找出 “支付平台”的li区块
                int myz = -1;
                for (int z = 0; z < lis.Count; z++)
                {
                    var divs = lis[z].FindElements(By.XPath(".//div[@class='e8HoverZtreeDiv']"));
                    for (int t = 0; t < timeout; t++)
                    {
                        if (divs.Count > 0) { break; }
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                        divs = lis[z].FindElements(By.XPath(".//div[@class='e8HoverZtreeDiv']"));
                        if (t + 1 >= timeout) { throw new System.Exception($"未能找到元素.//div[@class='e8HoverZtreeDiv']"); }
                    }

                    var myas = divs[0].FindElements(By.XPath(".//a[@class='level0 e8menu']"));
                    for (int t = 0; t < timeout; t++)
                    {
                        if (divs.Count > 0) { if (myas[0].GetAttribute("title").Contains("支付平台")) { myz = z; break; } }
                        await Task.Delay(500, token); token.ThrowIfCancellationRequested();
                        divs[0].FindElements(By.XPath(".//a[@class='level0 e8menu']"));
                    }
                }
                if (myz == -1) { res = false; return res; }

                //“支付平台”的li区块
                IWebElement zhifu_li = lis[myz];
                //“支付平台”的li下的ul区块,只有一个ul
                var zhifu_uls = zhifu_li.FindElements(By.XPath(".//ul"));
                for (int t = 0; t < timeout; t++)
                {
                    if (zhifu_uls.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    zhifu_uls = zhifu_li.FindElements(By.XPath(".//ul"));
                    if (t + 1 >= timeout) { throw new System.Exception(@"未能找到元素。zhifu_uls"); }
                }
                IWebElement zhifu_ul = zhifu_uls[0];


                //支付平台的  分类列表项：li   class=level1 e8_z_toplevel
                var ClassItems_lis = zhifu_ul.FindElements(By.XPath(".//li[@class='level1 e8_z_toplevel']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (ClassItems_lis.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    ClassItems_lis = zhifu_ul.FindElements(By.XPath(".//li[@class='level1 e8_z_toplevel']"));
                    if (t + 1 >= timeout) { throw new System.Exception(@"未能找到元素。ClassItems_lis"); }
                }

                //找出“ZF003-预付款申请流程”分类
                int myzf004 = -1;
                for (int z = 0; z < ClassItems_lis.Count; z++)
                {
                    var divs = ClassItems_lis[z].FindElements(By.XPath(".//div[@class='e8HoverZtreeDiv']"));
                    for (int t = 0; t < timeout; t++)
                    {
                        if (divs.Count > 0) { break; }
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                        divs = ClassItems_lis[z].FindElements(By.XPath(".//div[@class='e8HoverZtreeDiv']"));
                        if (t + 1 >= timeout) { throw new System.Exception($"未能找到元素.//div[@class='e8HoverZtreeDiv']"); }
                    }
                    var mya = divs[0].FindElement(By.XPath(".//a[@class='level1 e8menu']"));
                    if (mya.GetAttribute("title").Contains("ZF004-到付"))
                    {
                        myzf004 = z;
                        res = true;
                        mya.Click();//找到直接点击即可.(点击分类，不会有新窗口)
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                        break;
                    }
                }
                if (myzf004 == -1) { res = false; return res; }

            }
            catch (System.Exception ex) { throw ex; }
            return res;
        }






        private async Task ClickItem(CancellationToken token,int timeout=60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();              
                await SwitchToIframe(token, By.Id("mainFrame"), "RequestView.jsp");
                await SwitchToIframe(token, By.Id("myFrame"), "wfTabFrame.jsp");               
                await SwitchToIframe(token, By.Id("tabcontentframe"), "WFSearchResult.jsp");  

                // var tbodys = MyDriver.FindElements(By.XPath("//form[@id='weaver']//div[@id='_xTable']//table//tbody"));
                var tbodys = MyDriver.FindElements(By.XPath(".//form[@id='weaver']//div[@id='_xTable']//div[@class='table']//table[@class='ListStyle']//tbody"));
                for (int t = 0; t < timeout; t++)
                {
                    if (tbodys.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    MyDriver.SwitchTo().DefaultContent();
                    await SwitchToIframe(token, By.Id("mainFrame"), "RequestView.jsp");
                    await SwitchToIframe(token, By.Id("myFrame"), "wfTabFrame.jsp");
                    await SwitchToIframe(token, By.Id("tabcontentframe"), "WFSearchResult.jsp");
                    tbodys = MyDriver.FindElements(By.XPath("//div[@id='_xTable']//div[@class='table']//table[@class='ListStyle']//tbody"));
                    if (t + 1 >= timeout) {  throw new System.Exception($"未能找到元素tbodys"); }
                }

                List<string> originalWindow = MyDriver.WindowHandles.ToList(); // 获取所有窗口句柄    
                var trs= tbodys[0].FindElements(By.XPath(".//tr"));
                foreach (IWebElement tr in trs)
                {
                    if (tr.GetDomAttribute("class")!=null&&tr.GetDomAttribute("class").Contains("Spacing")) { continue; }

                    var tds = tr.FindElements(By.XPath(".//td"));
                    if (tds.Count >= 2)
                    {
                        IWebElement td2 = tds[1];
                        if (td2.GetDomAttribute("title").Contains("ZF003") || td2.GetDomAttribute("title").Contains("ZF004"))
                        {
                            var mya = td2.FindElements(By.XPath(".//a"));
                            if (mya.Count > 0)
                            {
                                mya[0].Click();
                                await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                                await SwitchToWindow(token, originalWindow, 60);    //切换新窗口
                                await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                            }
                           
                        }
                    }
                    break; //仅处理一条，下次循环需重新点击《ZF003-预付款申请流程》


                }
            }
            catch (System.Exception ex) { throw ex; }
        }

        private async Task<bool> LoadedDelBox(CancellationToken token,int timeout=60)
        {
            bool res = false;
            try 
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/ManageRequestNoFormIframe.jsp");
                await SwitchToIframe(token, By.Id("rightMenuIframe"), "workflow/request/ManageRequestNoFormIframe.jsp");

                var divs = MyDriver.FindElements(By.XPath(".//div[@id='menuTable']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (divs.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    MyDriver.SwitchTo().DefaultContent();
                    await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/ManageRequestNoFormIframe.jsp");
                    await SwitchToIframe(token, By.Id("rightMenuIframe"), "workflow/request/ManageRequestNoFormIframe.jsp");
                    divs = MyDriver.FindElements(By.XPath(".//div[@id='menuTable']"));
                    if (t + 1 >= timeout) {  throw new System.Exception($"未能找到元素menuTable"); }
                }

                var items = divs[0].FindElements(By.XPath(".//div[@class='b-m-item']"));               
                foreach (IWebElement item in items)
                {
                    IWebElement div = item.FindElement(By.XPath(".//div[@class='b-m-ibody']"));
                    IWebElement nobr = div.FindElement(By.XPath(".//nobr"));
                    IWebElement button =nobr.FindElement(By.XPath(".//button"));
                    if (button.GetAttribute("title") != null && button.GetAttribute("title").Contains("删除"))
                    {
                        res = true;
                        break;
                    }
                }

            }
            catch (System.Exception ex) { throw ex; }
            return res;
        }

        private async Task ClickCornerMenu(CancellationToken token, int timeout = 60)
        {
            try 
            {
                MyDriver.SwitchTo().DefaultContent();

               var divs = MyDriver.FindElements(By.XPath("//div[@id='rightBox']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (divs.Count > 0) {  break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    MyDriver.SwitchTo().DefaultContent();                  
                    divs = MyDriver.FindElements(By.XPath("//div[@id='rightBox']"));
                    if (t + 1 >= timeout) {  throw new System.Exception($"未能找到元素 rightBox"); }
                }                            

                IWebElement span = divs[0].FindElement(By.XPath(".//span[@id='rightclickcornerMenu']")); 
                if (span.GetAttribute("title") != null && span.GetAttribute("title").Contains("菜单")) { span.Click(); }

            }
            catch (System.Exception ex) { throw ex; }
        }
        private async Task DelRequest(CancellationToken token, int timeout = 60)
        {
            try {

                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/ManageRequestNoFormIframe.jsp");
                await SwitchToIframe(token, By.Id("rightMenuIframe"), "workflow/request/ManageRequestNoFormIframe.jsp");


                var items = MyDriver.FindElements(By.XPath(".//div[@id='menuTable']//div[@class='b-m-item']"));  
                for (int t = 0; t < timeout; t++)
                {
                    if (items.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    MyDriver.SwitchTo().DefaultContent();
                    await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/ManageRequestNoFormIframe.jsp");
                    await SwitchToIframe(token, By.Id("rightMenuIframe"), "workflow/request/ManageRequestNoFormIframe.jsp");
                    items = MyDriver.FindElements(By.XPath(".//div[@id='menuTable']//div[@class='b-m-item']"));
                    if (t + 1 >= timeout) { throw new System.Exception($"未能找到元素 b-m-item"); }
                }

                foreach (IWebElement item in items)
                {
                    IWebElement div = item.FindElement(By.XPath(".//div[@class='b-m-ibody']"));
                    IWebElement nobr = div.FindElement(By.XPath(".//nobr"));
                    IWebElement button = nobr.FindElement(By.XPath(".//button"));
                    if (button.GetAttribute("title") != null && button.GetAttribute("title").Contains("删除"))
                    {

                        button.Click();
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                        IAlert alert = MyDriver.SwitchTo().Alert();
                        alert.Accept();
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();

                        bool LoadedAlert2 = false;
                        for (int wi = 0; wi <= timeout; wi++)
                        {
                            try
                            {
                                IAlert alert2 = MyDriver.SwitchTo().Alert();
                                alert2.Accept();
                                LoadedAlert2 = true;
                            }
                            catch { }
                            if (LoadedAlert2) { break; }
                            await Task.Delay(500, token); token.ThrowIfCancellationRequested();
                        }

                        break;
                    }
                }

            }
            catch (System.Exception ex) { throw ex; }
        }

        private async Task SwitchMainWindow(CancellationToken token, int timeout = 60)
        {
            try 
            {
                List<string> originalWindow = MyDriver.WindowHandles.ToList(); // 获取所有窗口句柄 

                //切换主窗口：仅有一个时为主窗口，否则直接抛异常下次循环重新执行。
                for (int t = 0; t < timeout; t++)
                {
                    if (originalWindow.Count == 1) { MyDriver.SwitchTo().Window(originalWindow[0]);break; }
                    originalWindow = MyDriver.WindowHandles.ToList();
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    if (t + 1 >= timeout) { throw new System.Exception($"未能找到元素menuTable"); }
                }

            }
            catch (System.Exception ex) { throw ex; }
        }

        private async Task RefreshPage(CancellationToken token, int timeout = 60)
        {
            MyDriver.SwitchTo().DefaultContent();
            MyDriver.Navigate().Refresh();//刷新网页
            await Task.Delay(1000, token); token.ThrowIfCancellationRequested();

            //刷新后 可能出现的提示框           
            for (int t = 0; t <= timeout; t++)
            {
                MyDriver.SwitchTo().DefaultContent();
                var boxs = MyDriver.FindElements(By.CssSelector(".zd_btn_cancle.btn_submit"));
                if (boxs.Count > 0) { boxs[0].Click();break; }
                await Task.Delay(100, token); token.ThrowIfCancellationRequested();
            }


        }





    }
}
