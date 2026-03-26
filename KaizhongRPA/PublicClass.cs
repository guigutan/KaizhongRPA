using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support;
using OpenQA.Selenium.Support.UI;
using Org.BouncyCastle.Crypto.Encodings;
using SAPFEWSELib;
using SapROTWr;
using static Microsoft.IO.RecyclableMemoryStreamManager;

namespace KaizhongRPA
{
    public class PublicClass
    {
        #region WindowsAPI
        [DllImport("user32.dll", EntryPoint = "FindWindow")] private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", EntryPoint = "FindWindowEx")] private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpClassName, string lpWindowName);
        [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
        [DllImport("User32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        #endregion




        /// <summary>
        /// 
        /// <summary>
        /// 替换*通配符,timestr=时间字符串,IsMin=是否是开始时间（是=年份被替换为1900年，否=年份被替换为3000年）
        /// </summary>
        /// <param name="timestr">时间字符串</param>
        /// <param name="IsMin">是否属于开始时间</param>
        /// <returns></returns>
        public string ReplaceTime(string timestr, bool IsMin)
        {

            string result = timestr.Replace("：", ":");
            //年
            if (result.Substring(0, 4) == "****")
            {
                string yyyy = (IsMin ? "1900" : "3000");
                result = result.Substring(0, 0) + yyyy + result.Substring(4);
            }
            //月
            if (result.Substring(5, 2) == "**")
            {
                string MM = (IsMin ? "01" : "12");
                result = result.Substring(0, 5) + MM + result.Substring(7);
            }
            //日
            if (result.Substring(8, 2) == "**")
            {
                string MM = (IsMin ? "01" : "31");
                result = result.Substring(0, 8) + MM + result.Substring(10);
            }
            //时
            if (result.Substring(11, 2) == "**")
            {
                string MM = (IsMin ? "00" : "23");
                result = result.Substring(0, 11) + MM + result.Substring(13);
            }
            //分
            if (result.Substring(14, 2) == "**")
            {
                string MM = (IsMin ? "00" : "59");
                result = result.Substring(0, 14) + MM + result.Substring(16);
            }
            //秒
            if (result.Substring(17, 2) == "**")
            {
                string MM = (IsMin ? "00" : "59");
                result = result.Substring(0, 17) + MM + result.Substring(19);
            }
            return result;

        }

        /// <summary>
        /// Excel转DataTable，首行作为列名
        /// </summary>
        /// <param name="filePath">Excel路径</param>
        /// <param name="sheetIndex">sheet页</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public async Task<DataTable> ExcelToDataTable(CancellationToken token, string filePath, int sheetIndex = 0)
        {
            DataTable dt = new DataTable();
            try
            {
                using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    token.ThrowIfCancellationRequested();  await Task.Delay(0, token);
                    string extension = Path.GetExtension(filePath);
                    if (extension.ToLower() != ".xls" && extension.ToLower() != ".xlsx") { return dt; }
                    IWorkbook workbook;
                    if (extension.ToLower() == ".xls") { workbook = new HSSFWorkbook(file); } else { workbook = new XSSFWorkbook(file); }
                    ISheet sheet = workbook.GetSheetAt(sheetIndex);

                    // 从第一行读取列名
                    IRow Row0 = sheet.GetRow(sheet.FirstRowNum);
                    string[] headerArr = new string[Row0.Cells.Count];
                    for (int cell = 0; cell < Row0.Cells.Count; cell++)
                    {
                        token.ThrowIfCancellationRequested();  await Task.Delay(0, token);
                        string str = Row0.GetCell(cell) != null ? Row0.GetCell(cell).ToString().Trim() : "unknow" + cell;
                        if (headerArr.Contains(str)) { str = str + cell; }//处理重复列名
                        headerArr[cell] = str;
                        dt.Columns.Add(str, typeof(string));
                    }

                    // 遍历所有行，从第二行开始（假设第一行是列名）
                    for (int row = 1; row <= sheet.LastRowNum; row++)
                    {
                        token.ThrowIfCancellationRequested();  await Task.Delay(0, token);
                        IRow rowData = sheet.GetRow(row);
                        DataRow dtRow = dt.NewRow();
                        for (int cell = 0; cell < headerArr.Length; cell++)
                        {
                            token.ThrowIfCancellationRequested();  await Task.Delay(0, token);
                            dtRow[cell] = rowData.GetCell(cell) != null ? rowData.GetCell(cell).ToString().Trim() : "";
                        }
                        dt.Rows.Add(dtRow);
                    }
                }
            }
            catch (Exception ex) { throw ex; }
            return dt;
        }

        /// <summary>
        /// 查找dt指定LibName项的LibValue值，找不到返回空字符串。不向父级抛异常。（注意即使找到LibName项而其LibValue本身就为空字符串的情形）
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="LibName"></param>
        /// <returns></returns>
        public async Task<string> GetLibValue(CancellationToken token, DataTable dt, string LibName)
        {
            string result = "";
            if (!(dt != null && dt.Rows.Count > 0)) { return result; }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                await Task.Delay(0, token); token.ThrowIfCancellationRequested(); 
                try
                {
                    //最好是写成第一列：dt.Rows[i][0]，它的项的值是第二列：dt.Rows[i][1]
                    //if (dt.Rows[i]["LibName"].ToString() == LibName) { result = dt.Rows[i]["LibValue"].ToString(); break; }                   
                    if (dt.Rows[i][0].ToString() == LibName) { result = dt.Rows[i][1].ToString(); break; }
                }
                catch { }
            }

            return result;
        }

        public async Task NoteLog(CancellationToken token, Exception ex, DataTable dt_config, string[] filePaths = null)
        {
            try
            {
                string MailTo = await GetLibValue(token, dt_config, "MailTo");
                if (MailTo == null || MailTo == "") { return; }

                //1、依据版本引用：Microsoft Outlook 15.0 Object Library
                //2、设置：Outlook2013-- - 邮件管理员方式运行-- - 文件-- - 选项-- - 信任中心-- - 信任中心设置-- - 编程访问-- - 从不向我发出可疑活动警告
                Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.MailItem mailItem = (Microsoft.Office.Interop.Outlook.MailItem)outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                // 设置邮件属性
                mailItem.Subject = $"来自KaizhongRPA的异常通知";
                mailItem.To = MailTo;
                mailItem.Body = $"{ex.Message}\r\n{ex.StackTrace}\r\n{ex.HelpLink}\r\n{ex.Source}";


                // 添加附件
                if (filePaths != null && filePaths.Length > 0)
                {
                    foreach (string filePath in filePaths)
                    {
                        mailItem.Attachments.Add(filePath);
                    }
                }

                // 发送邮件
                mailItem.Send();
                await Task.Delay(1000, token); token.ThrowIfCancellationRequested(); 
            }
            catch { }
        }


        #region SAP登录相关(GotoSapHome)


        public async Task<bool> GotoSapHome(CancellationToken token, DataTable dtConf, int failCount = 0, int backspaceCount = 5)
        {
            bool result = false;
            try
            {
                //1、是否存在未失效的可用的Session
                //>>>>>>存在：退回首页。
                //>>>>>>不存在：关闭程序/结束进程（如果有），启动程序重新登录。
                if (MySap.Session != null && MySap.Session.FindById("wnd[0]/tbar[0]/okcd") != null)
                {
                    result = await Backspace(token, backspaceCount);
                }
                else
                {
                    await ExitSap(token);//关闭后需等待几秒，否则SapStartLogin检测到存在程序不会启动。
                    await Task.Delay(3000, token); token.ThrowIfCancellationRequested(); 

                    bool IsSetSapLocal = await SapLocalConfig(token);//①启用脚本；②安全模式；
                    if (!IsSetSapLocal) { MessageBox.Show($"以注册表的方式设置启动脚本和安全模式时错误，请手动检查权限问题。\r\n或者手动设置后，新建空白内容文件：{MyPath.SapLocalConfig}"); return IsSetSapLocal; }

                    result = await SapStartLogin(token, dtConf);

                    if (result)
                    {
                        //已优化：[用户]
                        result = await IsDialogBased(token, dtConf); //③对话模式；处理对话框模式                      
                    }
                }
            }
            catch
            {
                //人为注销SAP登录，仍有SAP Logon 740的Process，但已经不是一个MySap.Session 
                failCount += 1; result = false;
                MySap.Session = null; MySap.Connection = null;
                if (failCount == 1) { result = await GotoSapHome(token, dtConf, failCount); }
            }

            return result;
        }


        public async Task<bool> SapLocalConfig(CancellationToken token)
        {
            bool res = true;
            try
            {
                if (!File.Exists(MyPath.SapLocalConfig))
                {
                    res = false;
                    string UserSaplogonini = MyPath.AppData + @"Roaming\SAP\Common\saplogon.ini";
                    string DefaultSaplogonini = MyPath.App + @"config\Common\saplogon.ini";
                    if (!Directory.Exists(Path.GetDirectoryName(UserSaplogonini))) { Directory.CreateDirectory(Path.GetDirectoryName(UserSaplogonini)); }
                    File.Copy(DefaultSaplogonini, UserSaplogonini, true);

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested(); 

                    //设置注册表：计算机\HKEY_CURRENT_USER\SOFTWARE\SAP\SAPGUI Front\SAP Frontend Server\Security
                    string keyPath = @"SOFTWARE\SAP\SAPGUI Front\SAP Frontend Server\Security";
                    string[] valueNames = { "DefaultAction", "SecurityLevel", "UserScripting", "WarnOnAttach", "WarnOnConnection" };
                    int[] valueData = { 0, 0, 1, 0, 0 };
                    using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath, true))
                    {
                        if (key == null)
                        {
                            // 如果键不存在，则创建它
                            using (RegistryKey newKey = Registry.CurrentUser.CreateSubKey(keyPath))
                            {
                                for (int i = 0; i < valueNames.Length; i++)
                                {
                                    await Task.Delay(10, token); token.ThrowIfCancellationRequested(); 
                                    newKey.SetValue(valueNames[i], valueData[i], RegistryValueKind.DWord);
                                    res = true;
                                    if (!File.Exists(MyPath.SapLocalConfig)) { File.Create(MyPath.SapLocalConfig); }
                                }
                            }
                        }
                        else
                        {
                            // 如果键存在，则直接设置值
                            for (int i = 0; i < valueNames.Length; i++)
                            {
                                await Task.Delay(10, token); token.ThrowIfCancellationRequested(); 
                                key.SetValue(valueNames[i], valueData[i], RegistryValueKind.DWord);
                                res = true;
                                if (!File.Exists(MyPath.SapLocalConfig)) { File.Create(MyPath.SapLocalConfig); }
                            }
                        }
                    }

                }
            }

            catch { res = false; if (File.Exists(MyPath.SapLocalConfig)) { File.Delete(MyPath.SapLocalConfig); } }

            return res;
        }


        public async Task<bool> IsDialogBased(CancellationToken token, DataTable dtConf)
        {
            bool GotoSapHome_Res = false;
            try
            {
                if (!Directory.Exists(Path.GetDirectoryName(MyPath.DialogBox))) { Directory.CreateDirectory(Path.GetDirectoryName(MyPath.DialogBox)); }
                if (!File.Exists(MyPath.DialogBox)) { File.Create(MyPath.DialogBox).Close(); }
                string SapUserName = await GetLibValue(token, dtConf, "SapUserName");
                string[] lines = File.ReadAllLines(MyPath.DialogBox);
                foreach (string line in lines)
                {
                    //判断该用户是否为对话框模式
                    if (line.ToUpper().Contains(SapUserName.ToUpper()))
                    {
                        GotoSapHome_Res = true;
                        break;
                    }
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested(); 
                }


                if (!GotoSapHome_Res)
                {
                    //设置对话框模式
                    (MySap.Session.FindById("wnd[0]/mbar/menu[5]/menu[8]") as GuiMenu).Select(); //帮助-设置                    
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested(); 
                    (MySap.Session.FindById("wnd[1]/usr/tabsUSR_VALS/tabpF1HI") as GuiTab).Select(); //选择F1帮助选项卡
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested(); 
                    (MySap.Session.FindById("wnd[1]/usr/tabsUSR_VALS/tabpF1HI/ssubSUBSCREEN1:SAPLSR13:0103/radSF1SETTING-SETTING2") as GuiRadioButton).Select(); //选中对话框模式
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested(); 
                    (MySap.Session.FindById("wnd[1]/usr/tabsUSR_VALS/tabpF4HI") as GuiTab).Select(); ////选择F4帮助选项卡
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested(); 
                    (MySap.Session.FindById("wnd[1]/usr/tabsUSR_VALS/tabpF4HI/ssubSUBSCREEN1:SAPLSDH4:0703/radF4SETTING-NOACTIVEX") as GuiRadioButton).Select(); //选中对话（模式）
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested(); 
                    (MySap.Session.FindById("wnd[1]/tbar[0]/btn[0]") as GuiButton).Press();//点击应用
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested(); 
                    //将用户名追加写入文件
                    using (StreamWriter writer = File.AppendText(MyPath.DialogBox))
                    {
                        writer.WriteLine(SapUserName);
                    }
                    //重启生效
                    await ExitSap(token);
                    MySap.Session = null; MySap.Connection = null;
                    GotoSapHome_Res = await GotoSapHome(token, dtConf);
                }


            }
            catch { GotoSapHome_Res = true;/*异常暂时返回真*/ }
            return GotoSapHome_Res;
        }


        public async Task<bool> Backspace(CancellationToken token, int count)
        {
            bool result = false;
            for (int i = 0; i < count; i++)
            {
                try
                {
                    if (MySap.Session.FindById("wnd[0]/tbar[0]/okcd") != null)
                    {
                        (MySap.Session.FindById("wnd[0]/tbar[0]/okcd") as GuiOkCodeField).Text = "/n";
                        await Task.Delay(50, token); token.ThrowIfCancellationRequested(); 
                        (MySap.Session.FindById("wnd[0]") as GuiFrameWindow).SendVKey(0);
                        await Task.Delay(50, token); token.ThrowIfCancellationRequested(); 
                        result = true;
                    }
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested(); 
                }
                catch { }
            }
            return result;
        }


        /// 关闭程序/结束进程（如果有）
        /// </summary>
        public async Task ExitSap(CancellationToken token)
        {
            try
            {
                if (MySap.Session != null && MySap.Session.FindById("wnd[0]/tbar[0]/okcd") != null)
                {
                    await Backspace(token, 5);
                    (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(15);//15 Shift+F3
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested(); 
                    (MySap.Session.FindById("wnd[1]/usr/btnSPOP-OPTION1") as GuiButton).Press();
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested(); 
                    // Shift + F3注销关闭后，进入下面的：正常关闭[SAP Logon 740]
                }
            }
            catch { MySap.Session = null; MySap.Connection = null; }


            //正常关闭(仅第一个窗口[SAP Logon 740]，第二个主界面窗口无法定位,且关闭时弹注销询问框，待解决)
            IntPtr hWnd = FindWindow(null, "SAP Logon 740");
            if (hWnd != IntPtr.Zero) { PostMessage(hWnd, 0x0010, IntPtr.Zero, IntPtr.Zero); } //请求关闭窗口WM_CLOSE对应的ID=0x0010
            await Task.Delay(3000, token); token.ThrowIfCancellationRequested(); // //给足够时间右下角图标释放

            //直接杀死（会堆积任务栏右下图标，所以每次使用都应在MySap.Session失效前进行Shift+F3注销关闭）
            if (FindWindow(null, "SAP Logon 740") != IntPtr.Zero)
            {
                Process[] processs = Process.GetProcesses();
                for (int i = 0; i < processs.Length; i++)
                {
                    token.ThrowIfCancellationRequested();  await Task.Delay(0, token);
                    if (processs[i].ProcessName.ToString().ToLower().Contains("saplogon"))
                    {
                        processs[i].Kill();
                    }
                }
            }
            MySap.Session = null; MySap.Connection = null;
        }

        private async Task<bool> SapStartLogin(CancellationToken token, DataTable dtConf)
        {
            bool result = false;
            try
            {
                string SapLogon = await GetLibValue(token, dtConf, "SapLogon");
                string SapHost = await GetLibValue(token, dtConf, "SapHost");
                string SapPort = await GetLibValue(token, dtConf, "SapPort");
                string SapGroup = await GetLibValue(token, dtConf, "SapGroup");
                string SapClient = await GetLibValue(token, dtConf, "SapClient");
                string SapUserName = await GetLibValue(token, dtConf, "SapUserName");
                string SapPassWorld = await GetLibValue(token, dtConf, "SapPassWorld");
                string SapLanguage = await GetLibValue(token, dtConf, "SapLanguage");
                if (SapLogon == "" || SapHost == "" || SapPort == "" || SapGroup == "" || SapClient == "" || SapUserName == "" || SapPassWorld == "" || SapLanguage == "") { throw new Exception($"配置表中的SAP相关信息部分或全部为空字符"); }

                string ConnectionString = SapGroup.Contains("无") ? $"/H/{SapHost}/S/{SapPort}" : $"/M/{SapHost}/S/{SapPort}/G/{SapGroup}";

                var Application = await GetSAPGuiApp(token, SapLogon, 10);

                Application.OpenConnectionByConnectionString(ConnectionString);

                var index = Application.Connections.Count - 1;
                MySap.Connection = Application.Children.ElementAt(index) as GuiConnection;//设置Connection
                index = MySap.Connection.Sessions.Count - 1;
                if (MySap.Connection.Sessions.Count == 0) { throw new Exception("新会话没有发现，SAP客户端需开启脚本，请检查。"); }
                MySap.Session = MySap.Connection.Children.Item(index) as GuiSession;//设置Session

                (MySap.Session.FindById("wnd[0]/usr/txtRSYST-BNAME") as GuiTextField).Text = SapUserName;
                await Task.Delay(200, token); token.ThrowIfCancellationRequested(); 
                (MySap.Session.FindById("wnd[0]/usr/pwdRSYST-BCODE") as GuiTextField).Text = SapPassWorld;
                await Task.Delay(200, token); token.ThrowIfCancellationRequested(); 
                (MySap.Session.FindById("wnd[0]/usr/txtRSYST-MANDT") as GuiTextField).Text = SapClient;
                await Task.Delay(200, token); token.ThrowIfCancellationRequested(); 
                (MySap.Session.FindById("wnd[0]/usr/txtRSYST-LANGU") as GuiTextField).Text = SapLanguage;
                await Task.Delay(200, token); token.ThrowIfCancellationRequested(); 
                (MySap.Session.FindById("wnd[0]") as GuiFrameWindow).SendVKey(0);
                await Task.Delay(200, token); token.ThrowIfCancellationRequested(); 

                result = true;
            }
            catch { }
            return result;
        }

        private async Task<GuiApplication> GetSAPGuiApp(CancellationToken token, string SapLogon, int secondsOfTimeOut)
        {
            CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
            object SapGuiRot = sapROTWrapper.GetROTEntry("SAPGUI");
            if (secondsOfTimeOut < 0) { throw new TimeoutException($"secondsOfTimeOut时间内仍旧无法获取SAPGUI Application"); }
            else
            {
                if (SapGuiRot == null)
                {
                    Process[] resProcess = Process.GetProcessesByName("saplogon");
                    foreach (Process proc in resProcess)
                    {
                        proc.Kill();
                        token.ThrowIfCancellationRequested();  //await Task.Delay(10, token);
                    }
                    Process.Start(SapLogon);//启动后等3秒左右，否则SapGuiRot仍为NULL
                    token.ThrowIfCancellationRequested();  await Task.Delay(5000, token);
                    return await GetSAPGuiApp(token, SapLogon, secondsOfTimeOut - 1);
                }
                else
                {
                    // SapLogon32770(token); //不等待
                    object engine = SapGuiRot.GetType().InvokeMember("GetSCriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, SapGuiRot, null);

                    if (engine == null)
                    {
                        throw new NullReferenceException("SAPGUI Application没有发现。");
                    }
                    return engine as GuiApplication;
                }
            }

        }

        //sap logon 某脚本正试图访问SAP GUI
        //标题：SAP Logon 类：#32770 (对话框)
        //标题：确定(&O)  类：Button
        private async Task SapLogon32770(CancellationToken token)
        {
            //sap logon 某脚本正试图访问SAP GUI
            try
            {
                for (int i = 0; i < 5; i++)
                {
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested(); 
                    IntPtr logonGui = FindWindow(null, "SAP Logon");
                    if (logonGui != IntPtr.Zero)
                    {
                        IntPtr Button = FindWindowEx(logonGui, IntPtr.Zero, "Button", "确定(&O)");
                        PostMessage(Button, 0x00F5, IntPtr.Zero, IntPtr.Zero);//请求按钮点击BM_CLICK对于的ID=0x00F5                 
                    }
                }
            }
            catch { }
        }


        #endregion



        public async Task<string> GetDtValue(CancellationToken token, DataTable dt, string ColumnName, bool IsExactColumnName)
        {
            string result = "";
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                if ((IsExactColumnName && dt.Columns[c].ColumnName.Trim() == ColumnName) || (!IsExactColumnName && dt.Columns[c].ColumnName.Contains(ColumnName)))
                {
                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        token.ThrowIfCancellationRequested();  await Task.Delay(10, token);
                        if (dt.Rows[r][c] != null && dt.Rows[r][c].ToString().Trim() != "")
                        {
                            result = dt.Rows[r][c].ToString().Trim();
                            return result;
                        }
                    }
                }
                token.ThrowIfCancellationRequested();  await Task.Delay(10, token);
            }
            return result;
        }
        public async Task<List<string>> GetDtValue_List(CancellationToken token, DataTable dt, string ColumnName, bool IsExactColumnName)
        {
            List<string> result = new List<string>();
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                if ((IsExactColumnName && dt.Columns[c].ColumnName.Trim() == ColumnName) || (!IsExactColumnName && dt.Columns[c].ColumnName.Contains(ColumnName)))
                {
                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        token.ThrowIfCancellationRequested();  await Task.Delay(10, token);
                        if (dt.Rows[r][c] != null && dt.Rows[r][c].ToString().Trim() != "")
                        {
                            result.Add(dt.Rows[r][c].ToString().Trim());
                        }
                    }
                    break;
                }
                token.ThrowIfCancellationRequested();  await Task.Delay(10, token);
            }
            return result;
        }

        public async Task<bool> WinApi_ChooseFile32770(CancellationToken token, string SapResFile)
        {
            bool result = false;
            try
            {
                IntPtr hWnd = FindWindow("#32770", null);
                if (hWnd != IntPtr.Zero)
                {
                    result = true;
                    IntPtr ComboBoxEx32 = FindWindowEx(hWnd, IntPtr.Zero, "ComboBoxEx32", null);
                    if (ComboBoxEx32 != IntPtr.Zero)
                    {
                        IntPtr ComboBox = FindWindowEx(ComboBoxEx32, IntPtr.Zero, "ComboBox", null);
                        if (ComboBox != IntPtr.Zero)
                        {
                            IntPtr Edit = FindWindowEx(ComboBox, IntPtr.Zero, "Edit", null);
                            if (Edit != IntPtr.Zero)
                            {
                                await SetClipboardTextAsync(SapResFile);//设置剪贴板
                                string text = await GetClipboardTextAsync();//获取剪贴板
                                if (SapResFile == text)
                                {
                                    SetForegroundWindow(hWnd);
                                    PostMessage(Edit, 0x0007, IntPtr.Zero, IntPtr.Zero);//获得焦点后WM_SETFOCUS对于的ID=0x0007 
                                    SendKeys.SendWait("^v");
                                    await Task.Delay(500, token); token.ThrowIfCancellationRequested();
                                    //回车
                                    SendKeys.SendWait("{ENTER}");
                                    await Task.Delay(500, token); token.ThrowIfCancellationRequested();
                                }
                                else { throw new Exception("选择附件：将SapResFile设置到剪贴板时错误"); }
                            }
                        }
                    }

                    IntPtr open = FindWindowEx(hWnd, IntPtr.Zero, "Button", "打开(&O)");
                    PostMessage(open, 0x00F5, IntPtr.Zero, IntPtr.Zero);//请求按钮点击BM_CLICK对于的ID=0x00F5 
                    result = false;
                    token.ThrowIfCancellationRequested();  await Task.Delay(10000, token);  //待优化，直接等待10s

                    //#32770选完，会一直存在
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }


            return result;

        }
        public async Task DialogExcleOpen(CancellationToken token)
        {
            // Microsoft Excel  》》#32770 (对话框)
            //是(&Y) 》》Button
            //是否能抓到句柄 待确认
            try
            {
                Console.WriteLine("已经进入DialogExcleOpen");
                IntPtr hWnd = FindWindow("#32770", "Microsoft Excel");
                for (int wi = 0; wi < 30; wi++)
                {
                    Console.WriteLine($"wi={wi},hWnd={hWnd}");
                    if ((hWnd != IntPtr.Zero)) { break; }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested(); 
                    hWnd = FindWindow("#32770", "Microsoft Excel");
                }

                Console.WriteLine($"hWnd={hWnd}");

                if (hWnd != IntPtr.Zero)
                {
                    IntPtr ButtonY = FindWindowEx(hWnd, IntPtr.Zero, "Button", "是(&Y)");
                    if (ButtonY != IntPtr.Zero)
                    {
                        PostMessage(ButtonY, 0x00F5, IntPtr.Zero, IntPtr.Zero);             //请求按钮点击BM_CLICK对于的ID=0x00F5 
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested(); 
                    }
                    //SetForegroundWindow(hWnd);
                    //SendKeys.SendWait("%y");                   
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



        public async Task SetClipboard(CancellationToken token, string txtPath, string text)
        {
            try
            {

                using (FileStream fs = new FileStream(txtPath, FileMode.Create))
                {
                    using (StreamWriter writer = new StreamWriter(fs))
                    {
                        writer.Write(text);
                    }

                }
                await Task.Delay(1000, token); token.ThrowIfCancellationRequested(); 

                Process process = new Process();
                process.StartInfo.FileName = txtPath;
                process.Start();
                await Task.Delay(500, token); token.ThrowIfCancellationRequested(); 
                IntPtr notepad = FindWindow(null, "TempNote.txt - 记事本");
                SetForegroundWindow(notepad);
                SendKeys.SendWait("^a");
                token.ThrowIfCancellationRequested();  await Task.Delay(10, token);
                SendKeys.SendWait("^c"); //发送Ctrl+C ，必须是小写！
                await Task.Delay(1000, token); token.ThrowIfCancellationRequested(); 

                try
                {
                    process.CloseMainWindow(); //正常关闭
                    if (!process.WaitForExit(1000))
                    {
                        process.Kill(); //强制关闭
                    }
                }
                catch { }
                //PostMessage(notepad, 0x0010, IntPtr.Zero, IntPtr.Zero); // //关闭txtPath文档,请求关闭窗口WM_CLOSE对应的ID=0x0010           
            }
            catch (Exception ex) { throw ex; }


        }

        public async Task CloseExcle(CancellationToken token)
        {
            try
            {
                await Task.Delay(3000, token); token.ThrowIfCancellationRequested(); 
                Process[] excelProcesses = Process.GetProcessesByName("Excel");

                foreach (Process process in excelProcesses)
                {
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested(); 
                    process.Kill();
                }
            }
            catch { }

        }

        public static async Task SetClipboardTextAsync(string text)
        {
            await Task.Run(() =>
            {
                // 确保在UI线程上执行剪贴板操作
                if (Application.OpenForms.Count > 0) // 检查是否有打开的Form
                {
                    Application.OpenForms[0].Invoke(new Action(() =>
                    {
                        Clipboard.SetText(text);
                    }));
                }
                else
                {
                    // 如果没有打开的Form，假设我们是在非UI线程中运行，需要创建一个临时Form
                    using (var tempForm = new Form())
                    {
                        tempForm.CreateControl(); // 确保Form已创建
                        tempForm.Invoke(new Action(() =>
                        {
                            Clipboard.SetText(text);
                        }));
                    }
                }
            });
        }
        public static async Task<string> GetClipboardTextAsync()
        {
            string res = "";
            await Task.Run(() =>
            {
                // 确保在UI线程上执行剪贴板操作
                if (Application.OpenForms.Count > 0) // 检查是否有打开的Form
                {
                    Application.OpenForms[0].Invoke(new Action(() =>
                    {
                        res = Clipboard.GetText();
                    }));
                }
                else
                {
                    // 如果没有打开的Form，假设我们是在非UI线程中运行，需要创建一个临时Form
                    using (var tempForm = new Form())
                    {
                        tempForm.CreateControl(); // 确保Form已创建
                        tempForm.Invoke(new Action(() =>
                        {
                            res = Clipboard.GetText();
                        }));
                    }
                }
            });

            return res;
        }





        /// <summary>
        /// 获取数据量连接字符串。向父级跑异常。指定列名[RpaDataSource,RpaDataUser,RpaDataPwd,RpaDatabase]
        /// </summary>
        /// <param name="token"></param>
        /// <param name="dt_config"></param>
        /// <returns></returns>
        public async Task<string> GetConnstr(CancellationToken token, DataTable dt_config)
        {
            string connstr = "";
            try
            {
                string RpaDataSource = await GetLibValue(token, dt_config, "RpaDataSource");
                string RpaDataUser = await GetLibValue(token, dt_config, "RpaDataUser");
                string RpaDataPwd = await GetLibValue(token, dt_config, "RpaDataPwd");
                string RpaDatabase = await GetLibValue(token, dt_config, "RpaDatabase");
                connstr = $"data source={RpaDataSource};user={RpaDataUser};pwd={RpaDataPwd};database={RpaDatabase};Connect Timeout=5";
            }
            catch (Exception ex) { throw ex; }
            return connstr;
        }



        public async Task ClearDir(CancellationToken token, string dirPath, int count = 0)
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
                            await Task.Delay(500, token); token.ThrowIfCancellationRequested(); 
                        }
                    }
                }

                DirectoryInfo di2 = new DirectoryInfo(dirPath);
                if (di2.Exists && di.GetFiles().Length > 0)
                {
                    await ClearDir(token, dirPath, count + 1);
                    if (count > 5) { throw new Exception($"我需要删除{dirPath}下所有文件，尝试多次失败，请帮我删除。"); }
                }
            }

            catch (Exception ex) { throw ex; }
        }

        /// <summary>
        /// Ctrl+F9选择布局
        /// </summary>
        /// <param name="token"></param>
        /// <param name="NameStr">布局名称</param>
        /// <param name="isExactNameStr">布局名称是否完全等于</param>
        /// <param name="UpCount"></param>
        /// <returns></returns>
        public async Task<bool> ChooseLayout(CancellationToken token, string NameStr, bool isExactNameStr = false, int UpCount = 10)
        {
            bool res = false;
            try
            {
                (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(33);//33 Ctrl+F9 选择布局
                await Task.Delay(1000, token); token.ThrowIfCancellationRequested(); 
                IntPtr hwnd = FindWindow(null, "选择布局");
                SetForegroundWindow(hwnd); SendKeys.SendWait("{TAB}");
                await Task.Delay(100, token); token.ThrowIfCancellationRequested(); 
                SetForegroundWindow(hwnd); SendKeys.SendWait("{LEFT}");
                await Task.Delay(100, token); token.ThrowIfCancellationRequested(); 

                SetForegroundWindow(hwnd); SendKeys.SendWait("{PGUP}");//PageUp 置于首行
                SetForegroundWindow(hwnd); SendKeys.SendWait("{PGUP}");//PageUp 置于首行 再一次容错 

                for (int i = 0; i < UpCount; i++)
                {
                    SetForegroundWindow(hwnd); SendKeys.SendWait("^c"); //发送Ctrl+C ，必须是小写！
                    await Task.Delay(200, token); token.ThrowIfCancellationRequested(); 
                    string text = await GetClipboardTextAsync();
                    bool isit = ((isExactNameStr && text == NameStr) || (!isExactNameStr && text.Contains(NameStr)));
                    if (isit) { SetForegroundWindow(hwnd); SendKeys.SendWait("{ENTER}"); res = true; break; }
                    SetForegroundWindow(hwnd); SendKeys.SendWait("{DOWN}");
                }
                await Task.Delay(1000, token); token.ThrowIfCancellationRequested(); 
            }
            catch (Exception ex) { throw ex; }

            return res;
        }


        public async Task<string> GetClass2(CancellationToken token, string PGroup, DataTable dt_class2)
        {
            string result = "";
            if (dt_class2 != null && dt_class2.Rows.Count > 0)
            {
                for (int i = 0; i < dt_class2.Rows.Count; i++)
                {
                    if (PGroup.Trim() == dt_class2.Rows[i]["采购组"].ToString().Trim())
                    {
                        result = dt_class2.Rows[i]["二级分类"].ToString().Trim();
                        break;
                    }
                }
                await Task.Delay(0, token); token.ThrowIfCancellationRequested(); 
            }
            return result;
        }


        public async Task<List<IWebElement>> FindByDriver22(CancellationToken token, IWebDriver myDriver, By by, int timeout = 60, int ckType = 1)
        {
            try
            {
                var wait = new WebDriverWait(myDriver, TimeSpan.FromSeconds(timeout));// 使用 WebDriverWait 进行高效等待                
                wait.PollingInterval = TimeSpan.FromMilliseconds(500);// 设置检查间隔为500毫秒                
                List<IWebElement> resElements = new List<IWebElement>();// 收集符合条件的所有元素
                await Task.Delay(500, token); token.ThrowIfCancellationRequested(); 
                wait.Until(dd =>
                {
                    token.ThrowIfCancellationRequested();
                    var elements = myDriver.FindElements(by);
                    foreach (var ele in elements)
                    {
                        if (MeetsCondition(ele, ckType))
                        {
                            resElements.Add(ele);
                        }
                    }
                    return resElements.Count > 0; // 当至少找到一个元素时，停止等待
                });

                return resElements;
            }
            catch (WebDriverTimeoutException)
            {
                throw new Exception($"在{timeout}秒内未找到符合条件为 '{by.Mechanism}' 的元素。");
            }
            catch (OperationCanceledException)
            {
                throw new OperationCanceledException("操作已被取消。");
            }
            catch (Exception ex)
            {
                throw new Exception($"查找符合条件为 '{by.Mechanism}' 的元素时发生错误：{ex.Message}", ex);
            }

        }

        public async Task<List<IWebElement>> FindByElement22(CancellationToken token, IWebDriver myDriver, IWebElement myParentbElement, By by, int timeout = 60, int ckType = 1)
        {
            try
            {
                var wait = new WebDriverWait(myDriver, TimeSpan.FromSeconds(timeout));// 使用 WebDriverWait 进行高效等待                
                wait.PollingInterval = TimeSpan.FromMilliseconds(500);// 设置检查间隔为500毫秒                
                List<IWebElement> resElements = new List<IWebElement>();// 收集符合条件的所有元素
                await Task.Delay(500, token); token.ThrowIfCancellationRequested(); 
                wait.Until(dd =>
                {
                    token.ThrowIfCancellationRequested();
                    var elements = myParentbElement.FindElements(by);
                    foreach (var ele in elements)
                    {
                        if (MeetsCondition(ele, ckType))
                        {
                            resElements.Add(ele);
                        }
                    }
                    return resElements.Count > 0; // 当至少找到一个元素时，停止等待
                });
                return resElements;
            }
            catch (WebDriverTimeoutException)
            {
                throw new Exception($"在{timeout}秒内未在元素内找到为 '{by.Criteria}' 的子元素。");
            }
            catch (OperationCanceledException)
            {
                throw new OperationCanceledException("操作已被取消。");
            }
            catch (Exception ex)
            {
                throw new Exception($"在元素内查找为 '{by.Criteria}' 的子元素时发生错误：{ex.Message}", ex);
            }
        }

        private bool MeetsCondition(IWebElement element, int ckType)
        {
            #region ckType
            //1、Displayed:检查元素是否可见
            //2、Enabled: 检查元素是否启用（可交互）。
            //3、Selected: 对于可选元素（如复选框或单选按钮），检查是否选中。

            //4、Location: 获取元素在页面上的位置。
            //5、Size: 获取元素的尺寸。
            //6、TagName: 获取元素的标签名。
            //7、Text: 获取元素的文本内容。 
            #endregion
            switch (ckType)
            {
                case 0: return true;
                case 1: return element.Displayed;
                case 2: return element.Enabled;
                case 3: return element.Selected;
                case 12: return element.Displayed && element.Enabled;
                case 13: return element.Displayed && element.Selected;
                default: return false;
            }

        }


        public async Task SwitchToWindow(CancellationToken token, IWebDriver myDriver, string url, int timeout = 60)
        {
            try
            {
                for (int wi = 0; wi <= timeout; wi++)
                {
                    var windowHandles = myDriver.WindowHandles;
                    foreach (var wh in windowHandles)
                    {
                        myDriver.SwitchTo().Window(wh); //先切换
                        if (myDriver.Url.ToUpper().Contains($"{url}".ToUpper())) { return; }//再判断                       
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested(); 
                    if (wi == timeout) { throw new Exception($"未找到弹出的新窗口{url}"); }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public async Task SwitchToIFrame22(CancellationToken token, IWebDriver myDriver, By by, string documentUrl, int timeout = 60, int ckType = 1)
        {
            try
            {
                List<IWebElement> iframes = await FindByDriver22(token, myDriver, by, timeout, ckType);
                foreach (var iframe in iframes)
                {
                    string url = (string)((IJavaScriptExecutor)myDriver).ExecuteScript("return arguments[0].contentDocument.URL;", iframe); //#document
                    if (url.ToUpper().Contains(documentUrl.ToUpper())) { myDriver.SwitchTo().Frame(iframe); return; }
                }
            }
            catch (Exception ex) { throw ex; }
        }


        public async Task PGDN_CSNO(CancellationToken token)
        {
            //翻页（待修改）                   
            try
            {
                IntPtr hWnd = FindWindow(null, "供应商对账表-修改");//001E0AF6
                if (hWnd != IntPtr.Zero)
                {
                    SetForegroundWindow(hWnd); // 激活窗口
                    SendKeys.SendWait("{PGDN}");// 发送 Page Down 键                          
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested(); 
                }
            }
            catch { }

        }


       
         
        [DllImport("user32.dll")] private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);
        [DllImport("user32.dll")] private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);
        public async Task DisableScreen(CancellationToken token)
        {           

            //鼠标事件
            mouse_event(0x0001, 1, 0, 0, 0); //模拟鼠标轻微移动  MOUSEEVENTF_MOVE=0x0001
            mouse_event(0x0001, -1, 0, 0, 0);// 再移回原位
            await Task.Delay(10, token); token.ThrowIfCancellationRequested();

            ////键盘事件
            keybd_event(0x90, 0, 0, 0);     //按下Num Lock 键      VK_NUMLOCK = 0x90;
            keybd_event(0x90, 0, 0x0002, 0);//释放Num Lock 键      KEYEVENTF_KEYUP = 0x0002;
            await Task.Delay(10,token); token.ThrowIfCancellationRequested();
            keybd_event(0x90, 0, 0, 0);     //按下Num Lock 键      VK_NUMLOCK = 0x90;
            keybd_event(0x90, 0, 0x0002, 0);//释放Num Lock 键      KEYEVENTF_KEYUP = 0x0002;

        }


        public async Task WechatPost(string key, string jsonContent)
        {
            string url = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key="+key.Trim(); 
            using (HttpClient client = new HttpClient())
            {
                HttpContent content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                try
                {
                    HttpResponseMessage response = await client.PostAsync(url, content);
                    response.EnsureSuccessStatusCode();
                    string responseBody = await response.Content.ReadAsStringAsync();
                    Console.WriteLine("Response: " + responseBody);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Request failed: " + ex.Message);
                }
            }
        }




    }
}

