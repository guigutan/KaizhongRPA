using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SAPFEWSELib;
using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace KaizhongRPA
{       


    public static class MyGlobal
    {
        public static string Version = "1.0.0";
        public static CancellationTokenSource MyCts;
        public static Task MyTask;

        public static bool Odd=true;
    }

    public static class MyPath
    {
        public static string App { get; } = System.AppDomain.CurrentDomain.BaseDirectory; //含“/”
        public static string Documents { get; } = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);       
        public static string UserRpaList { get; } = Path.Combine(MyPath.Documents, typeof(MyPath).Namespace, "UserRpaList.conf");
        public static string RunToStart { get; } = Path.Combine(MyPath.Documents, typeof(MyPath).Namespace, "RunToStart.conf");
        public static string DialogBox { get; } = Path.Combine(MyPath.Documents, typeof(MyPath).Namespace, "DialogBoxF1F4.conf");
        public static string SapLocalConfig { get; } = Path.Combine(MyPath.Documents, typeof(MyPath).Namespace, "SapLocalConfig.conf"); //文件存在即判断为已设置

        public static string AppData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);       

    }

    public static class RpaApi
    {
        public static string GetThisRpaInfo { get; } = "GetThisRpaInfo";
        public static string RpaMain { get; } = "RpaMain";     
    }

    public class RpaInfo
    {
        public string RpaClassName { get; set; }
        public string RpaName { get; set; }
        public string DefaultRunTime1 { get; set; }
        public string DefaultRunTime2 { get; set; }
        public string DefaultStatus { get; set; }
        public string DefaultPathStype { get; set; }
        public string DefaultConfigPath { get; set; }      
    }

    public static class Dvg1CN
    {
        public static string ClassName { get; } = "流程ID";
        public static string Name { get; } = "流程名称";
        public static string RunTime1 { get; } = "运行条件(从)";
        public static string RunTime2 { get; } = "运行条件(至)";
        public static string Status { get; } = "状态";
        public static string PathStype { get; } = "路径类型";
        public static string ConfigPath { get; } = "配置文件";
    }

    public static class LogRes
    {
        public static string E { get; } = "错误";
        public static string W { get; } = "提示";
        public static string S { get; } = "成功";
    }

    public static class MyWeb
    {
        public static string Chrome { get; set; } = MyPath.App + @"browser\chrome\chrome.exe";
        public static string ChromeDriver { get; set; } = MyPath.App + @"browser\chromedriver\chromedriver.exe";
        //public static IWebDriver Driver { get; set; }
        //public static ChromeOptions Options { get; set; }= new ChromeOptions() { BinaryLocation = Chrome };
        //public static  WebDriverWait Wait60 { get; set; } = new WebDriverWait(MyBS.Driver, TimeSpan.FromSeconds(60));
        //public static  WebDriverWait Wait30 { get; set; } = new WebDriverWait(MyBS.Driver, TimeSpan.FromSeconds(30));
        //public static  WebDriverWait Wait10 { get; set; } = new WebDriverWait(MyBS.Driver, TimeSpan.FromSeconds(10));
        //public static  WebDriverWait Wait3 { get; set; } = new WebDriverWait(MyBS.Driver, TimeSpan.FromSeconds(3));
        //public static  WebDriverWait Wait1 { get; set; } = new WebDriverWait(MyBS.Driver, TimeSpan.FromSeconds(1));

    }

    public static class MySap
    {
        public static GuiConnection Connection { get; set; }
        public static GuiSession Session { get; set; }
    }




}
