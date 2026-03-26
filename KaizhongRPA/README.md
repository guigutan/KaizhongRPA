部分文件超过100M，未提交。
请补充以下文件：
1、浏览器：chrome 129.0.6632.0  （chromium）
2、浏览器驱动：chromedriver 
3、...


WEB自动化：
	浏览器：chrome
	浏览器驱动：chromedriver
	控制框架：Selenium

SAP自动化：
	（程序集com组件找不到说明SAP GUI没有正确安装)
	1、Sap Gui Scripting Api
	2、SapROTWr 1.0 Type Library
		引用：
		1、using SAPFEWSELib;
		2、using SapROTWr;
	3、SAP元素定位工具（Scripting Tracker）：Tracker.exe
		位置：C:\Users\00127185\AppData\Local\Programs\Tracker
	
	4、SAP Scripting文档
	
		文档：https://help.sap.com/doc/ee3b7c74d6d545468a82460d0a8bbbe6/760.00/en-US/sap_gui_scripting_api.pdf
		下载：https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/babdf65f4d0a4bd8b40f5ff132cb12fa.html?version=760.00&locale=en-US
		
	5、SPA相关设置
		1、启用脚本：登录前/后主界面>>>定制本地布局Alt+F12（图标）>>>选项>>>左侧的《辅助功能与脚本-脚本》>>>勾选《启用脚本》>>>取消勾选其他>>>✔确定
			位于：计算机\HKEY_CURRENT_USER\SOFTWARE\SAP\SAPGUI Front\SAP Frontend Server\Security的[UserScripting、WarnOnAttach、WarnOnConnection]
		2、安全模式：登录前/后主界面>>>定制本地布局Alt+F12（图标）>>>选项>>>左侧的《安全性-安全设置》>>>安全模式的状态改为《禁用》>>>✔确定
			位于：
		3、对话模式：登录后主界面>>>帮助>>>F1帮助>>>对话框模式>>>F4帮助>>>对话(模式)>>>✔确定
			位于：


	NuGet包：
		1、Newtonsoft.Json
		2、控制框架：Selenium（Selenium.WebDriver+Selenium.Support）已安装4.27.0版本
			说明：Selenium 4.xx.x 默认与 Chrome v75 及更高版本兼容
		3、NPOI


	邮件通知：
		1、依据版本引用：Microsoft Outlook 15.0 Object Library
		2、设置：Outlook2013---邮件管理员方式运行---文件---选项---信任中心---信任中心设置---编程访问---从不向我发出可疑活动警告