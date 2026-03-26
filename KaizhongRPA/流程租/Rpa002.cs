using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Newtonsoft.Json;
using System.Threading;
using System.Threading.Tasks;
using SAPFEWSELib;

using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using NPOI.SS.Formula.Functions;
using System.Diagnostics.Eventing.Reader;
using NPOI.Util;
using System.Reflection.Emit;
using System.Security.Cryptography;
using OpenQA.Selenium.BiDi.Modules.Input;
using static System.Runtime.CompilerServices.RuntimeHelpers;



namespace KaizhongRPA
{
    public class Rpa002:InvokeCenter
    {
        public RpaInfo GetThisRpaInfo()
        {
            RpaInfo rpaInfo = new RpaInfo();
            rpaInfo.RpaClassName = this.GetType().Name;
            rpaInfo.RpaName = "采购订单预付流程";
            rpaInfo.DefaultRunTime1 = "****-**-** **:**:**";
            rpaInfo.DefaultRunTime2 = "****-**-** **:**:**";
            rpaInfo.DefaultStatus = "有效";
            rpaInfo.DefaultPathStype = "相对路径";
            rpaInfo.DefaultConfigPath = @"config\RpaGroup\采购订单预付.xlsx";
            return rpaInfo;
        }

        #region 头部声明
        PublicClass publicClass = new PublicClass();
        SqlClass sqlClass = new SqlClass();        
        public IWebDriver MyDriver;
        public WebDriverWait wait60;      
        public string ScadaPODir = $@"{MyPath.Documents}\{typeof(MyPath).Namespace}\ZF003\ScadaPODir\";
        public string PDFDir = $@"{MyPath.Documents}\{typeof(MyPath).Namespace}\ZF003\PDF\";
        public string ScadaPOFN = "ScadaPO.txt";       
        public DataTable dt_config;
        public string Connstr;
        #endregion
        public async Task RpaMain(CancellationToken token, RpaInfo rpaInfo)
        {
            try
            {
                if (!Directory.Exists(ScadaPODir)) { Directory.CreateDirectory(ScadaPODir); }
                if (!Directory.Exists(PDFDir)) { Directory.CreateDirectory(PDFDir); }
                dt_config = await publicClass.ExcelToDataTable(token, rpaInfo.DefaultConfigPath);               
                if (!(dt_config != null && dt_config.Rows.Count > 0)) { return; }
                await ScadaPO(token);      //采集Y020的PO
                await Fill_ME23N(token);   //填写其他字段的值，
                await Fill_ZMMR010(token); //填写其他字段的值，银行国家代码
                await Fill_ZMMF007(token); //填写其他字段的值，下载附件
                await Fill_Conver(token);  //填写其他字段的值，计算值
                await Fill_Default(token); //填写其他字段的值，固定值
                await IsScadaFinsh(token); //是否采集完毕
                await IsIntact(token);     //是否填写完整
                await publicClass.ExitSap(token);
                await SpecialExec(token); //特殊处理


                await ExecOA(token);                                 
            }
            catch (Exception ex)
            {
                await publicClass.NoteLog(token, ex, dt_config);
            }
            finally
            {
                await publicClass.ExitSap(token);
            }
        }

       
        #region ScadaPO             

        private async Task ScadaPO(CancellationToken token)
        {
            try
            {
                await publicClass.DisableScreen(token);

                string paneStr = "";
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                await publicClass.GotoSapHome(token, dt_config,0,2);//登录SAP（已登录则退回首页）
                await EnterTransaction(token, "SE11");  //SE11事务码

                (MySap.Session.FindById("wnd[0]/usr/radRSRD1-TBMA") as GuiRadioButton).Selected = true;
                (MySap.Session.FindById("wnd[0]/usr/ctxtRSRD1-TBMA_VAL") as GuiCTextField).Text = "EKKO";
                (MySap.Session.FindById("wnd[0]/usr/btnPUSHSHOW") as GuiButton).Press();
                await Task.Delay(200, token); token.ThrowIfCancellationRequested();
                paneStr = (MySap.Session.FindById("wnd[0]/sbar/pane[0]") as GuiStatusPane).Text;
                if (paneStr != "") { throw new Exception(paneStr); }              


                (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(46); //Ctrl+Shift+F10 内容
                string ScadaBeforeDay = await publicClass.GetLibValue(token, dt_config, "ScadaBeforeDay");
                int days = 0 - Convert.ToInt32(ScadaBeforeDay.Trim());
                (MySap.Session.FindById("wnd[0]/usr/ctxtI8-LOW") as GuiCTextField).Text = DateTime.Now.AddDays(days).ToString("yyyy.MM.dd");//AEDAT             
                (MySap.Session.FindById("wnd[0]/usr/ctxtI8-HIGH") as GuiCTextField).Text = DateTime.Now.ToString("yyyy.MM.dd");//AEDAT
                (MySap.Session.FindById("wnd[0]/usr/ctxtI17-LOW") as GuiCTextField).Text = "G";//FRGKE  原16变17 2025.1.9
                (MySap.Session.FindById("wnd[0]/usr/txtMAX_SEL") as GuiTextField).Text = "5000";//最大命中数              
                (MySap.Session.FindById("wnd[0]/tbar[1]/btn[8]") as GuiButton).Press();//执行
                await Task.Delay(500, token); token.ThrowIfCancellationRequested();
                paneStr = (MySap.Session.FindById("wnd[0]/sbar/pane[0]") as GuiStatusPane).Text;
                if (paneStr != "") { throw new Exception(paneStr); }
               

                (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(33); //Ctrl+F9 选择字段
                await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                (MySap.Session.FindById("wnd[1]/tbar[0]/btn[14]") as GuiButton).Press(); //14  Shift+F2 取消全选              
                await SelectedColumn(token);//选中BEBLN和ZTERM字段
                (MySap.Session.FindById("wnd[1]/tbar[0]/btn[6]") as GuiButton).Press();//选好后F6应用
                await Task.Delay(200, token); token.ThrowIfCancellationRequested();            

                (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(20); //下载 20 Shift+F8
                await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                await SelectedDownStyle(token); //选择格式：含标签的文本
                await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                (MySap.Session.FindById("wnd[1]/tbar[0]/btn[0]") as GuiButton).Press(); //继续
                await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                await publicClass.ClearDir(token, ScadaPODir); //先清空下载目录
                (MySap.Session.FindById("wnd[1]/usr/ctxtDY_PATH") as GuiCTextField).Text = ScadaPODir; //下载目录
                (MySap.Session.FindById("wnd[1]/usr/ctxtDY_FILENAME") as GuiCTextField).Text = ScadaPOFN; //保存的文件名
                (MySap.Session.FindById("wnd[1]/tbar[0]/btn[0]") as GuiButton).Press(); //生成
                await Task.Delay(500, token); token.ThrowIfCancellationRequested();
                await InsertPO(token, TablePO, Connstr);//处理ScadaPO.txt数据并存入数据库              

                await publicClass.Backspace(token,3); 
            }
            catch (Exception ex) {await publicClass.ExitSap(token); throw ex; }
        }

        public async Task EnterTransaction(CancellationToken token, string transactionCode)
        {
            try
            {
                (MySap.Session.FindById("wnd[0]/tbar[0]/okcd") as GuiOkCodeField).Text = transactionCode;//输入事务码
                await Task.Delay(50, token); token.ThrowIfCancellationRequested();
                (MySap.Session.FindById("wnd[0]") as GuiFrameWindow).SendVKey(0);//回车
                await Task.Delay(200, token); token.ThrowIfCancellationRequested();
            }
            catch (Exception ex) { throw ex; }
        }

        public async Task SelectedColumn(CancellationToken token)
        {
            try
            {
                //选中BEBLN和ZTERM字段
                int SelectOK = 0;
                GuiComponentCollection allColl = (MySap.Session.FindById("wnd[1]/usr/") as GuiUserArea).Children;               
                foreach (GuiComponent myGuiComponent in allColl)
                {                   
                    if (myGuiComponent is GuiLabel thisGui)
                    {
                        if (thisGui.Text == "EBELN" || thisGui.Text == "ZTERM")
                        {
                            SelectOK +=1;
                            string ID= thisGui.Id.ToString();                                   //>>>     /app/con[0]/ses[0]/wnd[1]/usr/lbl[3,4]
                            string chkID= "wnd[1]/usr/chk[1" + ID.Substring(ID.IndexOf(","));  //假设在GuiCheckBox在第1列
                            (MySap.Session.FindById(chkID) as GuiCheckBox).Selected = true;
                            if (SelectOK >= 2) { break; }
                            await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        }
                    }
                    await Task.Delay(10, token); token.ThrowIfCancellationRequested();
                }
            }
            catch (Exception ex) { throw ex; }
        }

        public async Task SelectedDownStyle(CancellationToken token)
        {
            try
            {
                bool finded = false;
                //选择下载的格式：含标签的文本
                GuiComponentCollection allColl = (MySap.Session.FindById("wnd[1]/usr/") as GuiUserArea).FindAllByName("SPOPLI-SELFLAG", "GuiRadioButton");
                foreach (GuiComponent myGuiComponent in allColl)
                {
                    if (myGuiComponent is GuiRadioButton thisGui)
                    {
                        if (thisGui.Text == "含标签的文本")
                        {
                            finded=true;
                            string ID = thisGui.Id.ToString();//>>  /app/con[0]/ses[0]/wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]                           
                            (MySap.Session.FindById(ID.Substring(ID.IndexOf("wnd[1]/usr/"))) as GuiRadioButton).Selected = true;
                            break;
                        }
                    }
                    await Task.Delay(10, token); token.ThrowIfCancellationRequested();
                }
                if (!finded) { throw new Exception("未找到<含标签的文本>选项"); }
            }
            catch (Exception ex) { throw ex; }
        }

        public async Task InsertPO(CancellationToken token, string TablePO, string Connstr)
        {
            
            try
            {
                bool downFinish=false;
                string filePath = ScadaPODir + ScadaPOFN;
                for (int i = 0; i <30; i++)
                {
                    if (File.Exists(filePath)) { downFinish = true;break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                }
                if (!downFinish) { throw new Exception("SAP下载采集数据，已超过30秒仍未完成。"); }

                using (StreamReader reader = new StreamReader(filePath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        string PayCode = "Y020";
                        //Replace方式：必须是仅有BEBLN和ZTERM两列的情况时才可行
                        if (line.Contains(PayCode))
                        {
                            string PO = line.Replace(PayCode, "").Replace("\t", "").Replace(" ","").Replace("|", "");
                            string sql = $"insert into {TablePO}(PO,PayCode) values('{PO}','{PayCode}')";
                            sqlClass.InsSQL(sql, Connstr);
                            await Task.Delay(10, token); token.ThrowIfCancellationRequested();
                        }
                    }
                }
            }
            catch (Exception ex) { throw ex; }
        }

        #endregion


        #region Fill_ME23N       
        private async Task Fill_ME23N(CancellationToken token)
        {
            try
            {
                List<string> list1 = new List<string>() { "PUser", "Supplier", "SupplierCode", "SupplierShortName", "Currency", "PGroup", "CompanyCode", "DeliveryDate" };//ME23N
                List<string> list = list1.Distinct().ToList();

                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                string sql = $"select * from {TablePO} where IsScadaFinsh=0 and  IsRequest<=0";
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
               
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        await publicClass.DisableScreen(token);

                        bool b2 = await IsFillScada(token, dt, i, list);
                        if (b2) { continue; } //优化：list列都有值时跳过采集

                        string ID = dt.Rows[i]["ID"].ToString();
                        string PO = dt.Rows[i]["PO"].ToString();
                        await publicClass.GotoSapHome(token, dt_config,0,3);//登录SAP（已登录则退回首页）
                        await EnterTransaction(token, "ME23N");  //ME23N采购订单事务码
                        (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(17); //Shift+F5 其他采购订单
                        ((MySap.Session.FindById("wnd[1]/usr/") as GuiUserArea).FindAllByName("MEPO_SELECT-BSTYP_F", "GuiRadioButton").Item(0) as GuiRadioButton).Selected = true;  //选中采购订单
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        ((MySap.Session.FindById("wnd[1]/usr/") as GuiUserArea).FindAllByName("MEPO_SELECT-EBELN", "GuiCTextField").Item(0) as GuiCTextField).Text = PO;  //输入采购订单
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        (MySap.Session.FindById("wnd[1]") as GuiModalWindow).SendVKey(0); //回车
                        await Task.Delay(500, token); token.ThrowIfCancellationRequested();

                        //表头：PO号，供应商
                        string thisPO = ((MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).FindAllByName("MEPO_TOPLINE-EBELN", "GuiTextField").Item(0) as GuiTextField).Text.Trim();
                        if (!thisPO.Contains(PO)) { await publicClass.Backspace(token, 3); continue; }
                        string Supplier = ((MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).FindAllByName("MEPO_TOPLINE-SUPERFIELD", "GuiCTextField").Item(0) as GuiTextField).Text.Trim();
                        string SupplierCode = await GetSupplierCode(token, Supplier.Trim());
                        string SupplierShortName = SupplierCode;


                        //选项卡：《机构数据》
                        (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(26);//26 Ctrl+F2 扩展抬头
                        ((MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).FindAllByName("TABHDT8", "GuiTab").Item(0) as GuiTab).Select(); //点击《机构数据》选项卡
                        await Task.Delay(200, token); token.ThrowIfCancellationRequested();
                        string PGroup = ((MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).FindAllByName("MEPO1222-EKGRP", "GuiCTextField").Item(0) as GuiCTextField).Text.Trim(); //获取采购组的值
                        string CompanyCode = ((MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).FindAllByName("MEPO1222-BUKRS", "GuiCTextField").Item(0) as GuiCTextField).Text.Trim(); //获取公司代码的值


                        //选项卡：《采购员信息》
                        (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(26);//26 Ctrl+F2 扩展抬头
                        ((MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).FindAllByName("TABHDT9", "GuiTab").Item(0) as GuiTab).Select();//点击《采购员信息》选项卡
                        await Task.Delay(200, token); token.ThrowIfCancellationRequested();
                        string PUser = ((MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).FindAllByName("CI_EKKODB-ZCGYTMS", "GuiTextField").Item(0) as GuiTextField).Text.Trim(); //获取采购员描述的值

                        //选项卡：《支付/开票》
                        (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(26);//26 Ctrl+F2 扩展抬头               
                        ((MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).FindAllByName("TABHDT1", "GuiTab").Item(0) as GuiTab).Select(); //点击《支付/开票》选项卡
                        await Task.Delay(200, token); token.ThrowIfCancellationRequested();
                        string Currency = ((MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).FindAllByName("MEPO1226-WAERS", "GuiCTextField").Item(0) as GuiCTextField).Text.Trim(); //获取货币的值

                        //扩展：《项目概况》
                        (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(27); //27 Ctrl+F3 扩展项目概况
                        string DeliveryDate_JsonStr = await GetDeliveryDate(token); //获取交货日期的值（多个，json）
                       
                        /*
                            PUser varchar(255),				--采购员描述
                            Supplier varchar(255),			--供应商
                            SupplierCode varchar(255),		--供应商代码
                            SupplierShortName varchar(255),	--供应商简称
                            Currency varchar(255),			--货币
                            PGroup varchar(255),			--采购组	
                            CompanyCode varchar(255),		--公司代码
                            DeliveryDate nvarchar(max),		--交货日期(json)

                         */
                        string sqlUpd = $"update t_PO set PUser='{PUser}',Supplier='{Supplier}',SupplierCode='{SupplierCode}',SupplierShortName='{SupplierShortName}',Currency='{Currency}',PGroup='{PGroup}',CompanyCode='{CompanyCode}',DeliveryDate='{DeliveryDate_JsonStr}' ";
                        sqlUpd += $" where ID={ID}";
                        sqlClass.UpdSQL(sqlUpd, Connstr);

                        await publicClass.Backspace(token, 3);                      
                    }
                    catch (Exception ex)
                    { 
                        await publicClass.NoteLog(token, ex,dt_config);
                        /*异常>>继续下一个迭代*/ 
                    }
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                }
               
            }
            catch (Exception ex) { throw ex; }
        }

        private async Task<string> GetSupplierCode(CancellationToken token,string supplier)
        {
            char[] nums = new char[10] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
            string SupplierNO = "";
            foreach (char item in supplier)
            {
                await Task.Delay(0, token); token.ThrowIfCancellationRequested();
                if (nums.Contains(item))
                {
                    SupplierNO += item.ToString();
                }
                else
                {
                    //不再是数字时，忽略后面跳出
                    break;
                }
            }
            return SupplierNO;
        }

        public async Task<string> GetDeliveryDate(CancellationToken token)
        {
            string DeliveryDateStr = "";

            List<string> dateList=new List<string>();
            GuiComponentCollection allColl = (MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).FindAllByName("MEPO1211-EEIND", "GuiCTextField");
            foreach (GuiComponent myGuiComponent in allColl)
            {
                if (myGuiComponent is GuiCTextField thisGui)
                {
                    if (thisGui.Text != "")
                    {
                        dateList.Add(Convert.ToDateTime(thisGui.Text).ToString("yyyy-MM-dd"));
                    }
                    else { break; }
                }
                await Task.Delay(0, token); token.ThrowIfCancellationRequested();
            }
            if (dateList.Count > 0) { DeliveryDateStr = JsonConvert.SerializeObject(dateList, Formatting.Indented); }

            return DeliveryDateStr;           
        }

        #endregion


        #region Fill_ZMMR010
        private async Task Fill_ZMMR010(CancellationToken token)
        {
            try
            {
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
               // string sql = $"select * from {TablePO} where IsScadaFinsh=0 and  IsRequest<=0";
                string sql = $"select * from {TablePO} where IsScadaFinsh=0 and  IsRequest<=0 and (SWIFTCode is null or SWIFTCode='') ";//优化限定SWIFTCode，这里的只upd SWIFTCode时才可用
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string paneText = "";
                    try
                    {
                        await publicClass.DisableScreen(token);

                        string ID = dt.Rows[i]["ID"].ToString();
                        string SupplierCode = dt.Rows[i]["SupplierCode"].ToString();
                        string CompanyCode = dt.Rows[i]["CompanyCode"].ToString();
                        if (SupplierCode == "" || CompanyCode == "") { continue; }
                        await publicClass.GotoSapHome(token, dt_config, 0, 3);//登录SAP（已登录则退回首页）
                        await EnterTransaction(token, "ZMMR010"); //事务ZMMR010 >银行国家代码
                        (MySap.Session.FindById("wnd[0]/usr/ctxtS_LIFNR-LOW") as GuiCTextField).Text = SupplierCode; //输入供应商代码                 
                        (MySap.Session.FindById("wnd[0]/usr/ctxtS_BUKRS-LOW") as GuiCTextField).Text = CompanyCode;//输入公司代码
                        (MySap.Session.FindById("wnd[0]/usr/radP_R3") as GuiRadioButton).Selected = true;//选中供应商清单选项
                        (MySap.Session.FindById("wnd[0]/tbar[1]/btn[8]") as GuiButton).Press();//F8执行
                        await Task.Delay(500, token); token.ThrowIfCancellationRequested();

                        try { paneText = (MySap.Session.FindById("wnd[0]/sbar/pane[0]") as GuiStatusPane).Text.Trim(); } catch { }
                        if (paneText != "") { throw new Exception(paneText); }

                        bool chooselayout = await publicClass.ChooseLayout(token, "001SCODE", false, 20);//选择布局：仅显示银行国家代码(RPA采集)
                        if (!chooselayout) { throw new Exception("选择001SCODE布局错误"); }

                        (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(46);//Ctrl+Shift+F10打印预览
                        await Task.Delay(500, token); token.ThrowIfCancellationRequested();

                        string SWIFTCode = "";
                        //获取<银行国家代码>标题列的ID
                        string swiftcodeid = "";
                        GuiComponentCollection allColl = (MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).Children;
                        foreach (GuiComponent myGuiComponent in allColl)
                        {
                            if (myGuiComponent is GuiLabel thisGui)
                            {
                                if (thisGui.Text == "银行国家代码")
                                {
                                    swiftcodeid = thisGui.Id;
                                }
                            }
                        }
                        if (swiftcodeid == "") { throw new Exception("未能找到<银行国家代码>的标题列"); }

                        //获取<银行国家代码>内容列的值
                        ///app/con[0]/ses[0]/wnd[0]/usr/lbl[1,2]
                        string before_swiftcodeid = swiftcodeid.Substring(0, swiftcodeid.IndexOf(",") + 1);
                        foreach (GuiComponent myGuiComponent in allColl)
                        {
                            if (myGuiComponent is GuiLabel thisGui)
                            {
                                if (thisGui.Text != "" && thisGui.Text != "银行国家代码" && thisGui.Id.Contains(before_swiftcodeid))
                                {
                                    SWIFTCode = thisGui.Text.Trim();
                                    break;
                                }
                            }
                        }

                        if (SWIFTCode == "") { throw new Exception("未能找到<银行国家代码>的值"); }
                        string sqlUpd = $"update t_PO set SWIFTCode='{SWIFTCode}' where ID={ID}";
                        sqlClass.UpdSQL(sqlUpd, Connstr);


                        await publicClass.Backspace(token, 3);
                    }
                    catch (Exception ex)
                    {
                        await publicClass.NoteLog(token, ex, dt_config);
                        /*异常>>继续下一个迭代*/
                    }
                    await Task.Delay(200, token); token.ThrowIfCancellationRequested();

                }
            }
            catch (Exception ex) { throw ex; }
        }

        #endregion


        #region Fill_ZMMF007
        private async Task Fill_ZMMF007(CancellationToken token)
        {
            try
            {
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                string sql = $"select * from {TablePO} where IsScadaFinsh=0 and IsRequest<=0";
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        await publicClass.DisableScreen(token);

                        string ID = dt.Rows[i]["ID"].ToString();
                        string PO = dt.Rows[i]["PO"].ToString();
                        if (PO == "") { continue; }

                        //优化：这里仅下载时方适用
                        if (dt.Rows[i]["FilePath"].ToString() != "") 
                        {
                            string destPath = $"{PDFDir}{ID}_{PO}{Path.GetExtension(dt.Rows[i]["FilePath"].ToString())}";
                            if (File.Exists(destPath)) { continue; } 
                        }
                        

                        await publicClass.GotoSapHome(token, dt_config, 0, 3);//登录SAP（已登录则退回首页）
                        await EnterTransaction(token, "ZMMF007"); //事务ZMMF007 >下载PDF
                        (MySap.Session.FindById("wnd[0]/usr/ctxtS_WERKS-LOW") as GuiCTextField).Text = "*";//输入工厂
                        (MySap.Session.FindById("wnd[0]/usr/ctxtS_EBELN-LOW") as GuiCTextField).Text = PO;//输入采购订单
                        (MySap.Session.FindById("wnd[0]/usr/radP4") as GuiRadioButton).Selected = true;//选择<PDF导出本地>选项
                        (MySap.Session.FindById("wnd[0]/tbar[1]/btn[8]") as GuiButton).Press();//执行
                        await Task.Delay(500, token); token.ThrowIfCancellationRequested();

                        string paneText = (MySap.Session.FindById("wnd[0]/sbar/pane[0]") as GuiStatusPane).Text;
                        string FilePath = @"D:\" + $"{DateTime.Now.ToString("yyyyMMdd")}-{PO}.PDF";
                        if (paneText.ToUpper().Contains(FilePath.ToUpper()))
                        {
                            string sqlUpd = $"update t_PO set FilePath='{FilePath}' where ID={ID}";
                            sqlClass.UpdSQL(sqlUpd, Connstr);
                        }
                        else
                        {
                            throw new Exception(paneText);
                        }
                    }
                    catch (Exception ex)
                    {
                        await publicClass.NoteLog(token, ex, dt_config);
                        /*异常>>继续下一个迭代*/
                    }
                    await Task.Delay(200, token); token.ThrowIfCancellationRequested();
                }
            }
            catch (Exception ex) { throw ex; }

        }

        #endregion


        #region Fill_Conver
        private async Task Fill_Conver(CancellationToken token)
        {
            try
            {
                List<string> list1 = new List<string>() { "Class2", "InvoiceDate", "PayFinish", "PayDateExplained" };//计算值
                List<string> list = list1.Distinct().ToList();

                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                string sql = $"select * from {TablePO} where IsIntact=0 and  IsRequest<=0";
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }

                string Class2Path = await publicClass.GetLibValue(token, dt_config, "Class2");
                if (!File.Exists(Class2Path)) { throw new Exception($"配置文件不存在。{Class2Path}"); }
                DataTable dt_class2 = await publicClass.ExcelToDataTable(token, Class2Path);
                if (!(dt_class2 != null && dt_class2.Rows.Count > 0)) { throw new Exception($"配置文件无内容。{Class2Path}"); }
               
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    await publicClass.DisableScreen(token);

                    bool b2 = await IsFillScada(token, dt, i, list);
                    if (b2) { continue; } //优化：list列都有值时跳过计算

                    string ID = dt.Rows[i]["ID"].ToString();
                    string PGroup = dt.Rows[i]["PGroup"].ToString();
                    string DeliveryDate_Json = dt.Rows[i]["DeliveryDate"].ToString();
                    if (PGroup == "" || DeliveryDate_Json == "") { continue; }
                    string Class2 = await publicClass.GetClass2(token, PGroup, dt_class2);
                    List<string> date_list = JsonConvert.DeserializeObject<List<string>>(DeliveryDate_Json);

                    string InvoiceDate = "";
                    string PayFinish = "";
                    string PayDateExplained = "款到发货";
                    if (date_list.Count > 0)
                    {
                        DateTime DeliveryDate0 = Convert.ToDateTime(date_list[0].ToString());
                        DateTime thisM = new DateTime(DeliveryDate0.Year, DeliveryDate0.Month, DateTime.DaysInMonth(DeliveryDate0.Year, DeliveryDate0.Month));
                        DateTime LastM = new DateTime(DeliveryDate0.AddMonths(1).Year, DeliveryDate0.AddMonths(1).Month, DateTime.DaysInMonth(DeliveryDate0.AddMonths(1).Year, DeliveryDate0.AddMonths(1).Month));
                        InvoiceDate = DeliveryDate0.Day < 20 ? thisM.ToString("yyyy-MM-dd") : LastM.ToString("yyyy-MM-dd");

                        PayFinish = DeliveryDate0.AddDays(-3).ToString("yyyy-MM-dd");
                    }
                    /*
                        Class2 varchar(255),			--二级分类（采购组+二级分类配置表）
                        InvoiceDate date,				--预计金税发票提供时间 （交货日期<20号为交货月月末，否则为交货月次月月末）
                        PayFinish date,					--付款要求完成时间（交期日期的前3天）
                        PayDateExplained varchar(255),	--付款日期说明（款到发货）
                    */
                    string _Class2 = Class2 == "" ? "NULL" : $"'{Class2}'";
                    string _InvoiceDate = InvoiceDate == "" ? "NULL" : $"'{InvoiceDate}'";
                    string _PayFinish = PayFinish == "" ? "NULL" : $"'{PayFinish}'";
                    string _PayDateExplained = PayDateExplained == "" ? "NULL" : $"'{PayDateExplained}'";

                    string sqlUpd = $"update t_PO set Class2={_Class2},InvoiceDate={_InvoiceDate},PayFinish={_PayFinish},PayDateExplained={_PayDateExplained} where ID={ID}";
                    sqlClass.UpdSQL(sqlUpd, Connstr);
                }
            }
            catch (Exception ex) { throw ex; }
        }

        #endregion


        #region Fill_Default
        private async Task Fill_Default(CancellationToken token)
        {
            await publicClass.DisableScreen(token);

            await Task.Delay(0, token); token.ThrowIfCancellationRequested();
            //数据库已设置Default,如有变动再在此处理。


            /*
                PayStype varchar(255) not null default('预付材料款'),	--支付类型	
                BankOf  varchar(255) not null default('对公'),			--对公对私
                IsLinkPR varchar(255)not null default('否'),			--是否关联资本性支出请购流程
                PurchaseStype varchar(255) not null default('款到发货'),--分期类型
                IsInfrastructure varchar(255) not null default('否'),	--是否基建类
                IsExistPO varchar(255) not null default('是'),			--是否有采购订单号

             */
        }

        #endregion


        #region IsScadaFinsh
        private async Task IsScadaFinsh(CancellationToken token)
        {
            try 
            {

                List<string> list1 = new List<string>() { "PUser", "Supplier", "SupplierCode", "SupplierShortName", "Currency", "PGroup", "CompanyCode", "DeliveryDate"};//ME23N
                List<string> list2 = new List<string>() { "SWIFTCode"};//ZMMR010
                List<string> list3 = new List<string>() { "FilePath" };//ZMMF007
                list1.AddRange(list2);
                list1.AddRange(list3);
                List<string> list= list1.Distinct().ToList();

                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                //string sql = $"select * from {TablePO} where IsScadaFinsh=0 and IsIntact=0 and  IsRequest<=0";
                string sql = $"select * from {TablePO} where IsRequest<=0 ";
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        await publicClass.DisableScreen(token);

                        string ID = dt.Rows[i]["ID"].ToString();
                        string PO = dt.Rows[i]["PO"].ToString();
                        string FilePath = dt.Rows[i]["FilePath"].ToString();
                        string destPath = $"{PDFDir}{ID}_{PO}{Path.GetExtension(FilePath)}";
                        await RemoveFile(token,ID, PO, FilePath, destPath);
                        bool b1 = File.Exists(destPath);
                        bool b2 = await IsFillScada(token, dt, i, list);

                        string IsScadaFinsh = (b1 && b2) ? "1" : "0";
                        string sqlUpd = "";
                        if (IsScadaFinsh == "1") { sqlUpd = $"update t_PO set IsScadaFinsh='{IsScadaFinsh}' where ID={ID}"; }
                        else { sqlUpd = $"update t_PO set IsScadaFinsh='{IsScadaFinsh}',IsIntact=0 where ID={ID}"; }                      
                        sqlClass.UpdSQL(sqlUpd, Connstr);
                    }
                    catch (Exception ex) { await publicClass.NoteLog(token, ex, dt_config); /*异常>>继续下一个迭代*/}
                }
            }
            catch (Exception ex) { throw ex; }
        }       

        private async Task RemoveFile(CancellationToken token, string ID, string PO, string FilePath,string destPath)
        {
            try 
            {
                if (File.Exists(FilePath))
                {                   
                    File.Move(FilePath, destPath);
                    for (int i = 0; i < 10; i++)
                    {                        
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                        if (File.Exists(destPath)) { break; }
                    }                   
                }
            }
            catch { }
        }

        private async Task<bool> IsFillScada(CancellationToken token, DataTable dt, int i, List<string> list)
        {
            bool res = true;
            foreach (string cname in list) 
            {
                string str = dt.Rows[i][$"{cname}"].ToString();
                if (str == null || str == "") { res = false;break; }
                await Task.Delay(0, token); token.ThrowIfCancellationRequested();
            }
            return res;
        }

        #endregion


        #region IsIntact
        private async Task IsIntact(CancellationToken token)
        {
            try
            {
                List<string> list1 = new List<string>() { "Class2", "InvoiceDate", "PayFinish", "PayDateExplained" };//计算值
                List<string> list = list1.Distinct().ToList();

                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");                
                string sql = $"select * from {TablePO} where IsScadaFinsh=1 and IsRequest<=0 "; //仅遍历采集完毕的
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        await publicClass.DisableScreen(token);

                        string ID = dt.Rows[i]["ID"].ToString();                      
                        bool b2 = await IsFillScada(token, dt, i, list);
                        string IsIntact = (b2) ? "1" : "0";
                        string sqlUpd = $"update t_PO set IsIntact='{IsIntact}' where ID={ID}";
                        sqlClass.UpdSQL(sqlUpd, Connstr);
                    }
                    catch (Exception ex) { await publicClass.NoteLog(token, ex, dt_config); /*异常>>继续下一个迭代*/}
                }
            }
            catch (Exception ex) { throw ex; }
        }

        #endregion


        #region SpecialExec
        private async Task SpecialExec(CancellationToken token)
        {
            try
            {

               
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                string sql = $"select * from {TablePO} where IsIntact=1 and IsRequest<=0 "; //仅遍历采集完毕的
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        await publicClass.DisableScreen(token);

                        //----处理1：要求付款完成时间小于当前日期----改为当前日期加3天-------
                        string ID = dt.Rows[i]["ID"].ToString();                       
                        string PayFinish = dt.Rows[i]["PayFinish"].ToString();
                        if (Convert.ToDateTime(PayFinish) < DateTime.Now)
                        {
                            string sqlUpd1 = $"update t_PO set PayFinish='{DateTime.Now.AddDays(3).ToString("yyyy-MM-dd")}' where ID={ID}";
                            sqlClass.UpdSQL(sqlUpd1, Connstr);
                        }

                        //----处理2：采购员离职替换-----------
                        string PUser = dt.Rows[i]["PUser"].ToString();
                        string NewPUser=  await GetReplacePUser(token, PUser);
                        if (NewPUser != "")
                        {
                            sqlClass.UpdSQL($"update {TablePO} set PUser='{NewPUser}' where ID={ID}", Connstr);
                        }


                    }
                    catch (Exception ex) { await publicClass.NoteLog(token, ex, dt_config); /*异常>>继续下一个迭代*/}
                }




                //------处理3：删除用户文件夹ZF003\PDF\下的文档-----仅保留最新的100个---------------
                FileInfo[] allPDF =new DirectoryInfo(PDFDir).GetFiles();
                int delCount = 100;
                if (dt != null && dt.Rows.Count > delCount) { delCount = dt.Rows.Count; }
                var filesToDelete = allPDF.OrderByDescending(f => f.CreationTimeUtc).Skip(delCount).ToList();
                foreach (var file in filesToDelete)
                {
                    try
                    {
                        file.Delete();
                    }
                    catch { }
                }

                ////----处理4：删除一年前的数据库--(确保客户机时间正确)>>改为手动删除--------------------------------
                //string sqlDel3 = $"delete from t_PO where ScadaDateTime<='{DateTime.Now.AddDays(-365).ToString("yyyy-MM-dd")}' ";
                //sqlClass.DelSQL(sqlDel3, Connstr);





            }
            catch (Exception ex) { throw ex; }
        }

        private async Task<string> GetReplacePUser(CancellationToken token, string PUser)
        {
            string NewPUser = "";
            try
            {
                DataTable dt_ReplacePUser = new DataTable();
                string pathReplacePUser = await publicClass.GetLibValue(token, dt_config, "ReplacePUser");
                if (File.Exists(pathReplacePUser)) { dt_ReplacePUser = await publicClass.ExcelToDataTable(token, pathReplacePUser); }

                if (dt_ReplacePUser == null || dt_ReplacePUser.Rows.Count <= 0) { return ""; }
                for (int i = 0; i < dt_ReplacePUser.Rows.Count; i++)
                {
                    if (PUser == dt_ReplacePUser.Rows[i]["离职采购员"].ToString())
                    {
                        NewPUser = dt_ReplacePUser.Rows[i]["顶替采购员"].ToString();
                        break;
                    }
                    await Task.Delay(0, token); token.ThrowIfCancellationRequested();
                }
            }
            catch { }

            return NewPUser;
        }


        #endregion





        private async Task ExecOA(CancellationToken token)
        {
            DataTable dt = new DataTable();
            string WechatKey = "";
            string TablePO = "";
            try {
                Connstr = await publicClass.GetConnstr(token, dt_config);
                TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                WechatKey = await publicClass.GetLibValue(token, dt_config, "WechatKey");
                string sql = $"select * from {TablePO} where IsIntact=1 and IsRequest<=0 ";
                dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
            }
            catch (Exception ex) { throw ex; }
            
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string ID = "";
                try
                {
                    await publicClass.DisableScreen(token);

                    string PUser = dt.Rows[i]["PUser"].ToString();
                    string Supplier = dt.Rows[i]["Supplier"].ToString();
                    string SupplierCode = dt.Rows[i]["SupplierCode"].ToString();
                    string Currency = dt.Rows[i]["Currency"].ToString();
                    string PayCode = dt.Rows[i]["PayCode"].ToString();
                    string PayStype = dt.Rows[i]["PayStype"].ToString();
                    string Class2 = dt.Rows[i]["Class2"].ToString();
                    string SupplierShortName = dt.Rows[i]["SupplierShortName"].ToString();
                    string CompanyCode = dt.Rows[i]["CompanyCode"].ToString();
                    string BankOf = dt.Rows[i]["BankOf"].ToString();
                    string InvoiceDate = dt.Rows[i]["InvoiceDate"].ToString();
                    string IsLinkPR = dt.Rows[i]["IsLinkPR"].ToString();
                    string PurchaseStype = dt.Rows[i]["PurchaseStype"].ToString();
                    string SWIFTCode = dt.Rows[i]["SWIFTCode"].ToString();
                    string IsInfrastructure = dt.Rows[i]["IsInfrastructure"].ToString();
                    string IsExistPO = dt.Rows[i]["IsExistPO"].ToString();
                    string PayFinish = dt.Rows[i]["PayFinish"].ToString();
                    string PO = dt.Rows[i]["PO"].ToString();
                    ID = dt.Rows[i]["ID"].ToString();
                    string RequestID = dt.Rows[i]["RequestID"].ToString();
                    //string FilePath = dt.Rows[i]["FilePath"].ToString();
                    string FilePath = $"{PDFDir}{ID}_{PO}{Path.GetExtension(dt.Rows[i]["FilePath"].ToString())}"; //待优化，固定

                     await  PostWillDo(PO,PUser,Supplier,CompanyCode,PayCode, WechatKey);//POST 执行前的预告


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
                        //----------------------------------------------------------------------------------------
                        await LoginOA(token, dt_config);    //01-登录OA
                        await Nav_Liucheng(token);          //02-选择顶部菜单《流程》
                        await New_Liucheng(token);          //03-新建流程
                        await ClickZF003(token);            //04-点击ZF003流程，并切换新窗口

                        await FillPUser(token, PUser);                          //01-填写申请人（新窗口）
                        await FillSupplierCode(token, SupplierCode);            //02-填写供应商代码
                        await FillCurrency(token, Currency);                    //03-填写支付币种（货币）
                        await FillPayStype(token, PayStype);                    //04-选择付款类型（支付类型）>>预付材料款[默]
                        await FillClass2(token, Class2);                        //05-填写二级分类
                        await FillSupplierShortName(token, SupplierShortName);  //06-填写供应商简称
                        await FillCompanyCode(token, CompanyCode);              //07-填写付款公司（公司代码）
                        await FillBankOf(token, BankOf);                        //08-选择对公对私
                        await FillInvoiceDate(token, InvoiceDate);              //09-填写预计金税发票提供时间
                        await FillIsLinkPR(token, IsLinkPR);                    //10-选择是否关联资本性支出请购流程
                        await FillPurchaseStype(token, PurchaseStype);          //11-选择分期类型
                        await FillSWIFTCode(token, SWIFTCode);                  //12-填写收款人国家/地区（银行国家代码）
                        await FillIsInfrastructure(token, IsInfrastructure);    //13-选择是否基建类
                        await FillIsIsExistPO(token, IsExistPO);                //14-选择是否有采购订单号
                        await FillPayFinish(token, PayFinish);                  //15-填写付款要求完成时间
                        await FillPO_NO(token, PO, ID);                         //16-填写PO号
                        await ResPayInfo_SaveBefore(token,ID);                  //17-第一次刷新支付信息-支付金额（必须在保存前） (有异常情况)
                        await SaveOA(token, ID);                                //18-第一次保存OA 
                        await FillPO_Detail(token, PO, ID);                     //19-填写PO价格等其他明细信息（必须在保存后）
                        await ResPayInfo_SaveAfter(token);                      //20-第二次刷新支付信息-支付金额（价格更新后再次刷新）
                        await FillFilePath(token, FilePath);                    //21-填写相关附件-上传附件
                        await SaveOA(token, ID);                                //22-第二次保存OA
                        await CheckPayInfo(token, ID);                          //23-检查支付信息必填项-需先保存OA
                        await SubOA(token, ID);                                 //24-提交OA

                    }
                }
                catch (Exception ex) { await publicClass.NoteLog(token, ex, dt_config); }

                await PostResInfo(TablePO, ID, WechatKey);  //POST 执行后的结果


            }
        
        }

        public async Task PostWillDo(string PO, string PUser, string Supplier, string CompanyCode, string PayCode, string WechatKey)
        {
            if (WechatKey != "")
            {
                var requestBody = new
                {
                    msgtype = "markdown",
                    markdown = new
                    {
                        content = "预付订单即将开跑，采购单号为：<font color=\"warning\">" + PO + "</font>，请<font color=\"warning\">" + PUser + "</font>注意执行结果。\n" +
                              ">供应商:<font color=\"comment\">" + Supplier + "</font>\n" +
                              ">公司代码:<font color=\"comment\">" + CompanyCode + "</font>\n" +
                              ">类型:<font color=\"comment\">" + PayCode + "</font>\n" +                             
                              ">执行结果:<font color=\"comment\">正在执行</font>"
                    }
                };
                string jsonContent = JsonConvert.SerializeObject(requestBody);
                await publicClass.WechatPost(WechatKey, jsonContent);
            }

        }

        public async Task PostResInfo(string TablePO,string ID,string WechatKey)
        {
            if (WechatKey != "")
            {
                string sql = $"select * from {TablePO} where ID={ID}";
                DataTable  dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
                string PO = dt.Rows[0]["PO"].ToString();
                string PUser = dt.Rows[0]["PUser"].ToString();
                string Supplier = dt.Rows[0]["Supplier"].ToString();                
                string CompanyCode = dt.Rows[0]["CompanyCode"].ToString();                
                string PayCode = dt.Rows[0]["PayCode"].ToString();
                string ResInfo = dt.Rows[0]["ResInfo"].ToString().Trim();

                var requestBody = new
                {
                    msgtype = "markdown",
                    markdown = new
                    {
                        content = "预付订单执行完毕，采购单号为：<font color=\"warning\">" + PO + "</font>，请<font color=\"warning\">" + PUser + "</font>注意执行结果。\n" +
                              ">供应商:<font color=\"comment\">" + Supplier + "</font>\n" +
                              ">公司代码:<font color=\"comment\">" + CompanyCode + "</font>\n" +
                              ">类型:<font color=\"comment\">" + PayCode + "</font>\n" +
                              ">执行结果:<font color=\"comment\">"+( ResInfo==""?"异常，等下次重试": ResInfo) + "</font>"
                    }
                };
                string jsonContent = JsonConvert.SerializeObject(requestBody);
                await publicClass.WechatPost(WechatKey, jsonContent);
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

        private async Task SwitchToIframe(CancellationToken token, By by,int timeout = 60)
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

        private async Task SwitchToIframe(CancellationToken token, By by,string documentUrl,int timeout=60)
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


        public async Task LoginOA(CancellationToken token, DataTable dt_config)
        {
            try
            {
                string OaLoginUrl = await publicClass.GetLibValue(token, dt_config, "OaLoginUrl");
                string OaUser = await publicClass.GetLibValue(token, dt_config, "OaUser");
                string OaPassWorld = await publicClass.GetLibValue(token, dt_config, "OaPassWorld");
                if (OaLoginUrl == "" || OaUser == "" || OaPassWorld == "") { throw new Exception($"配置表中的OA相关信息为空字符"); }

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
            catch (Exception ex) { throw ex; }

        }      
        private async Task Nav_Liucheng(CancellationToken token)
        {
            try
            {
                bool isClick=false;
                wait60.Until(d => d.FindElement(By.XPath("//div[@class='slideItemText' and text()='流程']")).Displayed);
                var divs=MyDriver.FindElements(By.XPath("//div[@class='slideItemText' and text()='流程']"));
                foreach (IWebElement div in divs)
                {
                    var parents = div.FindElements(By.XPath(".."));
                    foreach (IWebElement parent in parents)
                    {
                        if (parent.GetDomAttribute("title") == "我的流程") { parent.Click(); isClick = true; return; }
                    }
                    await Task.Delay(0, token); token.ThrowIfCancellationRequested();
                }
                if (!isClick) { throw new Exception("单击<我的流程>异常"); }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private async Task New_Liucheng(CancellationToken token)
        {            
            try
            {
                bool isClick = false;
                wait60.Until(d => d.FindElement(By.Id("drillmenu")).Displayed);
                var drillmenu = MyDriver.FindElement(By.Id("drillmenu"));
                var liCss2s= drillmenu.FindElements(By.CssSelector("li.liCss2"));                        
                foreach (IWebElement li in liCss2s)
                {
                    var Mya = li.FindElements(By.TagName("a"));
                    foreach (IWebElement a in Mya)
                    {
                        if (a.GetDomAttribute("href").Contains("RequestType.jsp")) { isClick = true; a.Click(); return; }
                        await Task.Delay(0, token); token.ThrowIfCancellationRequested();
                    }
                    await Task.Delay(0, token); token.ThrowIfCancellationRequested();
                }
                if (!isClick) { throw new Exception("单击<新建流程>异常"); }
            }
            catch (Exception ex) { throw ex; }
        }
        private async Task ClickZF003(CancellationToken token,int timeout=60)
        {
            try
            {                              
                List<string> originalWindow = MyDriver.WindowHandles.ToList(); // 获取所有窗口句柄                


                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token,By.XPath("//iframe[@id='mainFrame']"));               
                await SwitchToIframe(token,By.XPath("//iframe[@id='tabcontentframe']"));

                //点击<ZF003-预付款申请流程>
                string xpathToFind = "//table[@class='ViewForm']//a[@class='e8contentover']";
                IWebElement myClick = null;
                var e8contentovers = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                for (int t=0;t<timeout;t++)
                {
                    foreach (IWebElement e8contentover in e8contentovers)
                    {
                        if (e8contentover.Text.ToUpper().Contains("ZF003-预付款申请流程")) { myClick = e8contentover; break; }                      
                    }
                    if (myClick != null) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    e8contentovers = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                    if (t + 1 == timeout) { throw new Exception("未能找到<ZF003-预付款申请流程>的链接"); }
                }
                myClick.Click();//点击

                //切换新窗口
                await SwitchToWindow(token, originalWindow, 60);

            }
            catch (Exception ex) { throw ex; }
        }            
        private async Task FillPUser(CancellationToken token,string PUser,int timeout=60)
        {
            try
            {
                
                await SwitchToDefaultContent(token);
                await SwitchToIframe(token, By.XPath("//iframe[@id='bodyiframe']"));

                //点击《申请人》搜索按钮
                var field91646_browserbtns = MyDriver.FindElements(By.Id("field91646_browserbtn"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91646_browserbtns.Count > 0) { field91646_browserbtns[0].Click(); break; }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91646_browserbtns = MyDriver.FindElements(By.Id("field91646_browserbtn"));
                    if (t + 1 == timeout) { throw new Exception("未能找到申请人搜索按钮"); }
                }

                //切换iframe
                MyDriver.SwitchTo().DefaultContent();       
                await SwitchToIframe(token, By.TagName("iframe"), "BrowserMain.jsp");
                await SwitchToIframe(token, By.Id("main"), "ResourceBrowser.jsp");
                await SwitchToIframe(token, By.Id("frame1"), "Select.jsp");

                //《申请人》搜索输入姓名
                var flowTitles = MyDriver.FindElements(By.XPath("//input[@id='flowTitle']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (flowTitles.Count > 0) 
                    {
                        flowTitles[0].Clear(); 
                        flowTitles[0].SendKeys(PUser);
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        flowTitles[0].SendKeys(OpenQA.Selenium.Keys.Enter);
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    flowTitles = MyDriver.FindElements(By.XPath("//input[@id='flowTitle']"));
                    if (t + 1 == timeout) { throw new Exception("未能找到申请人搜索输入框"); }
                }

                //选中第一个匹配项               
                string xpathToFind = "//div[@id='e8_box_source_quick']//table[@class='e8_box_source']//tbody//tr//td[@id='lastname']";
                var tds = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                for (int t = 0; t < timeout; t++)
                {
                    if (tds.Count > 0) { tds[0].Click(); break; }//仅第一个

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    tds = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                    if (t + 1 == timeout) { throw new Exception("申请人选中第一个匹配项异常"); }
                }

            }
            catch (Exception ex){ throw ex; }
        }
        private async Task FillSupplierCode(CancellationToken token, string supplierCode,int timeout=60)
        {
            try
            {
               
                MyDriver.SwitchTo().DefaultContent();               
                await SwitchToIframe(token, By.XPath("//iframe[@id='bodyiframe']"));

                //输入供应商代码
                var field91650s = MyDriver.FindElements(By.Id("field91650"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91650s.Count > 0) 
                    { 
                        field91650s[0].Clear();
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        field91650s[0].SendKeys(supplierCode);
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        field91650s[0].SendKeys(OpenQA.Selenium.Keys.Tab);
                        break; 
                    }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91650s = MyDriver.FindElements(By.Id("field91650"));
                    if (t + 1 == timeout) { throw new Exception("未能找到供应商代码输入框"); }
                }

            }
            catch (Exception ex) { throw ex; }

        }

        private async Task FillCurrency(CancellationToken token, string currency,int timeout=60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.XPath("//iframe[@id='bodyiframe']"));


                //点击《支付币种》搜索按钮
                var field91744_browserbtns = MyDriver.FindElements(By.Id("field91744_browserbtn"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91744_browserbtns.Count > 0) { field91744_browserbtns[0].Click();break; }
                   
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91744_browserbtns = MyDriver.FindElements(By.Id("field91744_browserbtn"));
                    if (t + 1 == timeout) { throw new Exception("未能找到支付币种搜索按钮"); }
                }

                //切换iframe
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.TagName("iframe"), "BrowserMain.jsp");
                await SwitchToIframe(token, By.Id("main"), "CommonBrowser.jsp");



                //输入币种
                var Currency_codes = MyDriver.FindElements(By.XPath("//input[@name='Currency_code']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (Currency_codes.Count > 0)
                    {
                        Currency_codes[0].Clear();
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        Currency_codes[0].SendKeys(currency);
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        Currency_codes[0].SendKeys(OpenQA.Selenium.Keys.Enter);
                        break;
                    }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    Currency_codes = MyDriver.FindElements(By.XPath("//input[@name='Currency_code']"));
                    if (t + 1 == timeout) { throw new Exception("未能找到币种输入框"); }
                }


                //选中第一个匹配项
                bool IsSelected = false;
                string xpathToFind = "//div[@id='_xTable']//table//tbody//tr//td";
                var tds = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                for (int t = 0; t < timeout; t++)
                {

                    foreach (IWebElement td in tds)
                    {
                        if (td.Text.Contains(currency) && !td.GetDomAttribute("style").Replace(" ", "").Contains("display:none"))
                        {
                            td.Click(); IsSelected = true; break;
                        }
                    }
                    if (IsSelected) { break; }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    tds = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                    if (t + 1 == timeout) { throw new Exception("币种选中第一个匹配项异常"); }
                }


            }
            catch (Exception ex) { throw ex; }
        }

        private async Task FillPayStype(CancellationToken token, string PayStype,int timeout=60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //选择付款类型
                var field91654s = MyDriver.FindElements(By.Id("field91654"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91654s.Count > 0)
                    {
                        var select = new SelectElement(field91654s[0]); //Selenium.Support 包 
                        select.SelectByText(PayStype);
                        break; 
                    }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91654s = MyDriver.FindElements(By.Id("field91654"));
                    if (t + 1 == timeout) { throw new Exception("未能找到付款类型的选项框"); }
                }

            }
            catch (Exception ex) { throw ex; }


        }

        private async Task FillClass2(CancellationToken token, string Class2,int timeout=60)
        {
            try
            {

                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");


                //选择二级分类
                var field91832s = MyDriver.FindElements(By.Id("field91832"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91832s.Count > 0)
                    {
                        var select = new SelectElement(field91832s[0]); //Selenium.Support 包 
                        select.SelectByText(Class2);
                        break;
                    }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91832s = MyDriver.FindElements(By.Id("field91832"));
                    if (t + 1 == timeout) { throw new Exception("未能找到二级分类的选项框"); }
                }

            }
            catch (Exception ex) { throw ex; }

        }

        private async Task FillSupplierShortName(CancellationToken token, string SupplierShortName, int timeout = 60)
        {
            try
            {

                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");


                //输入供应商简称
                var field91655s = MyDriver.FindElements(By.Id("field91655"));              
                for (int t = 0; t < timeout; t++)
                {
                    if (field91655s.Count > 0)
                    {
                        field91655s[0].Clear();
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        field91655s[0].SendKeys(SupplierShortName);
                        break; 
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91655s = MyDriver.FindElements(By.Id("field91655"));
                    if (t + 1 == timeout) { throw new Exception("未能找到供应商简称的输入框"); }
                }
                

            }
            catch (Exception ex) { throw ex; }


        }
                
        private async Task FillCompanyCode(CancellationToken token, string CompanyCode,int timeout=60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");


                //点击搜索按钮
                var field91697_browserbtns = MyDriver.FindElements(By.Id("field91697_browserbtn"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91697_browserbtns.Count > 0) { field91697_browserbtns[0].Click(); break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91697_browserbtns = MyDriver.FindElements(By.Id("field91697_browserbtn"));
                    if (t + 1 == timeout) { throw new Exception("未能找到公司代码的搜索按钮"); }
                }

                //切换iframe
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.TagName("iframe"), "systeminfo/BrowserMain.jsp");
                await SwitchToIframe(token, By.Id("main"), "interface/CommonBrowser.jsp");


                //输入公司代码
                var gongsdms = MyDriver.FindElements(By.XPath("//input[@name='gongsdm']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (gongsdms.Count > 0)
                    {
                        gongsdms[0].Clear();
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        gongsdms[0].SendKeys(CompanyCode);
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        gongsdms[0].SendKeys(OpenQA.Selenium.Keys.Enter);
                        break;
                    }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    gongsdms = MyDriver.FindElements(By.XPath("//input[@name='gongsdm']"));
                    if (t + 1 == timeout) { throw new Exception("未能找到公司代码的输入框"); }
                }





                //选中第一个匹配项               
                bool IsSelected = false;
                string xpathToFind = "//div[@id='_xTable']//table//tbody//tr//td";
                var tds = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                for (int t = 0; t < timeout; t++)
                {
                    foreach (IWebElement td in tds)
                    {
                        if (td.Text.Contains(CompanyCode) && !td.GetDomAttribute("style").Replace(" ", "").Contains("display:none"))
                        {
                            IsSelected = true;
                            td.Click(); 
                            break;
                        }
                    }
                    if (IsSelected) { break; }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    tds = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                    if (t + 1 == timeout) { throw new Exception($"公司代码{CompanyCode}选中第一个匹配项异常"); }
                }

             

            }
            catch (Exception ex) { throw ex; }

        }

        private async Task FillBankOf(CancellationToken token, string BankOf,int timeout=60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //选择对公对私
                var field91821s = MyDriver.FindElements(By.Id("field91821"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91821s.Count > 0) 
                    {
                        var select = new SelectElement(field91821s[0]); //Selenium.Support 
                        select.SelectByText(BankOf);
                        break; 
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91821s = MyDriver.FindElements(By.Id("field91821"));
                    if (t + 1 == timeout) { throw new Exception("未能找到对公对私的选项框"); }
                }

              
            }
            catch (Exception ex) { throw ex; }
        }

        private async Task FillInvoiceDate(CancellationToken token, string InvoiceDate,int timeout=60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //点击日期图标
                var field91822browsers = MyDriver.FindElements(By.Id("field91822browser"));
                for (int t=0;t<timeout;t++)
                {
                    if (field91822browsers.Count > 0)
                    {
                        field91822browsers[0].Click();
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91822browsers = MyDriver.FindElements(By.Id("field91822browser"));
                    if (t + 1 == timeout) { throw new Exception("未能找到日期图标"); }
                }

                //切换iframe
                await SwitchToIframe(token, By.TagName("iframe"), "My97DatePicker.htm");


                //输入年份 
                IWebElement YInput = null;
                var YMenuDivs = MyDriver.FindElements(By.XPath("//div[@id='dpTitle']//div[@class='menuSel YMenu']"));
                for (int t = 0; t < timeout; t++)
                {                    
                    if (YMenuDivs.Count > 0)
                    {
                        IWebElement YInput_parent = YMenuDivs[0].FindElement(By.XPath("./.."));
                        var YInputs = YInput_parent.FindElements(By.TagName("input"));
                        if (YInputs.Count > 0)
                        {
                            YInput=YInputs[0];
                            break;
                        }
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    YMenuDivs = MyDriver.FindElements(By.XPath("//div[@id='dpTitle']//div[@class='menuSel YMenu']"));
                    if (t + 1 == timeout) { throw new Exception("未能找到年份的输入框"); }
                }
                YInput.Click();
                for (int i = 0; i < 5; i++) { YInput.SendKeys(OpenQA.Selenium.Keys.Backspace); await Task.Delay(50, token); token.ThrowIfCancellationRequested(); }
                for (int i = 0; i < 5; i++) { YInput.SendKeys(OpenQA.Selenium.Keys.Delete); await Task.Delay(50, token); token.ThrowIfCancellationRequested(); }
                YInput.SendKeys(Convert.ToDateTime(InvoiceDate).ToString("yyyy"));



                //输入月份 
                IWebElement MInput = null;
                var MMenuDiv = MyDriver.FindElements(By.XPath("//div[@id='dpTitle']//div[@class='menuSel MMenu']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (MMenuDiv.Count > 0)
                    {
                        IWebElement MMenuDiv_parent = MMenuDiv[0].FindElement(By.XPath("./.."));
                        var MInputs = MMenuDiv_parent.FindElements(By.TagName("input"));
                        if (MInputs.Count > 0)
                        {
                            MInput = MInputs[0];
                            break;
                        }
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    MMenuDiv = MyDriver.FindElements(By.XPath("//div[@id='dpTitle']//div[@class='menuSel MMenu']"));
                    if (t + 1 == timeout) { throw new Exception("未能找到月份 的输入框"); }
                }
                MInput.Click();
                for (int i = 0; i < 5; i++) { MInput.SendKeys(OpenQA.Selenium.Keys.Backspace); await Task.Delay(50, token); token.ThrowIfCancellationRequested(); }
                for (int i = 0; i < 5; i++) { MInput.SendKeys(OpenQA.Selenium.Keys.Delete); await Task.Delay(50, token); token.ThrowIfCancellationRequested(); }
                MInput.SendKeys(Convert.ToDateTime(InvoiceDate).ToString("MM"));
                await Task.Delay(100, token); token.ThrowIfCancellationRequested();

                //回车
                MInput.SendKeys(OpenQA.Selenium.Keys.Enter);
                await Task.Delay(100, token); token.ThrowIfCancellationRequested();


                //切换iframe
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");


               
                //点击日期图标
                var field91822browsers2 = MyDriver.FindElements(By.Id("field91822browser"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91822browsers2.Count > 0)
                    {
                        field91822browsers2[0].Click();
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91822browsers2 = MyDriver.FindElements(By.Id("field91822browser"));
                    if (t + 1 == timeout) { throw new Exception("未能找到日期图标"); }
                }


                //切换iframe                
                await SwitchToIframe(token, By.TagName("iframe"), "My97DatePicker.htm");

                //选择日期>>点击对应的天
                int year = Convert.ToDateTime(InvoiceDate).Year;
                int month = Convert.ToDateTime(InvoiceDate).Month;
                int day = Convert.ToDateTime(InvoiceDate).Day;
                string xpath_dayClick = $"//table[@class='WdayTable']// td[@onclick='day_Click({year},{month},{day});']";
                var tds = MyDriver.FindElements(By.XPath($"{xpath_dayClick}"));
                for (int t = 0; t < timeout; t++)
                {
                    if (tds.Count > 0) { tds[0].Click();break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    tds = MyDriver.FindElements(By.XPath($"{xpath_dayClick}"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到点击对应的天{xpath_dayClick}"); }
                }


                //切换iframe
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");


                //校验日期
                var spanCK = MyDriver.FindElements(By.XPath("//span[@id='field91822span']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (spanCK.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    spanCK = MyDriver.FindElements(By.XPath("//span[@id='field91822span']"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到校验日期的spanCK"); }
                }
                var InputCK = MyDriver.FindElements(By.XPath("//Input[@id='field91822']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (InputCK.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    InputCK = MyDriver.FindElements(By.XPath("//Input[@id='field91822']"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到校验日期的InputCK"); }
                }              
                string str1 = Convert.ToDateTime(InvoiceDate).ToString("yyyy-MM-dd");
                string str2 = Convert.ToDateTime(spanCK[0].Text.Trim()).ToString("yyyy-MM-dd");
                string str3 = Convert.ToDateTime(InputCK[0].GetDomAttribute("value").Trim()).ToString("yyyy-MM-dd");
                if (!(str1 == str2 && str1 == str3))
                {
                    throw new Exception($"预计金税发票提供时间与实际选中日期不符,InvoiceDate={str1},span={str2},Input={str3}");
                }

            }
            catch (Exception ex) { throw ex; }

        }

        private async Task FillIsLinkPR(CancellationToken token, string isLinkPR,int timeout=60)
        {
            try
            {
                //切换iframe
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //选择是否关联资本性支出请购流程
                var field91809s = MyDriver.FindElements(By.Id("field91809"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91809s.Count > 0) 
                    {
                        var select = new SelectElement(field91809s[0]); //Selenium.Support 
                        select.SelectByText(isLinkPR);
                        break; 
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91809s = MyDriver.FindElements(By.Id("field91809"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到是否关联资本性支出请购流程的选项框"); }
                }                               
            }
            catch (Exception ex) { throw ex; }
        }

        private async Task FillPurchaseStype(CancellationToken token, string PurchaseStype,int timeout=60)
        {
            try
            {

                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //选择分期类型
                var field91653s = MyDriver.FindElements(By.Id("field91653"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91653s.Count > 0)
                    {
                        var select = new SelectElement(field91653s[0]); //Selenium.Support 
                        select.SelectByText(PurchaseStype);
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91653s = MyDriver.FindElements(By.Id("field91653"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到分期类型的选项框"); }
                }


                
            }
            catch (Exception ex) { throw ex; }

        }

        private async Task FillSWIFTCode(CancellationToken token, string SWIFTCode,int timeout=60)
        {
            try
            {

                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //点击搜索按钮
                var field91829_browserbtns = MyDriver.FindElements(By.Id("field91829_browserbtn"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91829_browserbtns.Count > 0) { field91829_browserbtns[0].Click(); break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91829_browserbtns = MyDriver.FindElements(By.Id("field91829_browserbtn"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到SWIFTCode的搜索按钮"); }
                }

                //切换iframe
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.TagName("iframe"), "systeminfo/BrowserMain.jsp");
                await SwitchToIframe(token, By.Id("main"), "interface/CommonBrowser.jsp");


                //在编码栏输入银行国家代码
                var guojdms = MyDriver.FindElements(By.XPath("//input[@name='guojdm']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (guojdms.Count > 0) 
                    {
                        guojdms[0].SendKeys(SWIFTCode);
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        guojdms[0].SendKeys(OpenQA.Selenium.Keys.Enter);
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        break; 
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    guojdms = MyDriver.FindElements(By.XPath("//input[@name='guojdm']"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到银行国家代码输入框"); }
                }


                //选中第一个匹配项
                bool IsSelected = false;
                string xpathToFind = "//div[@id='_xTable']//table//tbody//tr//td";
                var tds = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                for (int t = 0; t < timeout; t++)
                {
                    foreach (IWebElement td in tds)
                    {
                        if (td.Text.Contains(SWIFTCode) && !td.GetDomAttribute("style").Replace(" ", "").Contains("display:none"))
                        {
                            IsSelected = true;
                            td.Click();
                            break;
                        }
                    }
                    if (IsSelected) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    tds = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                    if (t + 1 == timeout) { throw new Exception($"银行国家代码{SWIFTCode}选中第一个匹配项异常"); }
                }

            }
            catch (Exception ex) { throw ex; }
        }

        private async Task FillIsInfrastructure(CancellationToken token, string IsInfrastructure, int timeout = 60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //选择是否基建类
                var field91807s = MyDriver.FindElements(By.Id("field91807"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91807s.Count > 0)
                    {
                        var select = new SelectElement(field91807s[0]); //Selenium.Support 
                        select.SelectByText(IsInfrastructure);
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91807s = MyDriver.FindElements(By.Id("field91807"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到是否基建类的选项框"); }
                }
              
            }
            catch (Exception ex) { throw ex; }

        }

        private async Task FillIsIsExistPO(CancellationToken token, string IsExistPO,int timeout=60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");


                //选择是否有采购订单号
                var field91707s = MyDriver.FindElements(By.Id("field91707"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91707s.Count > 0)
                    {
                        var select = new SelectElement(field91707s[0]); //Selenium.Support 
                        select.SelectByText(IsExistPO);
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91707s = MyDriver.FindElements(By.Id("field91707"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到是否有采购订单号的选项框"); }
                }
               
            }
            catch (Exception ex) { throw ex; }

        }

        private async Task FillPayFinish(CancellationToken token, string PayFinish,int timeout=60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //点击日期图标
                var field91764browsers = MyDriver.FindElements(By.Id("field91764browser"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91764browsers.Count > 0)
                    {
                        field91764browsers[0].Click();
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91764browsers = MyDriver.FindElements(By.Id("field91764browser"));
                    if (t + 1 == timeout) { throw new Exception("未能找到日期图标"); }
                }

                //切换iframe               
                await SwitchToIframe(token, By.TagName("iframe"), "My97DatePicker.htm");

                //输入年份 
                IWebElement YInput = null;
                var YMenuDivs = MyDriver.FindElements(By.XPath("//div[@id='dpTitle']//div[@class='menuSel YMenu']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (YMenuDivs.Count > 0)
                    {
                        IWebElement YInput_parent = YMenuDivs[0].FindElement(By.XPath("./.."));
                        var YInputs = YInput_parent.FindElements(By.TagName("input"));
                        if (YInputs.Count > 0)
                        {
                            YInput = YInputs[0];
                            break;
                        }
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    YMenuDivs = MyDriver.FindElements(By.XPath("//div[@id='dpTitle']//div[@class='menuSel YMenu']"));
                    if (t + 1 == timeout) { throw new Exception("未能找到年份的输入框"); }
                }
                YInput.Click();
                for (int i = 0; i < 5; i++) { YInput.SendKeys(OpenQA.Selenium.Keys.Backspace); await Task.Delay(50, token); token.ThrowIfCancellationRequested(); }
                for (int i = 0; i < 5; i++) { YInput.SendKeys(OpenQA.Selenium.Keys.Delete); await Task.Delay(50, token); token.ThrowIfCancellationRequested(); }
                YInput.SendKeys(Convert.ToDateTime(PayFinish).ToString("yyyy"));



                //输入月份 
                IWebElement MInput = null;
                var MMenuDiv = MyDriver.FindElements(By.XPath("//div[@id='dpTitle']//div[@class='menuSel MMenu']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (MMenuDiv.Count > 0)
                    {
                        IWebElement MMenuDiv_parent = MMenuDiv[0].FindElement(By.XPath("./.."));
                        var MInputs = MMenuDiv_parent.FindElements(By.TagName("input"));
                        if (MInputs.Count > 0)
                        {
                            MInput = MInputs[0];
                            break;
                        }
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    MMenuDiv = MyDriver.FindElements(By.XPath("//div[@id='dpTitle']//div[@class='menuSel MMenu']"));
                    if (t + 1 == timeout) { throw new Exception("未能找到月份 的输入框"); }
                }
                MInput.Click();
                for (int i = 0; i < 5; i++) { MInput.SendKeys(OpenQA.Selenium.Keys.Backspace); await Task.Delay(50, token); token.ThrowIfCancellationRequested(); }
                for (int i = 0; i < 5; i++) { MInput.SendKeys(OpenQA.Selenium.Keys.Delete); await Task.Delay(50, token); token.ThrowIfCancellationRequested(); }
                MInput.SendKeys(Convert.ToDateTime(PayFinish).ToString("MM"));
                await Task.Delay(100, token); token.ThrowIfCancellationRequested();

                //回车
                MInput.SendKeys(OpenQA.Selenium.Keys.Enter);
                await Task.Delay(100, token); token.ThrowIfCancellationRequested();


                //切换iframe
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");



                //点击日期图标2
                var field91764browsers2 = MyDriver.FindElements(By.Id("field91764browser"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91764browsers2.Count > 0)
                    {
                        field91764browsers2[0].Click();
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91764browsers2 = MyDriver.FindElements(By.Id("field91764browser"));
                    if (t + 1 == timeout) { throw new Exception("未能找到日期图标"); }
                }


                //切换iframe                
                await SwitchToIframe(token, By.TagName("iframe"), "My97DatePicker.htm");


                //选择日期>>点击对应的天
                int year = Convert.ToDateTime(PayFinish).Year;
                int month = Convert.ToDateTime(PayFinish).Month;
                int day = Convert.ToDateTime(PayFinish).Day;
                string xpath_dayClick = $"//table[@class='WdayTable']// td[@onclick='day_Click({year},{month},{day});']";
                var tds = MyDriver.FindElements(By.XPath($"{xpath_dayClick}"));
                for (int t = 0; t < timeout; t++)
                {
                    if (tds.Count > 0) { tds[0].Click(); break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    tds = MyDriver.FindElements(By.XPath($"{xpath_dayClick}"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到点击对应的天{xpath_dayClick}"); }
                }


                //切换iframe
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");


                //校验日期
                var spanCK = MyDriver.FindElements(By.XPath("//span[@id='field91764span']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (spanCK.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    spanCK = MyDriver.FindElements(By.XPath("//span[@id='field91764span']"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到校验日期的spanCK"); }
                }
                var InputCK = MyDriver.FindElements(By.XPath("//Input[@id='field91764']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (InputCK.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    InputCK = MyDriver.FindElements(By.XPath("//Input[@id='field91764']"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到校验日期的InputCK"); }
                }
                string str1 = Convert.ToDateTime(PayFinish).ToString("yyyy-MM-dd");
                string str2 = Convert.ToDateTime(spanCK[0].Text.Trim()).ToString("yyyy-MM-dd");
                string str3 = Convert.ToDateTime(InputCK[0].GetDomAttribute("value").Trim()).ToString("yyyy-MM-dd");
                if (!(str1 == str2 && str1 == str3))
                {
                    throw new Exception($"预计金税发票提供时间与实际选中日期不符,InvoiceDate={str1},span={str2},Input={str3}");
                }



            }
            catch (Exception ex) { throw ex; }

        }

        private async Task FillPO_NO(CancellationToken token, string PO, string ID, int timeout = 60)
        {
            try 
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //采购订单查询               
                var field93862s = MyDriver.FindElements(By.Id("field93862"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field93862s.Count > 0)
                    {
                        field93862s[0].Clear();
                        field93862s[0].SendKeys(PO);
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        field93862s[0].SendKeys(OpenQA.Selenium.Keys.Tab);
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field93862s = MyDriver.FindElements(By.Id("field93862"));
                    if (t + 1 == timeout) { throw new Exception("未能找到采购订单查询输入框"); }
                }
            }
            catch (Exception ex) { throw ex; }
        }
        private async Task FillPO_Detail(CancellationToken token, string PO, string ID, int timeout = 60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "ManageRequestNoFormIframe.jsp"); //保存后bodyiframe是ManageRequestNoFormIframe

                //等待加载：oTable0,oTable9
                string oTable0 = "//table[@id='oTable0']//tbody//tr[@_target='datarow']";
                var trs = MyDriver.FindElements(By.XPath($"{oTable0}"));
                for (int t = 0; t < timeout; t++)
                {
                    if (trs.Count > 0) { break; }
                    //-------------------------------------
                    /*加此两行是因为本函数上个方法是保存，保存DOM发生变更，可能未能切换缘故*/
                    MyDriver.SwitchTo().DefaultContent();
                    await SwitchToIframe(token, By.Id("bodyiframe"), "ManageRequestNoFormIframe.jsp"); //保存后bodyiframe是ManageRequestNoFormIframe
                    //-------------------------------------
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    trs = MyDriver.FindElements(By.XPath($"{oTable0}"));
                    if (t + 1 == timeout) { throw new Exception("datarow数据行的tr行未完全加载"); }
                }

                //等待加载：数据行第一行的 相应的列加载完毕
                for (int t = 0; t < timeout; t++)
                {
                    var w1 = trs[0].FindElements(By.XPath(".//input[@temptitle='订单本币金额（未税）']"));
                    var w2 = trs[0].FindElements(By.XPath(".//input[@temptitle='本次付款金额(未税)']"));
                    var w3 = trs[0].FindElements(By.XPath(".//input[@temptitle='税率']"));
                    var w4 = trs[0].FindElements(By.XPath(".//input[@temptitle='本次付款金额（含税）']"));
                    if (w1.Count > 0 && w2.Count > 0 && w3.Count > 0 && w4.Count > 0) { break; }
                    if (t + 1 == timeout) { throw new Exception("datarow数据行的tr行的相应的列未完全加载"); }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    trs = MyDriver.FindElements(By.XPath($"{oTable0}"));
                }



                //遍历datarow数据行，并设置每行的必填列
                for (int i = 0; i < trs.Count; i++)
                {
                    //获取：本币(未)                   
                    string benbiwei = trs[i].FindElement(By.XPath(".//input[@temptitle='订单本币金额（未税）']")).GetDomAttribute("value").Trim();
                    if (benbiwei == "" || benbiwei == "0") { throw new Exception("获取'订单本币金额（未税）'失败！"); }

                    //获取：税率 
                    string shuilv = trs[i].FindElement(By.XPath(".//input[@temptitle='税率']")).GetDomAttribute("value").Trim().Replace("%", "");
                    if (shuilv == "") { shuilv = "0"; }


                    //计算：本付款(含)
                    double benfuhan = Convert.ToDouble(benbiwei) + (Convert.ToDouble(benbiwei) * Convert.ToDouble(shuilv) / 100);
                    benfuhan = Math.Round(benfuhan, 2);


                    //先填写：本付款(含)     [需先清空] 
                    IWebElement input15 = trs[i].FindElement(By.XPath(".//input[@temptitle='本次付款金额（含税）']"));
                    input15.Clear();
                    input15.SendKeys(benfuhan.ToString());
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                    input15.SendKeys(OpenQA.Selenium.Keys.Tab);
                    await Task.Delay(500, token); token.ThrowIfCancellationRequested();


                    //后填写：本付款(未)     [需先清空]
                    IWebElement input11 = trs[i].FindElement(By.XPath(".//input[@temptitle='本次付款金额(未税)']"));
                    input11.Clear();
                    input11.SendKeys(benbiwei);
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                    input11.SendKeys(OpenQA.Selenium.Keys.Tab);
                    await Task.Delay(500, token); token.ThrowIfCancellationRequested();


                    //单据完整性
                    IWebElement select24 = trs[i].FindElement(By.XPath(".//select[@temptitle='单据完整性']"));
                    var selectE24 = new SelectElement(select24); //Selenium.Support 
                    selectE24.SelectByText("单据缺失");


                    // 备注原因 
                    IWebElement input26 = trs[i].FindElement(By.XPath(".//input[@temptitle='备注原因']"));
                    input26.SendKeys("先付款后补发票");

                }

            }
            catch (Exception ex) { throw ex; }
        }

              


        private async Task ResPayInfo_SaveBefore(CancellationToken token, string ID,int timeout = 60)
        {
            try 
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp"); //保存前


                //点击“支付信息”选项卡
                bool IsSelected = false;
                var spans = MyDriver.FindElements(By.XPath("//table[@class='excelMainTable tablefixed']//div[@class='tab_head']//span"));
                for (int t = 0; t < timeout; t++)
                {
                    foreach (var span in spans)
                    {
                        if (span.Text.Contains("支付信息")) { span.Click(); IsSelected = true; break; }
                    }
                    if (IsSelected) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    spans = MyDriver.FindElements(By.XPath("//table[@class='excelMainTable tablefixed']//div[@class='tab_head']//span"));
                    if (t + 1 == timeout) { throw new Exception("未能找到支付信息选项卡"); }
                }


                //点击“刷新支付金额”
                var btn_synzf_buttoms = MyDriver.FindElements(By.Id("btn_synzf_buttom"));
                for (int t = 0; t < timeout; t++)
                {
                    if (btn_synzf_buttoms.Count > 0) { btn_synzf_buttoms[0].Click(); break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    btn_synzf_buttoms = MyDriver.FindElements(By.Id("btn_synzf_buttom"));
                    if (t + 1 == timeout) { throw new Exception("未能找到刷新支付金额按钮"); }
                }

                //异常情况
                string MesssageE = "";
                try 
                {                   
                    MyDriver.SwitchTo().DefaultContent();
                   var Message_undefineds = MyDriver.FindElements(By.Id("Message_undefined"));
                    for (int t = 0; t < timeout; t++)
                    {
                        if (Message_undefineds.Count>0) { MesssageE ="<OA异常>"+ Message_undefineds[0].Text.Trim(); break; }
                        await Task.Delay(200, token); token.ThrowIfCancellationRequested();
                        Message_undefineds = MyDriver.FindElements(By.Id("Message_undefined"));                      
                    }
                }
                catch { }

                if(MesssageE!="")
                {
                    Connstr = await publicClass.GetConnstr(token, dt_config);
                    string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                    sqlClass.UpdSQL($"update {TablePO} set IsRequest=2 , ResInfo='{MesssageE}' where ID='{ID}'", Connstr);
                    throw new Exception($"{MesssageE}");
                }


            }
            catch (Exception ex) { throw ex; }
        }


        private async Task ResPayInfo_SaveAfter(CancellationToken token,int timeout=60)
        {
            try
            {                
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "ManageRequestNoFormIframe.jsp"); //保存后bodyiframe是ManageRequestNoFormIframe



                //点击“支付信息”选项卡
                bool IsSelected = false;
                var spans = MyDriver.FindElements(By.XPath("//table[@class='excelMainTable tablefixed']//div[@class='tab_head']//span"));
                for (int t = 0; t < timeout; t++)
                {
                    foreach (var span in spans)
                    {
                        if (span.Text.Contains("支付信息")) {  span.Click(); IsSelected = true; break; }
                    }
                    if (IsSelected) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    spans = MyDriver.FindElements(By.XPath("//table[@class='excelMainTable tablefixed']//div[@class='tab_head']//span"));
                    if (t + 1 == timeout) { throw new Exception("未能找到支付信息选项卡"); }
                }


                //点击“刷新支付金额”
                var btn_synzf_buttoms = MyDriver.FindElements(By.Id("btn_synzf_buttom"));
                for (int t = 0; t < timeout; t++)
                {
                    if (btn_synzf_buttoms.Count > 0) { btn_synzf_buttoms[0].Click();break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    btn_synzf_buttoms = MyDriver.FindElements(By.Id("btn_synzf_buttom"));
                    if (t + 1 == timeout) { throw new Exception("未能找到刷新支付金额按钮"); }
                }

                //等待几秒以便后续判断是否能刷新出来
                await Task.Delay(3000, token); token.ThrowIfCancellationRequested();




                //检查必填项
                /*
                 * 保存后再检查，保存后才能取到各项值
                 * 
                 * 
                 * 
                 */
            }
            catch (Exception ex) { throw ex; }

        }

        private async Task FillFilePath(CancellationToken token, string FilePath,int timeout=60)
        {
            try
            {
                //MyDriver.SwitchTo().DefaultContent();
                //await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "ManageRequestNoFormIframe.jsp"); //保存后bodyiframe是ManageRequestNoFormIframe



                if (!File.Exists(FilePath)) { throw new Exception("ZMMF007采购订单{PO}文件视乎还未下载哦！"); }

                //上传附件
                var spanButtonPlaceHolder91658s = MyDriver.FindElements(By.Id("spanButtonPlaceHolder91658"));
                for (int t = 0; t < timeout; t++)
                {
                    if (spanButtonPlaceHolder91658s.Count > 0)
                    {
                        IWebElement parent = spanButtonPlaceHolder91658s[0].FindElement(By.XPath("./..")); ;
                        IWebElement input = parent.FindElement(By.TagName("input"));
                        input.SendKeys(FilePath);
                        break;
                    }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    spanButtonPlaceHolder91658s = MyDriver.FindElements(By.Id("spanButtonPlaceHolder91658"));
                    if (t + 1 == timeout) { throw new Exception("未能找到上传附件按钮"); }                   

                }

                await Task.Delay(3000, token); token.ThrowIfCancellationRequested();

            }
            catch (Exception ex) { throw ex; }
        }

        private async Task SaveOA(CancellationToken token, string ID,int timeout=60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();


                //点击保存按钮（必填未填仍可以保存无错误弹窗）
                var inputSaves = MyDriver.FindElements(By.XPath("//div[@id='null_box']//input[@value='保存']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (inputSaves.Count > 0)
                    {
                        inputSaves[0].Click();
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    inputSaves = MyDriver.FindElements(By.XPath("//div[@id='null_box']//input[@value='保存']"));
                    if (t + 1 == timeout) { throw new Exception("未能找到保存按钮"); }
                }


                //以是否含有requestid来判断是否保存成功
                string requestid = "";
                for (int t = 0; t < timeout; t++)
                {
                    string url = MyDriver.Url.ToLower();
                    if (url.Contains("requestid"))
                    {
                        requestid = url.Substring(url.IndexOf("requestid=") + "requestid=".Length);
                        requestid = requestid.Substring(0, requestid.IndexOf("&"));
                        break;
                    }
                    await Task.Delay(3000, token); token.ThrowIfCancellationRequested();//3*timeout=180秒
                }
                if (requestid == "") { throw new Exception("OA点击保存按钮超过10分钟仍未成功"); }


                //存储requestid
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                sqlClass.UpdSQL($"update {TablePO} set RequestID='{requestid}',IsRequest=0 where ID='{ID}'", Connstr);


                //等待保存成功              
                for (int t = 0; t < timeout; t++)
                {
                    await Task.Delay(2000, token); token.ThrowIfCancellationRequested();
                    bool isEnabled = MyDriver.FindElement(By.XPath("//div[@id='null_box']//input[@value='保存']")).Enabled;
                    if (isEnabled) {break; }
                    if (t + 1 == timeout) { throw new Exception("保存按钮超时仍不可交互"); }
                }




            }
            catch (Exception ex) { throw ex; }
        }


        private async Task CheckPayInfo(CancellationToken token, string ID,int timeout=60)
        {
            //检查支付信息必填项-需先保存OA
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "ManageRequestNoFormIframe.jsp"); 

                //点击“支付信息”选项卡
                bool IsSelected = false;
                var spans = MyDriver.FindElements(By.XPath("//table[@class='excelMainTable tablefixed']//div[@class='tab_head']//span"));
                for (int t = 0; t < timeout; t++)
                {
                    foreach (var span in spans)
                    {
                        if (span.Text.Contains("支付信息")) { span.Click(); IsSelected = true; break; }
                    }
                    if (IsSelected) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();

                    //-------------------------------------
                    /*加此两行是因为本函数上个方法是保存，保存DOM发生变更，可能未能切换缘故*/
                    MyDriver.SwitchTo().DefaultContent();
                    await SwitchToIframe(token, By.Id("bodyiframe"), "ManageRequestNoFormIframe.jsp"); //保存后bodyiframe是ManageRequestNoFormIframe
                    //-------------------------------------
                    spans = MyDriver.FindElements(By.XPath("//table[@class='excelMainTable tablefixed']//div[@class='tab_head']//span"));
                    if (t + 1 == timeout) { throw new Exception("未能找到支付信息选项卡"); }
                }


                //检查必填项     >>[事先需OA保存成功]
                List<string> temptitles = new List<string>() { "金额", "供应商代码", "供应商名称", "收款人", "银行联行号", "账号" };
                var trs = MyDriver.FindElements(By.XPath("//table[@id='oTable4']//tr[@_target='datarow']"));//支付信息 默认只一行
                for (int t = 0; t < timeout; t++)
                {
                    if (trs.Count > 0) 
                    {                        
                        foreach (var temptitle in temptitles)
                        {
                            var intputs = trs[0].FindElements(By.XPath($"//input[@temptitle='{temptitle}']"));
                            if (!(intputs.Count > 0)) { throw new Exception($"未能找到{temptitle}的列[无等待机制的情况下]"); }
                            if (intputs[0].GetDomAttribute("value") == null || intputs[0].GetDomAttribute("value").ToString().Trim() == "")
                            {                               
                                Connstr = await publicClass.GetConnstr(token, dt_config);
                                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                                sqlClass.UpdSQL($"update {TablePO} set IsRequest=2 , ResInfo='<支付信息>选项卡中[{temptitle}]为空。' where ID='{ID}'", Connstr);
                                throw new Exception($"<支付信息>选项卡中[{temptitle}]为空, ID='{ID}'");
                               
                            }                               
                        }
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    trs = MyDriver.FindElements(By.XPath("//table[@id='oTable4']//tr[@_target='datarow']"));//支付信息 默认只一行
                    if (t + 1 == timeout) { throw new Exception("未能找到支付信息datarow行"); }
                }


            }
            catch (Exception ex) { throw ex; }
        }


        private async Task SubOA(CancellationToken token, string ID, int timeout = 60)
        {
            try
            {
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");

                MyDriver.SwitchTo().DefaultContent();
              
                int count1 = MyDriver.WindowHandles.Count;
                if (count1 < 2) { throw new Exception("此时应该有两个窗口才对"); }


                //定位“提交按钮”元素
                string xpathToFind = "//div[@id='null_box']//input[@value='提交']";
                var inputSaves = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                for (int t = 0; t < timeout; t++)
                {
                    if (inputSaves.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    inputSaves = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                    if (t + 1 == timeout) { throw new Exception("未能找到提交按钮"); }
                }

                //点击提交
                for (int t = 0; t < timeout; t++)
                {
                    await Task.Delay(2000, token); token.ThrowIfCancellationRequested();
                    bool isEnabled = MyDriver.FindElement(By.XPath($"{xpathToFind}")).Enabled;
                    if (isEnabled) { inputSaves[0].Click(); break; }
                    if (t + 1 == timeout) { throw new Exception("提交按钮超时仍不可交互"); }
                }


                //捕获提交后的异常，并设置IsRequest=2              
                try
                {
                    string msg = "";
                    var Message_undefineds = MyDriver.FindElements(By.Id("Message_undefined"));
                    for (int t = 0; t < 10; t++)
                    {
                        if (Message_undefineds.Count > 0)
                        {
                            msg = Message_undefineds[0].Text.Trim();
                            break;
                        }
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                        Message_undefineds = MyDriver.FindElements(By.Id("Message_undefined"));
                        if (t + 1 == 10) { /*不抛异常*/}
                    }
                    //设置IsRequest=2               
                    if (msg != "")
                    {
                        sqlClass.UpdSQL($"update {TablePO} set IsRequest=2,ResInfo='{msg}' where ID={ID} ", Connstr);
                        throw new Exception(msg);
                    }
                }
                catch { }
              

               




                //只点一次 不成功直接下次迭代。判断是否成功（窗口两个变为一个）成功break 到时失败下次迭代
                for (int t = 0; t <= 300; t++)
                {
                    int count2 = MyDriver.WindowHandles.Count;
                    if (count1 > count2)
                    {
                        sqlClass.UpdSQL($"update {TablePO} set IsRequest=1,ResInfo='成功' where ID={ID} ", Connstr);
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    if (t == 300) { throw new Exception("已经超过5分钟仍未提交成功"); }
                }




            }
            catch (Exception ex) { throw ex; }

        }





    }
    }
