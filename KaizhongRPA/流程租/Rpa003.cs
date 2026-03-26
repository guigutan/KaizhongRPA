using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using SAPFEWSELib;
using static Org.BouncyCastle.Math.EC.ECCurve;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;
using NPOI.Util;
using static System.Runtime.CompilerServices.RuntimeHelpers;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Net.NetworkInformation;
using System.Collections;
using OpenQA.Selenium.Chrome;
using System.Diagnostics.Eventing.Reader;
using System.Runtime.InteropServices;
using System.Reflection;
using Newtonsoft.Json;

namespace KaizhongRPA
{
    public class Rpa003:InvokeCenter
    {
        public RpaInfo GetThisRpaInfo()
        {
            RpaInfo rpaInfo = new RpaInfo();
            rpaInfo.RpaClassName = this.GetType().Name;
            rpaInfo.RpaName = "采购订单到付流程";
            rpaInfo.DefaultRunTime1 = "****-**-** **:**:**";
            rpaInfo.DefaultRunTime2 = "****-**-** **:**:**";
            rpaInfo.DefaultStatus = "有效";
            rpaInfo.DefaultPathStype = "相对路径";
            rpaInfo.DefaultConfigPath = @"config\RpaGroup\采购订单到付.xlsx";
            return rpaInfo;
        }
       
        #region 头部声明
        PublicClass publicClass = new PublicClass();
        SqlClass sqlClass = new SqlClass();
        public IWebDriver MyDriver;
        public WebDriverWait wait60;
        public string ScadaCompanyDir = $@"{MyPath.Documents}\{typeof(MyPath).Namespace}\ZF004\ScadaCompanyDir\";
        public string ScadaCompanyFN = "ScadaCompany.txt";

        public string PODir = $@"{MyPath.Documents}\{typeof(MyPath).Namespace}\ZF004\PODir\";

        public string PDFDir = $@"{MyPath.Documents}\{typeof(MyPath).Namespace}\ZF004\PDFDir\";

        public DataTable dt_config;
        public DataTable dt_Class2;
        public DataTable dt_PayTerm;
        public DataTable dt_ReplacePUser;
        public List<string> list_DAOFU;
        public DataTable dt_NeedScada;



        public string Connstr;
        #endregion
        public async Task RpaMain(CancellationToken token, RpaInfo rpaInfo)
        {
            try
            {
                if (!Directory.Exists(ScadaCompanyDir)) { Directory.CreateDirectory(ScadaCompanyDir); }
                if (!Directory.Exists(PODir)) { Directory.CreateDirectory(PODir); }
                if (!Directory.Exists(PDFDir)) { Directory.CreateDirectory(PDFDir); }
                
                dt_config = await publicClass.ExcelToDataTable(token, rpaInfo.DefaultConfigPath);
                if (!(dt_config != null && dt_config.Rows.Count > 0)) { return; }



                await DownCompanyCode(token);   //下载并保存公司代码至本地ScadaCompanyDir
                await ScadaCS(token);           //采集对账表checksheet

                await SetDT_RpaConfig(token); //本流程用到的配置文件

                await Fill_ZFIF001M(token);     //填充 采购单相关信息
                await Fill_ME23N(token);        //填充 采购组
                await Fill_ZMMR010(token);      //填充 银行国家代码
                await Fill_ZMMF007(token);      //填充 下载附件
                await Fill_ConverFinishDate(token); //填充，计算值：付款要求完成时间
                await Fill_Class2(token);           //填充，计算值：二级分类
                await IsScadaFinsh(token);  //检查 采集完毕
                await IsIntact(token);      //检查 填写完整
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

      

        #region 对账表主表采集相关
        private async Task DownCompanyCode(CancellationToken token)
        {

            await publicClass.DisableScreen(token);

            //抓取所有的公司代码
            DataTable dt = new DataTable();
            List<string> list = new List<string>();

            bool logined= await publicClass.GotoSapHome(token, dt_config, 0, 2);//登录SAP（已登录则退回首页）
            if (!logined) { throw new Exception("SAP未能登录成功"); }
            await EnterTransaction(token, "SE16");  //SE16事务码 

            (MySap.Session.FindById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME") as GuiCTextField).Text = "T001";          
            (MySap.Session.FindById("wnd[0]/tbar[0]/btn[0]") as GuiButton).Press(); //F7表内容(Enter)
            await Task.Delay(200, token); token.ThrowIfCancellationRequested();
            //默认 最大命中500

            (MySap.Session.FindById("wnd[0]/tbar[1]/btn[8]") as GuiButton).Press(); //执行   (F8)
            await Task.Delay(500, token); token.ThrowIfCancellationRequested();


            (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(33); //33 Ctrl+F9选择字段
            await Task.Delay(100, token); token.ThrowIfCancellationRequested();

            (MySap.Session.FindById("wnd[1]/tbar[0]/btn[14]") as GuiButton).Press(); //取消全选   (Shift+F2)

            GuiComponentCollection allColl = (MySap.Session.FindById("wnd[1]/usr/") as GuiUserArea).Children;
            foreach (GuiComponent myGuiComponent in allColl)
            {
                if (myGuiComponent is GuiLabel thisGui && thisGui.Text == "BUKRS") //仅选择BUKRS（公司代码）
                {
                    string ckhid = "wnd[1]/usr/chk[1"+ thisGui.Id.Substring(thisGui.Id.IndexOf(","));
                    (MySap.Session.FindById($"{ckhid}") as GuiCheckBox).Selected=true;
                    (MySap.Session.FindById("wnd[1]/tbar[0]/btn[6]") as GuiButton).Press(); //应用   (F6)
                    break;                   
                }
            }
            await Task.Delay(200, token); token.ThrowIfCancellationRequested();
            
            (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(20); //20 Shift+F8 下载
            await Task.Delay(200, token); token.ThrowIfCancellationRequested();
            await SelectedDownStyle(token); //选择下载的格式：含标签的文本
            await Task.Delay(200, token); token.ThrowIfCancellationRequested();
            (MySap.Session.FindById("wnd[1]/tbar[0]/btn[0]") as GuiButton).Press(); //继续
            await Task.Delay(500, token); token.ThrowIfCancellationRequested();
            await publicClass.ClearDir(token, ScadaCompanyDir); //先清空下载目录
            await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
            (MySap.Session.FindById("wnd[1]/usr/ctxtDY_PATH") as GuiCTextField).Text = ScadaCompanyDir; //填写下载目录
            (MySap.Session.FindById("wnd[1]/usr/ctxtDY_FILENAME") as GuiCTextField).Text = ScadaCompanyFN; //填写保存的文件名
            (MySap.Session.FindById("wnd[1]/tbar[0]/btn[11]") as GuiButton).Press(); //替换
            await Task.Delay(500, token); token.ThrowIfCancellationRequested();

            await  publicClass.Backspace(token,2);
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
                            finded = true;
                            string ID = thisGui.Id.ToString();//>>  /app/con[0]/ses[0]/wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]                           
                            (MySap.Session.FindById(ID.Substring(ID.IndexOf("wnd[1]/usr/"))) as GuiRadioButton).Selected = true;
                            break;
                        }
                    }
                    await Task.Delay(10, token); token.ThrowIfCancellationRequested();
                }
                if (!finded) { throw new Exception("未找到<含标签的文本>选项"); }
            }
            catch (Exception ex) {  throw ex; }
        }

        private async Task ScadaCS(CancellationToken token)
        {
            try
            {
                await publicClass.DisableScreen(token);

                //读取本地ScadaCompanyDir的公司代码
                List<string> CompanyCodeList = new List<string>();
                string path = ScadaCompanyDir + ScadaCompanyFN;
                if (!File.Exists(path)) { throw new Exception($"{path}文件不存在"); }
                try
                {
                    using (StreamReader reader = new StreamReader(path))
                    {
                        string line;
                        bool isStart = false;
                        while ((line = reader.ReadLine()) != null)
                        {
                            if (isStart && line.Trim() != "") { CompanyCodeList.Add(line.Trim()); } //置前
                            if (line.Contains("BUKRS")) { isStart = true; }//置后
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                if (CompanyCodeList.Count <= 0) { throw new Exception($"{path}获取所有的公司代码时为空。"); }
               
                //-------------------------------
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");               
                int beforeDay = Convert.ToInt32(await publicClass.GetLibValue(token, dt_config, "ScadaBeforeDay")); //对账表采集几天前(正数)               
                int PageDownMax = Convert.ToInt32(await publicClass.GetLibValue(token, dt_config, "PageDownMax")); //最多翻几页             
               
                
               
                foreach (string CompanyCode in CompanyCodeList)
                {
                    try
                    {
                        await publicClass.DisableScreen(token);

                        await InsertCS(token, CompanyCode, TableCS, beforeDay, PageDownMax);
                    }
                    catch (Exception ex) { await publicClass.NoteLog(token, ex, dt_config); }                   
                }

            }
            catch (Exception ex) { await publicClass.ExitSap(token); throw ex; }
            finally { await publicClass.ExitSap(token); }
        }

        private async Task InsertCS(CancellationToken token, string CompanyCode,string TableCS, int beforeDay,int PageDownMax)
        {
            try
            {
                await publicClass.GotoSapHome(token, dt_config, 0, 2);  //登录SAP（已登录则退回首页）
                await EnterTransaction(token, "ZFIF001M");              //ZFIF001M事务码

                (MySap.Session.FindById("wnd[0]/usr/ctxtP_BUKRS") as GuiCTextField).Text = CompanyCode; //填写公司代码                       
                (MySap.Session.FindById("wnd[0]/usr/ctxtS_CREDT-LOW") as GuiCTextField).Text = DateTime.Now.AddDays(beforeDay * (-1)).ToString("yyyy.MM.dd"); //填写创建日期                          
                (MySap.Session.FindById("wnd[0]/usr/ctxtS_CREDT-HIGH") as GuiCTextField).Text = DateTime.Now.ToString("yyyy.MM.dd");//填写创建日期       
                (MySap.Session.FindById("wnd[0]/tbar[1]/btn[8]") as GuiButton).Press();
                await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                try
                {
                    string text = (MySap.Session.FindById("wnd[0]/sbar/pane[0]") as GuiStatusPane).Text.Trim(); //执行前或后都是wnd[0]
                    if (text != "") { await publicClass.Backspace(token, 2); return; } //无数据，不再执行
                }
                catch { await publicClass.Backspace(token, 2); return; }

                //------------------------------------------------------
                //结果是单行还是多行。需等待more和one必有一个是真
                bool more = false, one = false;
                try { more = MySap.Session.FindById("wnd[0]/usr/tblZFIF001TAB_MOD") != null; } catch { }
                try { one = MySap.Session.FindById("wnd[0]/usr/cntlGRID1") != null; } catch { }
                for (int t = 0; t < 60; t++)
                {
                    if (more || one) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    try { more = MySap.Session.FindById("wnd[0]/usr/tblZFIF001TAB_MOD") != null; } catch { }
                    try { one = MySap.Session.FindById("wnd[0]/usr/cntlGRID1") != null; } catch { }
                }
                if (one && beforeDay < 365)
                {
                    await publicClass.Backspace(token, 2);

                    if (beforeDay <= 10) { await InsertCS(token, CompanyCode, TableCS, beforeDay + 1, PageDownMax); }
                    else if (beforeDay > 10 && beforeDay <= 20) { await InsertCS(token, CompanyCode, TableCS, beforeDay + 3, PageDownMax); }
                    else if (beforeDay > 20 && beforeDay <= 30) { await InsertCS(token, CompanyCode, TableCS, beforeDay + 5, PageDownMax); }
                    else { await InsertCS(token, CompanyCode, TableCS, beforeDay + 7, PageDownMax); }

                } //单行
                          
                
                //多行↓↓↓↓↓↓↓↓↓↓↓↓↓↓
                //翻页：
                int countPGDN = 0;//其他异常导致的死循环
                while (true)
                {
                    await publicClass.DisableScreen(token);

                    List<string> listZGZLN = new List<string>();
                    List<string> listNAME1 = new List<string>();
                    List<string> listLIFNR = new List<string>();
                    List<string> listBUKRS = new List<string>();
                    List<string> listFDATE = new List<string>();
                    List<string> listTDATE = new List<string>();
                    List<string> listNETWR = new List<string>();
                    List<string> listWMWST = new List<string>();
                    List<string> listWRBTR = new List<string>();
                    List<string> listZFPZT = new List<string>();
                    List<string> listWAERS = new List<string>();                   
                    GuiComponentCollection ZGZLN = (MySap.Session.FindById("wnd[0]/usr/tblZFIF001TAB_MOD") as GuiTableControl).FindAllByName("GW_GZBH-ZGZLN", "GuiTextField");//对账表编号
                    GuiComponentCollection NAME1 = (MySap.Session.FindById("wnd[0]/usr/tblZFIF001TAB_MOD") as GuiTableControl).FindAllByName("GW_GZBH-NAME1", "GuiTextField");//供应商名称
                    GuiComponentCollection LIFNR = (MySap.Session.FindById("wnd[0]/usr/tblZFIF001TAB_MOD") as GuiTableControl).FindAllByName("GW_GZBH-LIFNR", "GuiTextField");//供应商代码
                    GuiComponentCollection BUKRS = (MySap.Session.FindById("wnd[0]/usr/tblZFIF001TAB_MOD") as GuiTableControl).FindAllByName("GW_GZBH-BUKRS", "GuiTextField");//公司代码
                    GuiComponentCollection FDATE = (MySap.Session.FindById("wnd[0]/usr/tblZFIF001TAB_MOD") as GuiTableControl).FindAllByName("GW_GZBH-FDATE", "GuiTextField");//从日期
                    GuiComponentCollection TDATE = (MySap.Session.FindById("wnd[0]/usr/tblZFIF001TAB_MOD") as GuiTableControl).FindAllByName("GW_GZBH-TDATE", "GuiTextField");//到日期
                    GuiComponentCollection NETWR = (MySap.Session.FindById("wnd[0]/usr/tblZFIF001TAB_MOD") as GuiTableControl).FindAllByName("GW_GZBH-NETWR", "GuiTextField");//净价
                    GuiComponentCollection WMWST = (MySap.Session.FindById("wnd[0]/usr/tblZFIF001TAB_MOD") as GuiTableControl).FindAllByName("GW_GZBH-WMWST", "GuiTextField");//税额
                    GuiComponentCollection WRBTR = (MySap.Session.FindById("wnd[0]/usr/tblZFIF001TAB_MOD") as GuiTableControl).FindAllByName("GW_GZBH-WRBTR", "GuiTextField");//含税金额
                    GuiComponentCollection ZFPZT = (MySap.Session.FindById("wnd[0]/usr/tblZFIF001TAB_MOD") as GuiTableControl).FindAllByName("GW_GZBH-ZFPZT", "GuiComboBox");//发票状态
                    GuiComponentCollection WAERS = (MySap.Session.FindById("wnd[0]/usr/tblZFIF001TAB_MOD") as GuiTableControl).FindAllByName("GW_GZBH-WAERS", "GuiTextField");//币种
                    foreach (GuiComponent myGuiComponent in ZGZLN) { if (myGuiComponent is GuiTextField thisGui) { listZGZLN.Add(thisGui.Text.Trim()); } }
                    foreach (GuiComponent myGuiComponent in NAME1) { if (myGuiComponent is GuiTextField thisGui) { listNAME1.Add(thisGui.Text.Trim()); } }
                    foreach (GuiComponent myGuiComponent in LIFNR) { if (myGuiComponent is GuiTextField thisGui) { listLIFNR.Add(thisGui.Text.Trim()); } }
                    foreach (GuiComponent myGuiComponent in BUKRS) { if (myGuiComponent is GuiTextField thisGui) { listBUKRS.Add(thisGui.Text.Trim()); } }
                    foreach (GuiComponent myGuiComponent in FDATE) { if (myGuiComponent is GuiTextField thisGui) { listFDATE.Add(thisGui.Text.Trim()); } }
                    foreach (GuiComponent myGuiComponent in TDATE) { if (myGuiComponent is GuiTextField thisGui) { listTDATE.Add(thisGui.Text.Trim()); } }
                    foreach (GuiComponent myGuiComponent in NETWR) { if (myGuiComponent is GuiTextField thisGui) { listNETWR.Add(thisGui.Text.Trim()); } }
                    foreach (GuiComponent myGuiComponent in WMWST) { if (myGuiComponent is GuiTextField thisGui) { listWMWST.Add(thisGui.Text.Trim()); } }
                    foreach (GuiComponent myGuiComponent in WRBTR) { if (myGuiComponent is GuiTextField thisGui) { listWRBTR.Add(thisGui.Text.Trim()); } }
                    foreach (GuiComponent myGuiComponent in ZFPZT) { if (myGuiComponent is GuiComboBox thisGui) { listZFPZT.Add(thisGui.Text.Trim()); } }
                    foreach (GuiComponent myGuiComponent in WAERS) { if (myGuiComponent is GuiTextField thisGui) { listWAERS.Add(thisGui.Text.Trim()); } }

                    int rows = listZGZLN.Count;
                    if (rows == listNAME1.Count && rows == listLIFNR.Count && rows == listBUKRS.Count && rows == listFDATE.Count && rows == listTDATE.Count && rows == listNETWR.Count && rows == listWMWST.Count && rows == listWRBTR.Count && rows == listZFPZT.Count && rows == listWAERS.Count)
                    {
                        for (int i = 0; i < rows; i++)
                        {
                            string CSNO = listZGZLN[i];
                            string Supplier = listNAME1[i];
                            string SupplierCode = listLIFNR[i];
                            string SupplierShortName = SupplierCode;
                            //string CompanyCode = list[i];
                            string SDate = listFDATE[i];
                            string EDate = listTDATE[i];
                            string Price = listNETWR[i];
                            string TaxAmount = listWMWST[i];
                            string PriceIncludingTax = listWRBTR[i];
                            string InvoiceStatus = listZFPZT[i];
                            string Currency = listWAERS[i];                           
                            if (CSNO == "" || !InvoiceStatus.Contains("新建")) { continue; }

                            try { SDate = Convert.ToDateTime(SDate).ToString("yyyy-MM-dd"); } catch { SDate = "1900-01-01"; }
                            try { EDate = Convert.ToDateTime(EDate).ToString("yyyy-MM-dd"); } catch { EDate = "1900-01-01"; }
                            try { Price=Convert.ToDouble(Price).ToString(); } catch { Price = "0"; }
                            try { TaxAmount=Convert.ToDouble(TaxAmount).ToString(); } catch { TaxAmount = "0"; }
                            try { PriceIncludingTax=Convert.ToDouble(PriceIncludingTax).ToString(); } catch { PriceIncludingTax = "0"; }

                            string sql = $"insert into {TableCS}(CSNO,Supplier,SupplierCode,SupplierShortName,CompanyCode,SDate,EDate,Price,TaxAmount,PriceIncludingTax,InvoiceStatus,Currency) ";
                            sql += $" values('{CSNO}','{Supplier}','{SupplierCode}','{SupplierShortName}','{CompanyCode}','{SDate}','{EDate}','{Price}','{TaxAmount}','{PriceIncludingTax}','{InvoiceStatus}','{Currency}') ";
                            sqlClass.InsSQL(sql, Connstr);
                        }
                    }

                    //翻页
                    if (rows < 5 || countPGDN >= PageDownMax) { break; } //不再翻页
                    await publicClass.PGDN_CSNO(token);
                    countPGDN += 1;
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();

                }


            }
            catch (Exception ex) {  throw ex; }
           // finally { await publicClass.ExitSap(token); }
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
            catch (Exception ex) {  throw ex; }
        }

        #endregion

        #region ZFIF001M填充明细表>>采购订单，采购员描述，订单付款条件，订单付款条件描述，过账日期，采购总价，税率，税额，含税金额

        private async Task SetDT_RpaConfig(CancellationToken token)
        {
            await publicClass.DisableScreen(token);

            string pathClass2 = await publicClass.GetLibValue(token, dt_config, "Class2");
            string pathPayTerm = await publicClass.GetLibValue(token, dt_config, "PayTerm");
            string pathReplacePUser = await publicClass.GetLibValue(token, dt_config, "ReplacePUser");
            if (File.Exists(pathClass2)) { dt_Class2 = await publicClass.ExcelToDataTable(token, pathClass2); }
            if (File.Exists(pathPayTerm)) { dt_PayTerm = await publicClass.ExcelToDataTable(token, pathPayTerm); }
            if (File.Exists(pathReplacePUser)) { dt_ReplacePUser = await publicClass.ExcelToDataTable(token, pathReplacePUser); }

            list_DAOFU = new List<string>();
            if (dt_PayTerm != null && dt_PayTerm.Rows.Count > 0)
            {
                for (int i = 0; i < dt_PayTerm.Rows.Count; i++)
                {
                    if (dt_PayTerm.Rows[i]["对应流程"].ToString() == "到付")
                    {
                        list_DAOFU.Add(dt_PayTerm.Rows[i]["代码"].ToString());
                    }
                }
            }


        }
        private async Task<DataTable> GetDT_NeedScada(CancellationToken token)
        {
            Connstr = await publicClass.GetConnstr(token, dt_config);
            string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
            string sql = $"select * from {TableCS} where IsScadaFinsh=0 and  IsRequest<=0";
            return  sqlClass.SlcSQL(sql, Connstr);
        }


        private async Task Fill_ZFIF001M(CancellationToken token)
        {
            try
            {
                DataTable dt = await GetDT_NeedScada(token);
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
                if (list_DAOFU.Count <= 0) { throw new Exception("[订单付款条件代码配置表]中的“到付”代码为0行"); }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try 
                    {
                        await publicClass.DisableScreen(token);

                        string CSID = dt.Rows[i]["CSID"].ToString();
                        string CSNO = dt.Rows[i]["CSNO"].ToString();
                        string CompanyCode = dt.Rows[i]["CompanyCode"].ToString();
                        if (CompanyCode == "") { continue; }

                        await publicClass.GotoSapHome(token, dt_config, 0, 2);  //登录SAP（已登录则退回首页）
                        await EnterTransaction(token, "ZFIF001M");              //ZFIF001M事务码

                        (MySap.Session.FindById("wnd[0]/usr/ctxtP_BUKRS") as GuiCTextField).Text = CompanyCode; //填写公司代码                                                 
                        (MySap.Session.FindById("wnd[0]/usr/txtS_ZGZLN-LOW") as GuiTextField).Text = CSNO;//填写对账表编号       
                        (MySap.Session.FindById("wnd[0]/tbar[1]/btn[8]") as GuiButton).Press();
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();

                        string paneText= (MySap.Session.FindById("wnd[0]/sbar/pane[0]") as GuiStatusPane).Text.Trim();                      
                        if (paneText != "") { throw new Exception(paneText); }

                        bool chooselayout = await publicClass.ChooseLayout(token, "ASCADAPO", false, 20);//选择布局：采集专用(RPA)
                        if (!chooselayout) { throw new Exception("选择ASCADAPO布局错误"); }

                        (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(46);//Ctrl+Shift+F10打印预览
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();

                        (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(43);//43 Ctrl+Shift+F7 电子表格
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();

                        (MySap.Session.FindById("wnd[1]/usr/radRB_OTHERS") as GuiRadioButton).Selected = true;//从所有可用格式中选择
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        (MySap.Session.FindById("wnd[1]/tbar[0]/btn[0]") as GuiButton).Press();//继续

                        await Task.Delay(500, token); token.ThrowIfCancellationRequested();
                        (MySap.Session.FindById("wnd[1]/usr/ctxtDY_PATH") as GuiCTextField).Text = PODir;//填写目录
                        (MySap.Session.FindById("wnd[1]/usr/ctxtDY_FILENAME") as GuiCTextField).Text = $"{CSNO}.XLSX";//填写目录
                        (MySap.Session.FindById("wnd[1]/tbar[0]/btn[11]") as GuiButton).Press();//替换
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                        await publicClass.CloseExcle(token);//关闭Excle
                        await publicClass.Backspace(token,2);



                        string pathPO= PODir+ $"{CSNO}.XLSX";
                        DataTable dt_PO= await publicClass.ExcelToDataTable(token, pathPO);
                        if (dt_PO == null || dt_PO.Rows.Count <= 0) { continue; }
                        sqlClass.DelSQL($"delete from {TablePO} where  CSID={CSID} ", Connstr);

                        bool isNullAllPO = true;
                        string sqlPO = $"insert into  {TablePO} (CSID,PONO,PUser,PayCode,PayCodeExplanation,PostingDate,Price,TaxRate,TaxAmount,PriceIncludingTax) values ";
                        for (int p = 0; p < dt_PO.Rows.Count; p++)
                        {
                            await publicClass.DisableScreen(token);

                            string PO = dt_PO.Rows[p]["采购订单"].ToString();
                            if (PO != "")
                            {
                                isNullAllPO=false;
                                string PUser = dt_PO.Rows[p]["采购员描述"].ToString();
                                string PayCode = dt_PO.Rows[p]["订单付款条件"].ToString();
                                string PayCodeExplanation = dt_PO.Rows[p]["订单付款条件描述"].ToString();
                                string PostingDate = dt_PO.Rows[p]["过账日期"].ToString();
                             
                                string Price = dt_PO.Rows[p]["采购总价"].ToString();
                                string TaxRate = dt_PO.Rows[p]["税率"].ToString();
                                string TaxAmount = dt_PO.Rows[p]["税额"].ToString();
                                string PriceIncludingTax = dt_PO.Rows[p]["含税金额"].ToString();

                                if (!list_DAOFU.Contains($"{PayCode}")) { sqlClass.UpdSQL($"update {TableCS} set IsRequest=2,ResInfo='非到付单[{PayCode}{PayCodeExplanation}]' where CSID={CSID}", Connstr);break; }

                                try {
                                    string format = "dd-MMM-yyyy";
                                    CultureInfo culture = new CultureInfo("zh-CN"); // 使用中文文化信息
                                    PostingDate = DateTime.ParseExact(PostingDate, format, culture).ToString("yyyy-MM-dd");                                   
                                } catch { PostingDate = "1900-01-01"; }
                                try { Price = Convert.ToDouble(Price).ToString(); } catch { Price = "0"; }
                                try { TaxRate = Convert.ToDouble(TaxRate.Replace(",","")).ToString(); } catch { TaxRate = "0"; }
                                try { TaxAmount = Convert.ToDouble(TaxAmount.Replace(",", "")).ToString(); } catch { TaxAmount = "0"; }
                                try { PriceIncludingTax = Convert.ToDouble(PriceIncludingTax.Replace(",", "")).ToString(); } catch { PriceIncludingTax = "0"; }
                                 sqlPO += $"('{CSID}','{PO}','{PUser}','{PayCode}','{PayCodeExplanation}','{PostingDate}','{Price}','{TaxRate}','{TaxAmount}','{PriceIncludingTax}'),";                               
                            }                           
                        }
                        if (isNullAllPO) { sqlClass.UpdSQL($"update {TableCS} set IsRequest=2,ResInfo='该对账表的SAP的采购订单号为空' where CSID={CSID}", Connstr); }
                        else { sqlClass.InsSQL(sqlPO.Substring(0, sqlPO.Length-1), Connstr); }

                    }
                    catch (Exception ex) { string ss = ex.Message; await publicClass.ExitSap(token); }
                }
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



        #endregion

        #region ME23N填充明细表>>采购组
        private async Task Fill_ME23N(CancellationToken token)
        {
            try
            {
               
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                string sql = $"select DISTINCT {TablePO}.CSID, {TablePO}.PONO from {TablePO} left join {TableCS}  on {TableCS}.CSID={TablePO}.CSID where IsScadaFinsh=0 and IsRequest<=0 and PGroup is null";
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }              
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try 
                    {
                        await publicClass.DisableScreen(token);

                        string CSID = dt.Rows[i]["CSID"].ToString();
                        string PONO = dt.Rows[i]["PONO"].ToString();
                        await publicClass.GotoSapHome(token, dt_config, 0, 2);  //登录SAP（已登录则退回首页）
                        await EnterTransaction(token, "ME23N");              //ZFIF001M事务码

                        (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(17); //Shift+F5 其他采购订单
                        await Task.Delay(500, token); token.ThrowIfCancellationRequested();

                        ((MySap.Session.FindById("wnd[1]/usr/") as GuiUserArea).FindAllByName("MEPO_SELECT-BSTYP_F", "GuiRadioButton").Item(0) as GuiRadioButton).Selected = true;                     
                        ((MySap.Session.FindById("wnd[1]/usr/") as GuiUserArea).FindAllByName("MEPO_SELECT-EBELN", "GuiCTextField").Item(0) as GuiCTextField).Text = PONO;                       
                        (MySap.Session.FindById("wnd[1]") as GuiModalWindow).SendVKey(0);
                        await Task.Delay(500, token); token.ThrowIfCancellationRequested();

                        (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(26);//26 Ctrl+F2扩展抬头
                        await Task.Delay(500, token); token.ThrowIfCancellationRequested();
                        ((MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).FindAllByName("TABHDT8", "GuiTab").Item(0) as GuiTab).Select();
                        try
                        {
                            string thisPO = ((MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).FindAllByName("MEPO_TOPLINE-EBELN", "GuiTextField").Item(0) as GuiTextField).Text.Trim();
                            if (!thisPO.Contains(PONO)) { throw new Exception($"[{thisPO}]vs[{PONO}]Shift+F5查询的订单与展示的订单不一致，请检查。"); }

                            string PGroup = ((MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).FindAllByName("MEPO1222-EKGRP", "GuiCTextField").Item(0) as GuiCTextField).Text.Trim();
                          
                            sqlClass.UpdSQL($"update {TablePO} set PGroup='{PGroup}' where CSID='{CSID}' and PONO={PONO} ", Connstr);                          
                          
                        }
                        catch (Exception ex) {  throw ex; }

                      await  publicClass.Backspace(token,2);

                    }
                    catch (Exception ex) { await publicClass.ExitSap(token); await publicClass.NoteLog(token, ex, dt_config); }
                }

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

        #endregion

        # region ZMMR010填充明细表>>收款人国家/地区--银行国家代码SWIFTCode
        private async Task Fill_ZMMR010(CancellationToken token)
        {
            DataTable dt = await GetDT_NeedScada(token);
            Connstr = await publicClass.GetConnstr(token, dt_config);
            string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
            string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
            if (!(dt != null && dt.Rows.Count > 0)) { return; }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                try
                {
                    await publicClass.DisableScreen(token);

                    string CSID = dt.Rows[i]["CSID"].ToString();
                    string SupplierCode = dt.Rows[i]["SupplierCode"].ToString();
                    string CompanyCode = dt.Rows[i]["CompanyCode"].ToString();

                    await publicClass.GotoSapHome(token, dt_config, 0, 2);  //登录SAP（已登录则退回首页）
                    await EnterTransaction(token, "ZMMR010");              //事务码

                    (MySap.Session.FindById("wnd[0]/usr/ctxtS_LIFNR-LOW") as GuiCTextField).Text = SupplierCode;
                    (MySap.Session.FindById("wnd[0]/usr/ctxtS_BUKRS-LOW") as GuiCTextField).Text = CompanyCode;
                    (MySap.Session.FindById("wnd[0]/usr/radP_R3") as GuiRadioButton).Selected = true;
                    (MySap.Session.FindById("wnd[0]/tbar[1]/btn[8]") as GuiButton).Press();
                    await Task.Delay(500, token); token.ThrowIfCancellationRequested();

                    bool chooselayout = await publicClass.ChooseLayout(token, "001SCODE", false, 20);//选择布局
                    if (!chooselayout) { throw new Exception("选择001SCODE布局错误"); }
                    (MySap.Session.FindById("wnd[0]") as GuiMainWindow).SendVKey(46);//Ctrl+Shift+F10打印预览
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();

                    string IDLike = "";
                    GuiComponentCollection allColl = (MySap.Session.FindById("wnd[0]/usr/") as GuiUserArea).Children;
                    foreach (GuiComponent myGuiComponent in allColl)
                    {
                        if (myGuiComponent is GuiLabel thisGui && thisGui.Text.Contains("银行国家代码"))
                        {
                            IDLike = thisGui.Id;
                            IDLike = IDLike.Substring(IDLike.IndexOf("wnd[0]/usr/lbl["));
                            IDLike = IDLike.Substring(0, IDLike.IndexOf(",")+1);
                            break;
                        }
                    }
                    if (IDLike != "")
                    {
                        foreach (GuiComponent myGuiComponent in allColl)
                        {
                            if (myGuiComponent is GuiLabel thisGui && thisGui.Id.Contains($"{IDLike}"))
                            { 
                                string SWIFTCode= thisGui.Text.Trim();
                                if (SWIFTCode != ""&& Regex.IsMatch(SWIFTCode, @"^[a-zA-Z]+$")) { sqlClass.UpdSQL($"update {TablePO} set SWIFTCode='{SWIFTCode}' where CSID='{CSID}'", Connstr); break; }
                            }
                        }
                    }




                }
                catch (Exception ex) { await publicClass.ExitSap(token); await publicClass.NoteLog(token, ex, dt_config); }
               
            }
        }
        #endregion

        #region Fill_ZMMF007 下载附件
        private async Task Fill_ZMMF007(CancellationToken token)
        {
            try
            {
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                string sql = $"select DISTINCT {TablePO}.CSID from {TablePO} left join {TableCS}  on {TableCS}.CSID={TablePO}.CSID where IsScadaFinsh=0 and IsRequest<=0 and FilePath is null";
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        await publicClass.DisableScreen(token);

                        string CSID = dt.Rows[i]["CSID"].ToString();

                        string sqlPO = $"select DISTINCT PONO from {TablePO} where CSID={CSID}";
                        DataTable dtPO = sqlClass.SlcSQL(sqlPO, Connstr);
                        if (!(dtPO != null && dtPO.Rows.Count > 0)) { break; }
                        string listPO = "";
                        for (int p = 0; p < dtPO.Rows.Count; p++) { listPO += $"{dtPO.Rows[p]["PONO"].ToString()}\r\n"; }
                        await PublicClass.SetClipboardTextAsync(listPO);  //设置剪贴板
                        await Task.Delay(500, token); token.ThrowIfCancellationRequested();
                        string temp = await PublicClass.GetClipboardTextAsync();//获取剪贴板
                        if (listPO == temp)
                        {
                            await publicClass.GotoSapHome(token, dt_config, 0, 3);//登录SAP（已登录则退回首页）
                            await EnterTransaction(token, "ZMMF007"); //事务ZMMF007 >下载PDF
                            (MySap.Session.FindById("wnd[0]/usr/ctxtS_WERKS-LOW") as GuiCTextField).Text = "*";//输入工厂
                            (MySap.Session.FindById("wnd[0]/usr/btn%_S_EBELN_%_APP_%-VALU_PUSH") as GuiButton).Press();//单击采购单号的多项选择按钮
                            await Task.Delay(500, token); token.ThrowIfCancellationRequested();
                            (MySap.Session.FindById("wnd[1]/tbar[0]/btn[24]") as GuiButton).Press();//自剪贴板上载   (Shift+F12)
                            (MySap.Session.FindById("wnd[1]/tbar[0]/btn[0]") as GuiButton).Press(); //检查条目
                            await Task.Delay(500, token); token.ThrowIfCancellationRequested();
                            (MySap.Session.FindById("wnd[1]/tbar[0]/btn[12]") as GuiButton).Press();//取消
                            try { (MySap.Session.FindById("wnd[2]/usr/btnSPOP-OPTION1") as GuiButton).Press(); } catch { } //是否复制变更》是

                             (MySap.Session.FindById("wnd[0]/usr/radP4") as GuiRadioButton).Selected = true;//PDF导出本地
                            (MySap.Session.FindById("wnd[0]/tbar[1]/btn[8]") as GuiButton).Press();//F8执行
                            await Task.Delay(1000, token); token.ThrowIfCancellationRequested();

                            string PathDefault = (MySap.Session.FindById("wnd[0]/sbar/pane[0]") as GuiStatusPane).Text;//下载 338 KB D:\20240819-3100039872.PDF
                            PathDefault = PathDefault.Substring(PathDefault.IndexOf(@"D:\")).Trim();
                            string PathPdf = PDFDir + $"{Path.GetFileName(PathDefault)}";
                            if (File.Exists(PathDefault)) { File.Copy(PathDefault, PathPdf,true); /* PathPdf 覆盖同名*/}
                            if (File.Exists(PathPdf)) { sqlClass.UpdSQL($"update {TablePO} set FilePath='{PathPdf}' where CSID='{CSID}' ", Connstr); File.Delete(PathDefault); }

                        }


                    }
                    catch (Exception ex) { await publicClass.NoteLog(token, ex, dt_config); }


                }
            }
            catch (Exception ex) { throw ex; }
            finally { await publicClass.ExitSap(token); }

        }

        #endregion

        #region 检查 采集和填写
        private async Task IsScadaFinsh(CancellationToken token)
        {
            try
            {

                List<string> list1 = new List<string>() { "PONO", "PUser", "PayCode", "PayCodeExplanation", "PostingDate", "Price", "TaxRate", "TaxAmount", "PriceIncludingTax", "PGroup", "SWIFTCode", "FilePath" };
                List<string> list = list1.Distinct().ToList();


                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");                
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");                
               
                string sql = $"select DISTINCT {TablePO}.CSID from {TablePO} left join {TableCS}  on {TableCS}.CSID={TablePO}.CSID where IsScadaFinsh=0 and IsRequest<=0 ";
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);

                if (!(dt != null && dt.Rows.Count > 0)) { return; }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        await publicClass.DisableScreen(token);

                        string CSID = dt.Rows[i]["CSID"].ToString();
                        string sqlPO = $"select * from {TablePO} where CSID={CSID}";
                        DataTable dtPO = sqlClass.SlcSQL(sqlPO, Connstr);
                        if (dtPO == null || dtPO.Rows.Count <= 0) { continue; }
                        bool isFinsh=true;
                        for (int p = 0; p < dtPO.Rows.Count; p++)
                        {
                            string FilePath = dtPO.Rows[p]["FilePath"].ToString();
                            bool b1 = await IsFillScada(token, dtPO, p, list1);
                            bool b2 = FilePath != "" && File.Exists(FilePath);
                            isFinsh=b1 && b2;
                            if (!isFinsh) { break; }
                        }

                        sqlClass.UpdSQL($"update {TableCS} set IsScadaFinsh='{(isFinsh ? "1" : "0")}' where CSID='{CSID}' ", Connstr);

                    }
                    catch (Exception ex) { await publicClass.ExitSap(token); await publicClass.NoteLog(token, ex, dt_config); }
                }
            }
            catch (Exception ex) {  throw ex; }
        }
        private async Task<bool> IsFillScada(CancellationToken token, DataTable dt, int i, List<string> list)
        {
            bool res = true;
            foreach (string cname in list)
            {
                string str = dt.Rows[i][$"{cname}"].ToString();
                if (str == null || str == "") { res = false; break; }
                await Task.Delay(0, token); token.ThrowIfCancellationRequested();
            }
            return res;
        }



        private async Task IsIntact(CancellationToken token)
        {
            try
            {
                List<string> list1 = new List<string>() { "PONO", "PUser", "PayCode", "PayCodeExplanation", "PostingDate", "Price", "TaxRate", "TaxAmount", "PriceIncludingTax", "PGroup", "SWIFTCode", "FilePath" };
                List<string> list2 = new List<string>() { "PayFinish", "PayFinishExplanation", "Class2" };
                list1.AddRange(list2);   
                List<string> list = list1.Distinct().ToList();            


                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                string sql = $"select DISTINCT {TablePO}.CSID from {TablePO} left join {TableCS}  on {TableCS}.CSID={TablePO}.CSID where IsIntact=0 and IsRequest<=0 ";
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);

                if (!(dt != null && dt.Rows.Count > 0)) { return; }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        await publicClass.DisableScreen(token);

                        string CSID = dt.Rows[i]["CSID"].ToString();

                        string sqlPO = $"select * from {TablePO} where CSID={CSID}";
                        DataTable dtPO = sqlClass.SlcSQL(sqlPO, Connstr);
                        if (dtPO == null || dtPO.Rows.Count <= 0) { continue; }
                        bool isFinsh = true;
                        for (int p = 0; p < dtPO.Rows.Count; p++)
                        {
                            string FilePath = dtPO.Rows[p]["FilePath"].ToString();
                            bool b1 = await IsFillScada(token, dtPO, p, list1);
                            bool b2 = FilePath != "" && File.Exists(FilePath);
                            isFinsh = b1 && b2;
                            if (!isFinsh) { break; }
                        }

                        sqlClass.UpdSQL($"update {TableCS} set IsIntact='{(isFinsh ? "1" : "0")}' where CSID='{CSID}' ", Connstr);

                    }
                    catch (Exception ex) { await publicClass.ExitSap(token); await publicClass.NoteLog(token, ex, dt_config); }
                }


            }
            catch (Exception ex) {  throw ex; }
        }



        #endregion

        #region 填充计算

        private async Task Fill_ConverFinishDate(CancellationToken token)
        {
            try 
            {
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                string sql = $"select DISTINCT {TablePO}.CSID from {TablePO} left join {TableCS}  on {TableCS}.CSID={TablePO}.CSID where IsIntact=0 and IsRequest<=0 ";
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }           

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        await publicClass.DisableScreen(token);

                        string CSID = dt.Rows[i]["CSID"].ToString();

                        //最大的交货日期
                        string PostingDate = "";
                        string sqlMaxPD = $"select max(PostingDate) from {TablePO} where CSID={CSID}";
                        DataTable dtMaxPD = sqlClass.SlcSQL(sqlMaxPD, Connstr);                      
                        if (dtMaxPD != null && dtMaxPD.Rows.Count > 0 && dtMaxPD.Rows[0][0].ToString() != "") { PostingDate = dtMaxPD.Rows[0][0].ToString(); }
                        if (PostingDate == "") { continue; }
                        PostingDate=Convert.ToDateTime(PostingDate).ToString("yyyy-MM-dd");

                        //付款条件代码
                        string PayCode = "";
                        string sqlPayCode = $"select top 1 PayCode from {TablePO} where CSID={CSID} and  PayCode is not null ";
                        DataTable dtPayCode = sqlClass.SlcSQL(sqlPayCode, Connstr);
                        if (dtPayCode != null && dtPayCode.Rows.Count > 0 && dtPayCode.Rows[0][0].ToString() != "") { PayCode = dtPayCode.Rows[0][0].ToString(); }
                        if (PayCode == "") { continue; }

                        //付款条件代码对应的类型和天数
                        string stype = "";
                        int day = -99;
                        if (dt_PayTerm == null|| dt_PayTerm.Rows.Count<=0) { continue; }
                        for (int pt = 0; pt <= dt_PayTerm.Rows.Count; pt++)
                        {
                            if (PayCode == dt_PayTerm.Rows[pt]["代码"].ToString())
                            {
                                stype = dt_PayTerm.Rows[pt]["类型"].ToString();
                                try { day = Convert.ToInt32(dt_PayTerm.Rows[pt]["天数"]); } catch { }
                                break;
                            }
                        }
                        if (stype == ""|| day == -99) { continue; }

                        //计算要求付款完成时间    >>>货到：最晚的过账日期+条件代码天数，票到：最晚的过账日期
                        string PayFinish = "";
                        string PayFinishExplanation = "";
                        if (stype.Contains("货到")) { PayFinish = Convert.ToDateTime(PostingDate).AddDays(Convert.ToInt32(day)).ToString("yyyy-MM-dd"); PayFinishExplanation = $"{PayCode}货到-最晚的过账日期{PostingDate}加{day}天"; }
                        if (stype.Contains("票到")) { PayFinish = PostingDate; PayFinishExplanation = $"{PayCode}票到-最晚的过账日期{PostingDate}"; }

                        //更新
                        if (PayFinish != "")
                        {
                            sqlClass.UpdSQL($"update {TablePO} set PayFinish='{PayFinish}' ,PayFinishExplanation='{PayFinishExplanation}' where CSID='{CSID}' ", Connstr);                            
                        }


                    }
                    catch (Exception ex) { await publicClass.ExitSap(token); await publicClass.NoteLog(token, ex, dt_config); }
                }
            }
            catch (Exception ex) {  throw ex; }
        }


        private async Task Fill_Class2(CancellationToken token)
        {
            try
            {
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                string sql = $"select DISTINCT {TablePO}.CSID from {TablePO} left join {TableCS}  on {TableCS}.CSID={TablePO}.CSID where IsIntact=0 and IsRequest<=0 ";
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        await publicClass.DisableScreen(token);

                        string CSID = dt.Rows[i]["CSID"].ToString();

                        //获取采购组
                        string sqlPO = $"select top 1 PGroup from {TablePO} where CSID={CSID} and PGroup is not null ";
                        DataTable dtPO = sqlClass.SlcSQL(sqlPO, Connstr);
                        if (dtPO == null || dtPO.Rows.Count <= 0) { continue; }
                        string PGroup= dtPO.Rows[0]["PGroup"].ToString();//同一对账表默认所有PGroup一致
                        if ( PGroup=="") { continue; }
                         
                        //通过采购组获取二级分类
                        string Class2 = "";
                        if (dt_Class2 == null|| dt_Class2.Rows.Count<=0) { continue; }
                        for (int c = 0; c < dt_Class2.Rows.Count; c++)
                        {
                            if (PGroup == dt_Class2.Rows[c]["采购组"].ToString()) { Class2 = dt_Class2.Rows[c]["二级分类"].ToString(); break; }
                        }
                        if (Class2 == "") { continue; }

                        //更新
                        sqlClass.UpdSQL($"update {TablePO} set Class2='{Class2}'  where CSID='{CSID}' ", Connstr);

                    }
                    catch (Exception ex) { await publicClass.NoteLog(token, ex, dt_config); }
                }
            }
            catch (Exception ex) {  throw ex; }
        }

        #endregion

        #region 特殊处理
        private async Task SpecialExec(CancellationToken token)
        {
            try
            {

               
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
                string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                string sql = $"select {TablePO}.POID,{TablePO}.PayFinish,{TablePO}.PUser  from {TablePO} left join {TableCS}  on {TableCS}.CSID={TablePO}.CSID where  IsIntact=1 and IsRequest<=0 "; //仅遍历采集完毕的
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        await publicClass.DisableScreen(token);

                        //----处理1：要求付款完成时间小于当前日期----改为当前日期加3天-------
                        string POID = dt.Rows[i]["POID"].ToString();
                        string PayFinish = dt.Rows[i]["PayFinish"].ToString();
                        if (Convert.ToDateTime(PayFinish) < DateTime.Now)
                        {                          
                            sqlClass.UpdSQL($"update {TablePO} set PayFinish='{DateTime.Now.AddDays(3).ToString("yyyy-MM-dd")}' where POID={POID}", Connstr);
                        }

                        //----处理2：采购员离职替换-----------
                        string PUser = dt.Rows[i]["PUser"].ToString();
                        string NewPUser =await GetReplacePUser(token,PUser);
                        if (NewPUser != "")
                        {
                            sqlClass.UpdSQL($"update {TablePO} set PUser='{NewPUser}' where POID={POID}", Connstr);
                        }

                    }
                    catch (Exception ex) { await publicClass.NoteLog(token, ex, dt_config); /*异常>>继续下一个迭代*/}
                }
                

             

                //------处理3：删除用户文件夹ZF004\PDFDir\下的文档-----仅保留最新的100个---------------
                FileInfo[] allPDF = new DirectoryInfo(PDFDir).GetFiles();
                int delCount = 100;
                if (dt != null && dt.Rows.Count > delCount) { delCount = dt.Rows.Count; }
                var filesToDelete1 = allPDF.OrderByDescending(f => f.CreationTimeUtc).Skip(delCount).ToList();
                foreach (var file in filesToDelete1)
                {
                    try
                    {
                        file.Delete();
                    }
                    catch { }
                }

                //------处理4：删除用户文件夹ZF004\PODir\下的文档-----仅保留最新的100个---------------
                FileInfo[] allPO = new DirectoryInfo(PODir).GetFiles();              
                if (dt != null && dt.Rows.Count > delCount) { delCount = dt.Rows.Count; }
                var filesToDelete2 = allPO.OrderByDescending(f => f.CreationTimeUtc).Skip(delCount).ToList();
                foreach (var file in filesToDelete2)
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
            catch (Exception ex) {  throw ex; }



            

        }

        private async Task<string> GetReplacePUser(CancellationToken token, string PUser)
        {
            string NewPUser = "";
            try
            {
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
            string TableCS = "";
            string TablePO = "";
            string WechatKey = "";
            try
            {
                Connstr = await publicClass.GetConnstr(token, dt_config);
                TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
                TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                WechatKey = await publicClass.GetLibValue(token, dt_config, "WechatKey");
                string sql = $"select * from {TableCS} where IsIntact=1 and IsScadaFinsh=1  and IsRequest<=0  ";
                dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
            }
            catch (Exception ex) {  throw ex; }

            for (int i = 0; i < dt.Rows.Count; i++)
            {

                await publicClass.DisableScreen(token);
                string CSID = dt.Rows[i]["CSID"].ToString();
                string Supplier = dt.Rows[i]["Supplier"].ToString();
                string PriceIncludingTax_parent = dt.Rows[i]["PriceIncludingTax"].ToString();
                try
                {                   

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


                       
                        DataTable dtPO = sqlClass.SlcSQL($"select * from {TablePO} where CSID={CSID}", Connstr);
                        if (dtPO == null || dtPO.Rows.Count <= 0) { continue; }
                        wait60 = new WebDriverWait(MyDriver, TimeSpan.FromSeconds(60));

                        string PUser = dtPO.Rows[0]["PUser"].ToString();
                        string SupplierCode = dt.Rows[i]["SupplierCode"].ToString();
                        string PurchaseStype = dtPO.Rows[0]["PurchaseStype"].ToString();
                        string PayStype = dtPO.Rows[0]["PayStype"].ToString();
                        string Class2 = dtPO.Rows[0]["Class2"].ToString();
                        string IsExistPO = dtPO.Rows[0]["IsExistPO"].ToString();
                        string CompanyCode = dt.Rows[i]["CompanyCode"].ToString();
                        string BankOf = dtPO.Rows[0]["BankOf"].ToString();
                        string SupplierShortName = dt.Rows[i]["SupplierShortName"].ToString();
                        string SWIFTCode = dtPO.Rows[0]["SWIFTCode"].ToString();
                        string IsLinkPR = dtPO.Rows[0]["IsLinkPR"].ToString();
                        string PayFinish = dtPO.Rows[0]["PayFinish"].ToString();
                        string PayFinishExplanation = dtPO.Rows[0]["PayFinishExplanation"].ToString();
                        string Currency = dt.Rows[i]["Currency"].ToString();
                        string IsInfrastructure = dtPO.Rows[0]["IsInfrastructure"].ToString();
                        string PriceIncludingTax = dt.Rows[i]["PriceIncludingTax"].ToString();
                        string FilePath = dtPO.Rows[0]["FilePath"].ToString();
                        string CSNO = dt.Rows[i]["CSNO"].ToString();

                        await PostWillDo(CSNO, PUser, Supplier, CompanyCode, PriceIncludingTax_parent, WechatKey);//POST 执行前的预告

                        await LoginOA(token, dt_config);    //01-登录OA
                        await Nav_Liucheng(token);          //02-选择顶部菜单《流程》
                        await New_Liucheng(token);          //03-新建流程
                        await ClickZF004(token);            //04-点击ZF004流程，并切换新窗口


                        await FillPUser(token, PUser);                      //01-填写申请人（新窗口）
                        await FillSupplierCode(token, SupplierCode, CSID);        //02-填写供应商代码
                        await FillPurchaseStype(token, PurchaseStype);      //03-选择<付款类型> 货到付款
                        await FillPayStype(token, PayStype);                //04-选择<支付类型>  材料款
                        await FillClass2(token, Class2);                    //05-填写二级分类
                        await FillIsIsExistPO(token, IsExistPO);            //06-选择是否有采购订单号
                        await FillCompanyCode(token, CompanyCode);          //07-填写付款公司（公司代码）
                        await FillBankOf(token, BankOf);                    //08-选择对公对私
                        await FillSupplierShortName(token, SupplierShortName);  //09-填写供应商简称
                        await FillSWIFTCode(token, SWIFTCode);                  //10-填写收款人国家/地区（银行国家代码）
                        await FillIsLinkPR(token, IsLinkPR);                    //11-选择是否关联资本性支出请购流程
                        await FillPayFinish(token, PayFinish);                  //12-填写付款要求完成时间                     
                        await FillCurrency(token, Currency);                    //13-支付币种
                        await FillIsInfrastructure(token, IsInfrastructure);    //14-选择是否基建类
                        await FillFilePath(token, FilePath, CSID);              //15-填写相关附件-上传附件


                        await FillPO_NO(token, dtPO, CSID);         //01-循环填入PO号
                        await ResPayInfo_SaveBefore(token, CSID);   //02-第一次刷新支付信息-支付金额（必须在保存前）                       
                        await SaveOA(token, CSID);                  //03-第一次保存OA----------------------------------注意保存后bodyiframe变化-----                                                                
                        await DelRepeatPO(token);                   //04-勾选删除重复PO
                        await FillPO_Detail(token, dtPO);           //05-填写PO价格等其他明细信息（必须在保存后）
                        await ResPayInfo_SaveAfter(token);          //06-第二次刷新支付信息-支付金额（价格更新后再次刷新）
                        await SaveOA(token, CSID);                  //07-第二次保存OA

                        await CheckPriceIncludingTax(token, PriceIncludingTax, CSID);   //01-校验含税总额是否一致 
                        await Countersign(token, CSNO);                                 //02-签字意见
                        await SubOA(token, CSID);                            //03-提交OA




                        await Task.Delay(5000, token);


                    }
                }
                catch (Exception ex) { sqlClass.UpdSQL($"update {TableCS} set ResInfo='{ex.Message}' where CSID={CSID} ",Connstr);   }

                await PostResInfo(TableCS, TablePO,CSID, WechatKey);  //POST 执行后的结果
            }
        }


        public async Task PostWillDo(string CSNO, string PUser, string Supplier,string CompanyCode, string PriceIncludingTax, string WechatKey)
        {
            if (WechatKey != "")
            {
                var requestBody = new
                {
                    msgtype = "markdown",
                    markdown = new
                    {
                        content = "到付订单即将开跑，对账表编号为：<font color=\"warning\">" + CSNO + "</font>，请<font color=\"warning\">" + PUser + "</font>注意执行结果。\n" +
                              ">供应商:<font color=\"comment\">" + Supplier + "</font>\n" +
                              ">公司代码:<font color=\"comment\">" + CompanyCode + "</font>\n" +
                              ">含税金额:<font color=\"comment\">" + PriceIncludingTax + "</font>\n" +
                              ">执行结果:<font color=\"comment\">正在执行</font>"
                    }
                };
                string jsonContent = JsonConvert.SerializeObject(requestBody);
                await publicClass.WechatPost(WechatKey, jsonContent);
            }

        }

        public async Task PostResInfo(string TableCS, string TablePO, string CSID,  string WechatKey)
        {
            if (WechatKey != "")
            {
                string sql = $"select  (select top 1 PUser from {TablePO} where {TablePO}.CSID={TableCS}.CSID) as 'PUser',* from  {TableCS} where CSID={CSID}";
                DataTable dt = sqlClass.SlcSQL(sql, Connstr);
                if (!(dt != null && dt.Rows.Count > 0)) { return; }
                string CSNO = dt.Rows[0]["CSNO"].ToString();             
                string Supplier = dt.Rows[0]["Supplier"].ToString();
                string CompanyCode = dt.Rows[0]["CompanyCode"].ToString();
                string PUser = dt.Rows[0]["PUser"].ToString();
                string PriceIncludingTax = dt.Rows[0]["PriceIncludingTax"].ToString();
                string ResInfo = dt.Rows[0]["ResInfo"].ToString().Trim();

                var requestBody = new
                {
                    msgtype = "markdown",
                    markdown = new
                    {
                        content = "到付订单执行完毕，对账表编号为：<font color=\"warning\">" + CSNO + "</font>，请<font color=\"warning\">" + PUser + "</font>注意执行结果。\n" +
                              ">供应商:<font color=\"comment\">" + Supplier + "</font>\n" +
                              ">公司代码:<font color=\"comment\">" + CompanyCode + "</font>\n" +
                              ">含税金额:<font color=\"comment\">" + PriceIncludingTax + "</font>\n" +
                              ">执行结果:<font color=\"comment\">"+ (ResInfo==""?"异常，等下次重试": ResInfo )+ "</font>"
                    }
                };
                string jsonContent = JsonConvert.SerializeObject(requestBody);
                await publicClass.WechatPost(WechatKey, jsonContent);

            }
        }










        //------------------------------------------------------------------------------------------------

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
            catch (Exception ex) {  throw ex; }
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
            catch (Exception ex) {  throw ex; }
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
            catch (Exception ex) {  throw ex; }
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
            catch (Exception ex) {  throw ex; }

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
                var liCss2s = drillmenu.FindElements(By.CssSelector("li.liCss2"));
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
            catch (Exception ex) {  throw ex; }
        }
        private async Task ClickZF004(CancellationToken token, int timeout = 60)
        {
            try
            {
                List<string> originalWindow = MyDriver.WindowHandles.ToList(); // 获取所有窗口句柄                


                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.XPath("//iframe[@id='mainFrame']"));
                await SwitchToIframe(token, By.XPath("//iframe[@id='tabcontentframe']"));

                //点击<ZF003-预付款申请流程>
                string xpathToFind = "//table[@class='ViewForm']//a[@class='e8contentover']";
                IWebElement myClick = null;
                var e8contentovers = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                for (int t = 0; t < timeout; t++)
                {
                    foreach (IWebElement e8contentover in e8contentovers)
                    {
                        if (e8contentover.Text.ToUpper().Contains("ZF004-到付")) { myClick = e8contentover; break; }
                    }
                    if (myClick != null) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    e8contentovers = MyDriver.FindElements(By.XPath($"{xpathToFind}"));
                    if (t + 1 == timeout) { throw new Exception("未能找到<ZF004-到付/质保金付款审批流程>的链接"); }
                }
                myClick.Click();//点击

                //切换新窗口
                await SwitchToWindow(token, originalWindow, 60);

            }
            catch (Exception ex) {  throw ex; }
        }
        private async Task FillPUser(CancellationToken token, string PUser, int timeout = 60)
        {
            try
            {

                await SwitchToDefaultContent(token);
                await SwitchToIframe(token, By.XPath("//iframe[@id='bodyiframe']"));

                //点击《申请人》搜索按钮
                var field91836_browserbtns = MyDriver.FindElements(By.Id("field91836_browserbtn"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91836_browserbtns.Count > 0) { field91836_browserbtns[0].Click(); break; }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91836_browserbtns = MyDriver.FindElements(By.Id("field91836_browserbtn"));
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
            catch (Exception ex) {  throw ex; }
        }
        private async Task FillSupplierCode(CancellationToken token, string supplierCode,string CSID, int timeout = 60)
        {
            try
            {

                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.XPath("//iframe[@id='bodyiframe']"));

                //输入供应商代码
                var field91845s = MyDriver.FindElements(By.Id("field91845"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91845s.Count > 0)
                    {
                        field91845s[0].Clear();
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        field91845s[0].SendKeys(supplierCode);
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        field91845s[0].SendKeys(OpenQA.Selenium.Keys.Tab);
                        break;
                    }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91845s = MyDriver.FindElements(By.Id("field91845"));
                    if (t + 1 == timeout) { throw new Exception("未能找到供应商代码输入框"); }
                }

                //弹出错误信息
                MyDriver.SwitchTo().DefaultContent();
                var Message_undefineds = MyDriver.FindElements(By.Id("Message_undefined"));
                for (int t = 0; t < timeout; t++)
                {
                    if (Message_undefineds.Count > 0)
                    {
                        string MesssageE = Message_undefineds[0].Text;
                        Connstr = await publicClass.GetConnstr(token, dt_config);
                        string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
                        sqlClass.UpdSQL($"update {TableCS} set IsRequest=2 , ResInfo='{MesssageE}' where CSID='{CSID}'", Connstr);                      
                        throw new Exception(MesssageE);
                        //break;
                    }
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested();//100*60=6秒
                    MyDriver.SwitchTo().DefaultContent();
                    Message_undefineds = MyDriver.FindElements(By.Id("Message_undefined"));
                }


            }
            catch (Exception ex) {  throw ex; }

        }     
        private async Task FillPurchaseStype(CancellationToken token, string PurchaseStype, int timeout = 60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //选择付款类型
                var field91842S = MyDriver.FindElements(By.Id("field91842"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91842S.Count > 0)
                    {
                        var select = new SelectElement(field91842S[0]); //Selenium.Support 包 
                        select.SelectByText(PurchaseStype);
                        break;
                    }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91842S = MyDriver.FindElements(By.Id("field91842"));
                    if (t + 1 == timeout) { throw new Exception("未能找到付款类型的选项框"); }
                }

            }
            catch (Exception ex) {  throw ex; }
        }
        private async Task FillPayStype(CancellationToken token, string PayStype, int timeout = 60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //选择 支付类型
                var field92024S = MyDriver.FindElements(By.Id("field92024"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field92024S.Count > 0)
                    {
                        var select = new SelectElement(field92024S[0]); //Selenium.Support 包 
                        select.SelectByText(PayStype);
                        break;
                    }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field92024S = MyDriver.FindElements(By.Id("field92024"));
                    if (t + 1 == timeout) { throw new Exception("未能找到付款类型的选项框"); }
                }

            }
            catch (Exception ex) {  throw ex; }


        }
        private async Task FillClass2(CancellationToken token, string Class2, int timeout = 60)
        {
            try
            {

                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");


                //选择二级分类
                var field92057s = MyDriver.FindElements(By.Id("field92057"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field92057s.Count > 0)
                    {
                        var select = new SelectElement(field92057s[0]); //Selenium.Support 包 
                        select.SelectByText(Class2);
                        break;
                    }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field92057s = MyDriver.FindElements(By.Id("field92057"));
                    if (t + 1 == timeout) { throw new Exception("未能找到二级分类的选项框"); }
                }

            }
            catch (Exception ex) {  throw ex; }

        }
        private async Task FillIsIsExistPO(CancellationToken token, string IsExistPO, int timeout = 60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");


                //选择是否有采购订单号
                var field91905s = MyDriver.FindElements(By.Id("field91905"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91905s.Count > 0)
                    {
                        var select = new SelectElement(field91905s[0]); //Selenium.Support 
                        select.SelectByText(IsExistPO);
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91905s = MyDriver.FindElements(By.Id("field91905"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到是否有采购订单号的选项框"); }
                }

            }
            catch (Exception ex) {  throw ex; }

        }
        private async Task FillCompanyCode(CancellationToken token, string CompanyCode, int timeout = 60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");


                //点击搜索按钮
                var field91995_browserbtns = MyDriver.FindElements(By.Id("field91995_browserbtn"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91995_browserbtns.Count > 0) { field91995_browserbtns[0].Click(); break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91995_browserbtns = MyDriver.FindElements(By.Id("field91995_browserbtn"));
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
            catch (Exception ex) {  throw ex; }

        }
        private async Task FillBankOf(CancellationToken token, string BankOf, int timeout = 60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //选择对公对私
                var field92045s = MyDriver.FindElements(By.Id("field92045"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field92045s.Count > 0)
                    {
                        var select = new SelectElement(field92045s[0]); //Selenium.Support 
                        select.SelectByText(BankOf);
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field92045s = MyDriver.FindElements(By.Id("field92045"));
                    if (t + 1 == timeout) { throw new Exception("未能找到对公对私的选项框"); }
                }


            }
            catch (Exception ex) {  throw ex; }
        }
        private async Task FillSupplierShortName(CancellationToken token, string SupplierShortName, int timeout = 60)
        {
            try
            {

                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");


                //输入供应商简称
                var field91847s = MyDriver.FindElements(By.Id("field91847"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91847s.Count > 0)
                    {
                        field91847s[0].Clear();
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        field91847s[0].SendKeys(SupplierShortName);
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91847s = MyDriver.FindElements(By.Id("field91847"));
                    if (t + 1 == timeout) { throw new Exception("未能找到供应商简称的输入框"); }
                }
            }
            catch (Exception ex) {  throw ex; }


        }
        private async Task FillSWIFTCode(CancellationToken token, string SWIFTCode, int timeout = 60)
        {
            try
            {

                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //点击搜索按钮
                var field92054_browserbtns = MyDriver.FindElements(By.Id("field92054_browserbtn"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field92054_browserbtns.Count > 0) { field92054_browserbtns[0].Click(); break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field92054_browserbtns = MyDriver.FindElements(By.Id("field92054_browserbtn"));
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
            catch (Exception ex) {  throw ex; }
        }
        private async Task FillIsLinkPR(CancellationToken token, string isLinkPR, int timeout = 60)
        {
            try
            {
                //切换iframe
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //选择是否关联资本性支出请购流程
                var field92030s = MyDriver.FindElements(By.Id("field92030"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field92030s.Count > 0)
                    {
                        var select = new SelectElement(field92030s[0]); //Selenium.Support 
                        select.SelectByText(isLinkPR);
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field92030s = MyDriver.FindElements(By.Id("field92030"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到是否关联资本性支出请购流程的选项框"); }
                }
            }
            catch (Exception ex) {  throw ex; }
        }
        private async Task FillPayFinish(CancellationToken token, string PayFinish, int timeout = 60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //点击日期图标
                var field91968browsers = MyDriver.FindElements(By.Id("field91968browser"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91968browsers.Count > 0)
                    {
                        field91968browsers[0].Click();
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91968browsers = MyDriver.FindElements(By.Id("field91968browser"));
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
                var field91968browsers2 = MyDriver.FindElements(By.Id("field91968browser"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91968browsers2.Count > 0)
                    {
                        field91968browsers2[0].Click();
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91968browsers2 = MyDriver.FindElements(By.Id("field91968browser"));
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
                var spanCK = MyDriver.FindElements(By.XPath("//span[@id='field91968span']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (spanCK.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    spanCK = MyDriver.FindElements(By.XPath("//span[@id='field91968span']"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到校验日期的spanCK"); }
                }
                var InputCK = MyDriver.FindElements(By.XPath("//Input[@id='field91968']"));
                for (int t = 0; t < timeout; t++)
                {
                    if (InputCK.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    InputCK = MyDriver.FindElements(By.XPath("//Input[@id='field91968']"));
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
            catch (Exception ex) {  throw ex; }

        }
        private async Task FillPayFinishExplanation(CancellationToken token, string PayFinishExplanation, int timeout = 60)
        {
            try
            {

                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");


                //付款日期说明
                var field99801s = MyDriver.FindElements(By.Id("field99801"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field99801s.Count > 0)
                    {
                        field99801s[0].Clear();
                        await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                        field99801s[0].SendKeys(PayFinishExplanation);
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field99801s = MyDriver.FindElements(By.Id("field99801"));
                    if (t + 1 == timeout) { throw new Exception("未能找到供应商简称的输入框"); }
                }
            }
            catch (Exception ex) {  throw ex; }
        }
        private async Task FillCurrency(CancellationToken token, string currency, int timeout = 60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.XPath("//iframe[@id='bodyiframe']"));


                //点击《支付币种》搜索按钮
                var field91939_browserbtns = MyDriver.FindElements(By.Id("field91939_browserbtn"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field91939_browserbtns.Count > 0) { field91939_browserbtns[0].Click(); break; }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field91939_browserbtns = MyDriver.FindElements(By.Id("field91939_browserbtn"));
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
            catch (Exception ex) {  throw ex; }
        }
        private async Task FillIsInfrastructure(CancellationToken token, string IsInfrastructure, int timeout = 60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //选择是否基建类
                var field92028s = MyDriver.FindElements(By.Id("field92028"));
                for (int t = 0; t < timeout; t++)
                {
                    if (field92028s.Count > 0)
                    {
                        var select = new SelectElement(field92028s[0]); //Selenium.Support 
                        select.SelectByText(IsInfrastructure);
                        break;
                    }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    field92028s = MyDriver.FindElements(By.Id("field92028"));
                    if (t + 1 == timeout) { throw new Exception($"未能找到是否基建类的选项框"); }
                }

            }
            catch (Exception ex) {  throw ex; }

        }


        private async Task FillFilePath(CancellationToken token, string FilePath, string CSID, int timeout = 60)
        {
            try
            {

                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp"); //保存前


                if (!File.Exists(FilePath))
                {
                    Connstr = await publicClass.GetConnstr(token, dt_config);
                    string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
                    string TablePO = await publicClass.GetLibValue(token, dt_config, "TablePO");
                    sqlClass.UpdSQL($"update {TablePO} set FilePath=null where CSID='{CSID}'", Connstr);
                    sqlClass.UpdSQL($"update {TableCS} set IsScadaFinsh=0,IsIntact=0,ResInfo='本地附件不存在。{FilePath}' where CSID='{CSID}'", Connstr);
                    throw new Exception("本地附件不存在。{FilePath}");
                }


                //上传附件
                var spanButtonPlaceHolder91996s = MyDriver.FindElements(By.Id("spanButtonPlaceHolder91996"));
                for (int t = 0; t < timeout; t++)
                {
                    if (spanButtonPlaceHolder91996s.Count > 0)
                    {
                        IWebElement parent = spanButtonPlaceHolder91996s[0].FindElement(By.XPath("./..")); ;
                        IWebElement input = parent.FindElement(By.TagName("input"));
                        input.SendKeys(FilePath);
                        break;
                    }

                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    spanButtonPlaceHolder91996s = MyDriver.FindElements(By.Id("spanButtonPlaceHolder91996"));
                    if (t + 1 == timeout) { throw new Exception("未能找到上传附件按钮"); }

                }

                //等待上传成功
                await Task.Delay(3000, token); token.ThrowIfCancellationRequested();

            }
            catch (Exception ex) {  throw ex; }
        }





        private async Task FillPO_NO(CancellationToken token, DataTable dtPO, string CSID, int timeout = 60)
        {
            try
            {
               

                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "workflow/request/AddRequestIframe.jsp");

                //去重后的PO
                if (dtPO == null|| dtPO.Rows.Count<=0) { return; }
                List<string> listPO = new List<string>();
                for (int p = 0; p < dtPO.Rows.Count; p++) { listPO.Add(dtPO.Rows[p]["PONO"].ToString()); }
                listPO = listPO.Distinct().ToList();

                //采购订单查询  
                foreach (string PO in listPO)
                {
                    var field93861S = MyDriver.FindElements(By.Id("field93861"));
                    for (int t = 0; t < timeout; t++)
                    {
                        if (field93861S.Count > 0)
                        {
                            field93861S[0].Clear();
                            field93861S[0].SendKeys(PO);
                            await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                            field93861S[0].SendKeys(OpenQA.Selenium.Keys.Tab);
                            break;
                        }
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                        field93861S = MyDriver.FindElements(By.Id("field93861"));
                        if (t + 1 == timeout) { throw new Exception("未能找到采购订单查询输入框"); }
                    }
                    await Task.Delay(500, token); token.ThrowIfCancellationRequested();
                }
                          
               
            }
            catch (Exception ex) {  throw ex; }
        }

        private async Task ResPayInfo_SaveBefore(CancellationToken token,string CSID, int timeout = 60)
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
                        if (Message_undefineds.Count > 0) { MesssageE = "<OA异常>" + Message_undefineds[0].Text.Trim(); break; }
                        await Task.Delay(200, token); token.ThrowIfCancellationRequested();
                        Message_undefineds = MyDriver.FindElements(By.Id("Message_undefined"));
                    }
                }
                catch { }
                if (MesssageE != "")
                {
                    Connstr = await publicClass.GetConnstr(token, dt_config);
                    string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
                    sqlClass.UpdSQL($"update {TableCS} set IsRequest=2 , ResInfo='{MesssageE}' where CSID='{CSID}'", Connstr);
                    throw new Exception($"{MesssageE}");
                }



            }
            catch (Exception ex) {  throw ex; }
        }
        private async Task SaveOA(CancellationToken token, string CSID, int timeout = 60)
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
                        if (inputSaves[0].Enabled)
                        {
                            inputSaves[0].Click();
                            break;
                        }                       
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
                string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
                sqlClass.UpdSQL($"update {TableCS} set RequestID='{requestid}',IsRequest=0 where CSID='{CSID}'", Connstr);




                //等待保存成功      
                var inputSavesW = MyDriver.FindElements(By.XPath("//div[@id='null_box']//input[@value='保存']"));
                for (int t = 0; t < timeout; t++)
                {                    
                    await Task.Delay(2000, token); token.ThrowIfCancellationRequested(); //置前

                    if (inputSavesW.Count > 0)
                    {
                        try
                        {
                            if (inputSavesW[0].Enabled) { break; }
                        }
                        catch { /*灰色按钮时，有元素但点击不到，这里仅做捕获，不做处理*/ }
                    }
                    inputSavesW = MyDriver.FindElements(By.XPath("//div[@id='null_box']//input[@value='保存']"));
                    if (t + 1 == timeout) { throw new Exception("未能找到保存按钮"); }
                }
              
               





            }
            catch (Exception ex) {  throw ex; }
        }
        //-----保存后------------------------
        private async Task DelRepeatPO(CancellationToken token, int timeout = 60)
        {            
            try
            {
              
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "ManageRequestNoFormIframe.jsp"); //保存后bodyiframe是ManageRequestNoFormIframe             

                //等待加载：oTable2
                string oTable2 = "//table[@id='oTable2']//tbody//tr[@_target='datarow']";
                var trs = MyDriver.FindElements(By.XPath($"{oTable2}"));
                for (int t = 0; t < timeout; t++)
                {
                    if (trs.Count > 0) { break; }
                    //-------------------------------------
                    /*加此两行是因为本函数上个方法是保存，保存DOM发生变更，可能未能切换缘故*/
                    MyDriver.SwitchTo().DefaultContent();
                    await SwitchToIframe(token, By.Id("bodyiframe"), "ManageRequestNoFormIframe.jsp"); //保存后bodyiframe是ManageRequestNoFormIframe
                    //-------------------------------------
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    trs = MyDriver.FindElements(By.XPath($"{oTable2}"));
                    if (t + 1 == timeout) { throw new Exception("datarow数据行的tr行未完全加载"); }
                }
                //等待加载：oTable2
                for (int t = 0; t < timeout; t++)
                {
                    var w0 = trs[0].FindElements(By.XPath(".//td"));
                    if (w0.Count <= 0) { continue; }
                    var w1 = w0[0].FindElements(By.XPath(".//input[@type='checkbox']"));
                    var w2 = trs[0].FindElements(By.XPath(".//input[@temptitle='采购订单2']"));                   

                    if (w1.Count > 0 && w2.Count > 0 ) { break; }
                    if (t + 1 == timeout) { throw new Exception("datarow数据行的tr行的相应的列未完全加载"); }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    trs = MyDriver.FindElements(By.XPath($"{oTable2}"));
                }

                //勾选重复的行
                bool IsNeedDel=false;
                List<string> list = new List<string>();
                foreach (var tr in trs)
                {
                    IWebElement td0 = tr.FindElement(By.XPath(".//td"));
                    IWebElement checkbox = td0.FindElement(By.XPath(".//input[@type='checkbox']"));
                    IWebElement input = tr.FindElement(By.XPath(".//input[@temptitle='采购订单2']"));

                    string thisPO = input.GetDomAttribute("value");
                    if (list.Contains(thisPO))
                    {
                        IsNeedDel = true;
                        checkbox.Click(); //勾选
                    }
                    list.Add(thisPO); //无论如何都加，后加
                }

                //单击减号 删除
                if (IsNeedDel)
                {
                    for (int t = 0; t < timeout; t++)
                    { 
                        var delbuttons = MyDriver.FindElements(By.XPath($"//div[@id='div2button']/button[@title='删除']"));
                        if (delbuttons.Count > 0)
                        { 
                            delbuttons[0].Click();  //点击删除
                            await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                            //处理删除询问的弹窗
                            IAlert alert = MyDriver.SwitchTo().Alert();
                            alert.Accept();
                            break;
                        }
                        if (t + 1 == timeout) { throw new Exception("找不到 减号 删除 按钮"); }
                    }                     
                }
            }

            catch (Exception ex) {  throw ex; }
        }

        private async Task FillPO_Detail(CancellationToken token, DataTable dtPO, int timeout = 60)
        {
            try
            {      

                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "ManageRequestNoFormIframe.jsp"); //保存后bodyiframe是ManageRequestNoFormIframe

                //等待加载：oTable2
                string oTable2 = "//table[@id='oTable2']//tbody//tr[@_target='datarow']";
                var trs = MyDriver.FindElements(By.XPath($"{oTable2}"));
                for (int t = 0; t < timeout; t++)
                {
                    if (trs.Count > 0) { break; }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    trs = MyDriver.FindElements(By.XPath($"{oTable2}"));
                    if (t + 1 == timeout) { throw new Exception("datarow数据行的tr行未完全加载"); }
                }
                //等待加载：oTable2
                for (int t = 0; t < timeout; t++)
                {                   
                    var w1 = trs[0].FindElements(By.XPath(".//input[@temptitle='采购订单2']"));
                    var w2 = trs[0].FindElements(By.XPath(".//input[@temptitle='本次付款金额(含税)1']"));
                    if (w1.Count > 0 && w2.Count > 0) { break; }
                    if (t + 1 == timeout) { throw new Exception("datarow数据行的tr行的相应的列未完全加载"); }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    trs = MyDriver.FindElements(By.XPath($"{oTable2}"));
                }

                //遍历datarow数据行，并设置每行的必填列
                for (int i = 0; i < trs.Count; i++)
                {                   
                    string PONO = trs[i].FindElement(By.XPath(".//input[@temptitle='采购订单2']")).GetDomAttribute("value");
                    double benfuhan =await GetBenFuHan(token, PONO, dtPO);

                    //先填写：本付款(含)[需先清空]    {本付未为不可填写状态}
                    IWebElement w2 = trs[i].FindElement(By.XPath(".//input[@temptitle='本次付款金额(含税)1']"));
                    w2.Clear();
                    w2.SendKeys(benfuhan.ToString());
                    await Task.Delay(100, token); token.ThrowIfCancellationRequested();
                    w2.SendKeys(OpenQA.Selenium.Keys.Tab);
                    await Task.Delay(500, token); token.ThrowIfCancellationRequested();


                    //单据完整性
                    IWebElement wzx = trs[i].FindElement(By.XPath(".//select[@temptitle='单据完整性']"));
                    var wzxE = new SelectElement(wzx); //Selenium.Support 
                    wzxE.SelectByText($"{dtPO.Rows[0]["IsOrderComplete"].ToString()}");


                    //// 备注原因 
                    //IWebElement bzyy = trs[i].FindElement(By.XPath(".//input[@temptitle='备注原因']"));
                    //bzyy.SendKeys("非必填");

                }

            }
            catch (Exception ex) {  throw ex; }
        }

        private async Task<double> GetBenFuHan(CancellationToken token, string PONO, DataTable dtPO)
        {
            double benfuhan = 0;
            try 
            {
                if (dtPO == null || dtPO.Rows.Count <=0) { return benfuhan; }
                for (int i = 0; i < dtPO.Rows.Count; i++)
                {
                    if (PONO == dtPO.Rows[i]["PONO"].ToString())
                    {
                        benfuhan += Convert.ToDouble(dtPO.Rows[i]["PriceIncludingTax"]);
                    }
                    await Task.Delay(0, token); token.ThrowIfCancellationRequested();
                }
            }
            catch { }
            return benfuhan;
        }

        private async Task ResPayInfo_SaveAfter(CancellationToken token, int timeout = 60)
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
            catch (Exception ex) {  throw ex; }

        }

        private async Task CheckPriceIncludingTax(CancellationToken token, string PriceIncludingTax, string CSID,int timeout = 60)
        {
            //校验含税总额是否一致         
            try
            {
                bool res = false;
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "ManageRequestNoFormIframe.jsp"); //保存后bodyiframe是ManageRequestNoFormIframe


                //等待加载：oTable2
                string oTable2 = "//table[@id='oTable2']//tbody//tr[@_target='datarow']";
                var trs = MyDriver.FindElements(By.XPath($"{oTable2}"));
                for (int t = 0; t < timeout; t++)
                {
                    if (trs.Count > 0) { break; }

                    //-------------------------------------
                    /*加此两行是因为本函数上个方法是保存，保存DOM发生变更，可能未能切换缘故*/
                    MyDriver.SwitchTo().DefaultContent();
                    await SwitchToIframe(token, By.Id("bodyiframe"), "ManageRequestNoFormIframe.jsp"); //保存后bodyiframe是ManageRequestNoFormIframe
                    //-------------------------------------
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    trs = MyDriver.FindElements(By.XPath($"{oTable2}"));
                    if (t + 1 == timeout) { throw new Exception("datarow数据行的tr行未完全加载-CheckPriceIncludingTax01"); }
                }
                //等待加载：oTable2
                for (int t = 0; t < timeout; t++)
                {
                    var w1 = trs[0].FindElements(By.XPath(".//input[@temptitle='采购订单2']"));
                    var w2 = trs[0].FindElements(By.XPath(".//input[@temptitle='本次付款金额(含税)1']"));
                    if (w1.Count > 0 && w2.Count > 0) { break; }
                    if (t + 1 == timeout) { throw new Exception("datarow数据行的tr行的相应的列未完全加载-CheckPriceIncludingTax02"); }
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    trs = MyDriver.FindElements(By.XPath($"{oTable2}"));
                }

                //遍历datarow数据行，并设置每行的必填列
                double OAPriceIncludingTax = 0;
                for (int i = 0; i < trs.Count; i++)
                {
                    IWebElement w2 = trs[i].FindElement(By.XPath(".//input[@temptitle='本次付款金额(含税)1']"));
                    string value = w2.GetDomAttribute("value");
                    OAPriceIncludingTax += Convert.ToDouble(value.Replace(",",""));
                }

                res = Math.Round(OAPriceIncludingTax, 2) == Math.Round(Convert.ToDouble(PriceIncludingTax), 2);

                if (CSID == "314")                
                {
                    string dd= "断点专用";
                }
              

                if (!res)
                {


                    Connstr = await publicClass.GetConnstr(token, dt_config);
                    string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");
                    sqlClass.UpdSQL($"update {TableCS} set IsRequest=2,ResInfo='填写OA的付款总额（含税）和SAP不一致，{OAPriceIncludingTax}>>{PriceIncludingTax}' where CSID='{CSID}'", Connstr);
                    throw new Exception("填写OA的付款总额（含税）和SAP不一致");
                }

            }
            catch (Exception ex) {  throw ex; }
           
        }

        private async Task Countersign(CancellationToken token, string CSNO, int timeout = 60)
        {
            try
            {
                MyDriver.SwitchTo().DefaultContent();
                await SwitchToIframe(token, By.Id("bodyiframe"), "ManageRequestNoFormIframe.jsp"); //保存后bodyiframe是ManageRequestNoFormIframe

               
                await SwitchToIframe(token, By.Id("ueditor_0"), "about");

                var viewPs=MyDriver.FindElements(By.XPath($"//body[@class='view']/p"));
                for (int t = 0; t < timeout; t++)
                {                   
                    if (viewPs.Count > 0 ) { break; }
                   
                    await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                    viewPs = MyDriver.FindElements(By.XPath($"//body[@class='view']/p"));
                    if (t + 1 == timeout) { throw new Exception("未能定位签字意见元素"); }
                }
              
                viewPs[0].SendKeys($"来自RPA提交。所属对账表编号：{CSNO}");

            }
            catch (Exception ex) {  throw ex; }
        }
        private async Task SubOA(CancellationToken token, string CSID, int timeout = 60)
        {
            try
            {
                Connstr = await publicClass.GetConnstr(token, dt_config);
                string TableCS = await publicClass.GetLibValue(token, dt_config, "TableCS");

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
                string msg = "";
                try
                {
                   
                    MyDriver.SwitchTo().DefaultContent();
                    var Message_undefineds = MyDriver.FindElements(By.Id("Message_undefined"));
                    for (int t = 0; t < 10; t++)
                    {
                        if (Message_undefineds.Count > 0)
                        {
                            msg = Message_undefineds[0].Text.Trim();
                            break;
                        }
                        await Task.Delay(1000, token); token.ThrowIfCancellationRequested();
                        MyDriver.SwitchTo().DefaultContent();
                        Message_undefineds = MyDriver.FindElements(By.Id("Message_undefined"));                       
                    }                   
                }
                catch { }
                //设置IsRequest=2               
                if (msg != "")
                {
                    sqlClass.UpdSQL($"update {TableCS} set IsRequest=2,ResInfo='{msg}' where CSID={CSID} ", Connstr);
                    throw new Exception(msg);
                }




                //只点一次 不成功直接下次迭代。判断是否成功（窗口两个变为一个）成功break 到时失败下次迭代
                for (int t = 0; t <= 300; t++)
                {
                    int count2 = MyDriver.WindowHandles.Count;
                    if (count1 > count2)
                    {
                        sqlClass.UpdSQL($"update {TableCS} set IsRequest=1,ResInfo='成功' where CSID={CSID} ", Connstr);
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
