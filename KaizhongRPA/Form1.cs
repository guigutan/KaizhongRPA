using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace KaizhongRPA
{
    public partial class Form1 : Form
    {
        PublicClass publicClass = new PublicClass();
        public Form1()
        {
            InitializeComponent();
            LoadCheckBox3();
            SetRpaList_User();
        }
        private void Form1_Shown(object sender, EventArgs e)
        {
            ShowDataGridView();
            if (checkBox3.Checked) { btn_Start_Click(null, null); }
        }

        #region 运行/停止按钮
        private void SafeInvoke(Action action)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(action);
            }
            else
            {
                action();
            }
        }              

        private void LabMsg(string msg, bool IsAppend = false)
        {
            SafeInvoke(() =>
            {
                if (IsAppend) { lab_msg.Text += msg; }
                else { lab_msg.Text = msg; }
            });
        }

        private void ButtonEnabled(System.Windows.Forms.Button button, bool enabled = true)
        {
            SafeInvoke(() => button.Enabled = enabled);
        }

        public static bool IsThreadStop = true;
        private void btn_Start_Click(object sender, EventArgs e)
        {            
            if (MyGlobal.MyCts == null || MyGlobal.MyTask.IsCompleted)
            {
                LabMsg($"", false);
                ButtonEnabled(btn_Start, false);
                ButtonEnabled(btn_Stop, true);
                IsThreadStop = false;
                MyGlobal.MyCts = new CancellationTokenSource();
                List<RpaInfo> list = GetRpaList_User();
                //DoWork方法运行在后台任务中，检查取消请求和暂停状态
                MyGlobal.MyTask = Task.Run(() => DoWork(MyGlobal.MyCts.Token, list)); 
            }
        }
               
        private void btn_Stop_Click(object sender, EventArgs e)
        {           
            MyGlobal.MyCts?.Cancel();
            IsThreadStop = true;
            ButtonEnabled(btn_Start, true);
            ButtonEnabled(btn_Stop, false);
        }

        #endregion


        private async Task DoWork(CancellationToken token, List<RpaInfo> list)
        {
            try
            {
                await CountDown(token, 10, "倒计时");
                while (true)
                {
                    foreach (RpaInfo item in list)
                    {
                         await Task.Delay(1000, token);token.ThrowIfCancellationRequested(); 
                        if (item.DefaultStatus != "有效") { continue; }
                        if (!await IsTime(token, item.DefaultRunTime1, item.DefaultRunTime2)) { continue; }
                        this.Invoke(new Action(() => { LabMsg($"正在执行：{item.RpaName}..."); }));
                        Type type = Type.GetType(this.GetType().Namespace + "." + item.RpaClassName); if (type == null) { continue; }
                        object instance = Activator.CreateInstance(type);
                        MethodInfo method = type.GetMethod(RpaApi.RpaMain); if (method == null) { continue; }
                        await (Task)method.Invoke(instance, new object[] { token, item });
                        await Task.Delay(1000, token);token.ThrowIfCancellationRequested();                        
                    }
                    await Task.Delay(1000, token); if (token.IsCancellationRequested) { return; }                   
                    await CountDown(token,10, "休息");
                }
            }
            catch (OperationCanceledException)
            {
                LabMsg($"任务已取消");
            }
            finally
            {
                MyGlobal.MyCts.Dispose();
            }
        }


        #region 主要函数

        private async Task CountDown(CancellationToken token, int count,string Style)
        {
            for (int i = 0; i < count; i++)
            {
                token.ThrowIfCancellationRequested();                
                if (Style == "倒计时") { LabMsg($"程序将在{count-i}秒后执行流程租的流程，如取消请点击[停止]按钮"); }
                if (Style == "休息") { LabMsg($"流程租的流程已经全部执行完毕，休息{count-i}秒后将进入下一轮..."); }
                await Task.Delay(1000, token);
            }
        }

        private async Task<bool> IsTime(CancellationToken token, string t1, string t2)
        {
            bool res = false;
            try
            {
                token.ThrowIfCancellationRequested();
                DateTime d1 = Convert.ToDateTime(publicClass.ReplaceTime(t1, true));
                DateTime d2 = Convert.ToDateTime(publicClass.ReplaceTime(t2, false));
                DateTime now = DateTime.Now;
                res = now >= d1 && now <= d2;//错误的结果，待修改
            }
            catch { }
            await Task.Delay(0, token);
            return res;
        }


        private List<RpaInfo> GetRpaList_System()
        {
            List<RpaInfo> list = new List<RpaInfo>();
            Assembly assembly = Assembly.GetExecutingAssembly(); // 获取当前程序集            
            var types = assembly.GetTypes().Where(t => t.IsClass && !t.IsAbstract && t.IsSubclassOf(typeof(InvokeCenter)));// 获取所有继承自 InvokeCenter 的类型
            foreach (var type in types)
            {
                var instance = Activator.CreateInstance(type); // 创建实例              
                var method = type.GetMethod(RpaApi.GetThisRpaInfo);// 获取类中的"GetThisRpaInfo"方法
                if (method != null)
                {
                    RpaInfo rpaInfo = (RpaInfo)method.Invoke(instance, null);
                    list.Add(rpaInfo);
                }
            }
            var sortedList = list.OrderBy(x => x.RpaClassName).ToList(); //排序
            return sortedList;
        }
      
        private void SetRpaList_User()
        {
             //获取系统现存的RPA流程
            List<RpaInfo> list_sys=GetRpaList_System();           
            if (!Directory.Exists(Path.GetDirectoryName(MyPath.UserRpaList))) { Directory.CreateDirectory(Path.GetDirectoryName(MyPath.UserRpaList)); }
            if (!File.Exists(MyPath.UserRpaList))
            {
                File.Create(MyPath.UserRpaList).Close();
                using (StreamWriter writer = File.CreateText(MyPath.UserRpaList)) { writer.WriteLine(JsonConvert.SerializeObject(list_sys, Formatting.Indented)); }
                return;
            }

            //用户配置文件的RPA流程
            List<RpaInfo> list_old= GetRpaList_User();           

            //反序列化异常时设置为list_sys
            if (list_old.Count <= 0) { using (StreamWriter writer = File.CreateText(MyPath.UserRpaList)) { writer.WriteLine(JsonConvert.SerializeObject(list_sys, Formatting.Indented)); } return; }


            //修正用户配置文件
            List<RpaInfo> list_new = new List<RpaInfo>();
            foreach (RpaInfo rpaInfo in list_sys)
            {                 
                RpaInfo res= CompareRpaInfo(rpaInfo.RpaClassName,list_old);
                if (res == null) { res = rpaInfo; res.DefaultStatus = "禁用"; }
                list_new.Add(res);
            }
            using (StreamWriter writer = File.CreateText(MyPath.UserRpaList)) { writer.WriteLine(JsonConvert.SerializeObject(list_new, Formatting.Indented)); }

        }

        /// <summary>
        /// 系统现存的RPA流程对比用户配置文件的RPA流程,得出新的正确的用户配置文件
        /// </summary>
        /// <param name="rpaInfo">系统现存的RPA流程</param>
        /// <param name="list_old">用户配置文件的RPA流程 组列表</param>
        /// <returns></returns>       
        private RpaInfo CompareRpaInfo(string RpaClassName, List<RpaInfo> list_old)
        {
            RpaInfo res = null;
            foreach (RpaInfo rpaInfo_old in list_old)
            { 
                if(RpaClassName== rpaInfo_old.RpaClassName)
                {
                    res= rpaInfo_old;
                    break;
                }
            }
            return res;
        }

        private List<RpaInfo> GetRpaList_User()
        {
            //读取用户配置文件的RPA流程，前提：存在MyPath.UserRpaList文件
            List<RpaInfo> list_user=new List<RpaInfo>();
            if (File.Exists(MyPath.UserRpaList))
            {
                using (StreamReader reader = File.OpenText(MyPath.UserRpaList))
                {
                    try { list_user = JsonConvert.DeserializeObject<List<RpaInfo>>(reader.ReadToEnd()); }
                    catch { }
                    finally { reader.Close(); }
                }               
            }
            return list_user;
        }

       
        private void ShowDataGridView()
        {
            dataGridView1.DataSource = null;
            DataTable dt=new DataTable ();
            dt.Columns.Add($"{Dvg1CN.ClassName}", typeof(String));
            dt.Columns.Add($"{Dvg1CN.Name}", typeof(String));
            dt.Columns.Add($"{Dvg1CN.RunTime1}", typeof(String));
            dt.Columns.Add($"{Dvg1CN.RunTime2}", typeof(String));
            dt.Columns.Add($"{Dvg1CN.Status}", typeof(String));
            dt.Columns.Add($"{Dvg1CN.PathStype}", typeof(String));
            dt.Columns.Add($"{Dvg1CN.ConfigPath}", typeof(String));

            List<RpaInfo> list_user = GetRpaList_User();
            if (list_user.Count <= 0) { return; }
            foreach (RpaInfo rpaInfo in list_user)
            {
                if (checkBox1.Checked && rpaInfo.DefaultStatus == "无效")
                {
                    continue;
                }
                DataRow dr = dt.NewRow();
                dr[$"{Dvg1CN.ClassName}"] = rpaInfo.RpaClassName;
                dr[$"{Dvg1CN.Name}"] = rpaInfo.RpaName;
                dr[$"{Dvg1CN.RunTime1}"] = rpaInfo.DefaultRunTime1;
                dr[$"{Dvg1CN.RunTime2}"] = rpaInfo.DefaultRunTime2;
                dr[$"{Dvg1CN.Status}"] = rpaInfo.DefaultStatus;
                dr[$"{Dvg1CN.PathStype}"] = rpaInfo.DefaultPathStype;
                dr[$"{Dvg1CN.ConfigPath}"] = rpaInfo.DefaultConfigPath;
                dt.Rows.Add(dr);
            }
            dataGridView1.DataSource = dt;

            if (!(dataGridView1.DataSource != null && dataGridView1.Rows.Count > 0)) { return; }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            { 
                string DefualtStatus = dataGridView1.Rows[i].Cells[$"{Dvg1CN.Status}"].Value.ToString();
                if (DefualtStatus == "有效")
                {
                    dataGridView1.Rows[i].Cells[$"{Dvg1CN.Status}"].Style.ForeColor= Color.Green;
                }
                else 
                {
                    dataGridView1.Rows[i].Cells[$"{Dvg1CN.Status}"].Style.ForeColor = Color.Red;
                }
            }




        }

        #endregion

        #region  隐藏禁用，CheckBox1相关
        private void PleaseStopThread()
        {
            MessageBox.Show("RPA正在运行，先请停止后再来操作。","提示");
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (!IsThreadStop) { PleaseStopThread(); return; }
            ShowDataGridView();
        }
        #endregion

        #region 启动即运行，CheckBox3相关
        bool CheckBox3Loaded =false;
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {            
            if (!IsThreadStop) { PleaseStopThread();return; }
            if (!CheckBox3Loaded) {  return; }

            if (!Directory.Exists(Path.GetDirectoryName(MyPath.RunToStart))) { Directory.CreateDirectory(Path.GetDirectoryName(MyPath.RunToStart)); }
            if (!File.Exists(MyPath.RunToStart)) { File.Create(MyPath.RunToStart).Close(); }
            using (StreamWriter writer = File.CreateText(MyPath.RunToStart))
            {
                writer.WriteLine(checkBox3.Checked.ToString());
            }
            using (StreamReader reader = File.OpenText(MyPath.RunToStart))
            {
                string str = reader.ReadLine();
                bool IsChecked = (str != null && str.ToUpper() == "True".ToUpper());
                checkBox3.Checked = IsChecked;
            }
            MessageBox.Show("设置成功，下次启动时启用","提示");
        }
        private void LoadCheckBox3()
        {
            if (!IsThreadStop) { PleaseStopThread(); return; }
            if (!Directory.Exists(Path.GetDirectoryName(MyPath.RunToStart))) { Directory.CreateDirectory(Path.GetDirectoryName(MyPath.RunToStart)); }
            if (!File.Exists(MyPath.RunToStart)) { File.Create(MyPath.RunToStart).Close(); }
            using (StreamReader reader = File.OpenText(MyPath.RunToStart))
            {
                string str = reader.ReadLine();
                bool IsChecked = (str != null && str.ToUpper() == "True".ToUpper());
                checkBox3.Checked = IsChecked;
            }
            CheckBox3Loaded = true;
        }

        #endregion

        #region 开机启动相关
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (!IsThreadStop) { PleaseStopThread(); return; }
            MessageBox.Show(@"请自行将快捷方式放入cmd>>shell:startup>>目录下，这里不再管理。", "提示");
        }
        #endregion

        #region 列表的右键相关操作
        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (!(e.RowIndex != -1 && e.ColumnIndex != -1 && e.Button == MouseButtons.Right)) { return; }
            if (!IsThreadStop) { PleaseStopThread(); return; }
            dataGridView1.ClearSelection();
            dataGridView1.Rows[e.RowIndex].Selected = true;
            dataGridView1.CurrentCell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];           
            contextMenuStrip1.Show(MousePosition.X, MousePosition.Y);
        }
        private void 有效ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsThreadStop) { PleaseStopThread(); return; }
            List<RpaInfo> list_new = new List<RpaInfo>();
            int rowindex = dataGridView1.SelectedRows[0].Index;
            string RpaClassName = dataGridView1.Rows[rowindex].Cells[$"{Dvg1CN.ClassName}"].Value.ToString();
            List<RpaInfo> list_user = GetRpaList_User();
            if (list_user.Count <= 0) { return; }
            foreach (RpaInfo item in list_user)
            {
                if (RpaClassName == item.RpaClassName)
                {
                    item.DefaultStatus = "有效";
                }                
                list_new.Add(item);
            }
            //已经在GetRpaList_User()处理文件不存在问题，这里直接写入。
            using (StreamWriter writer = File.CreateText(MyPath.UserRpaList)) { writer.WriteLine(JsonConvert.SerializeObject(list_new, Formatting.Indented)); }

            ShowDataGridView();

        }

        private void 无效ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsThreadStop) { PleaseStopThread(); return; }
            List<RpaInfo> list_new = new List<RpaInfo>();
            int rowindex = dataGridView1.SelectedRows[0].Index;
            string RpaClassName = dataGridView1.Rows[rowindex].Cells[$"{Dvg1CN.ClassName}"].Value.ToString();
            List<RpaInfo> list_user = GetRpaList_User();
            if (list_user.Count <= 0) { return; }
            foreach (RpaInfo item in list_user)
            {
                if (RpaClassName == item.RpaClassName)
                {
                    item.DefaultStatus = "无效";
                }
                list_new.Add(item);
            }
            //已经在GetRpaList_User()处理文件不存在问题，这里直接写入。
            using (StreamWriter writer = File.CreateText(MyPath.UserRpaList)) { writer.WriteLine(JsonConvert.SerializeObject(list_new, Formatting.Indented)); }

            ShowDataGridView();
        }

        private void 全部有效ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsThreadStop) { PleaseStopThread(); return; }
            List<RpaInfo> list_new = new List<RpaInfo>(); 
            List<RpaInfo> list_user = GetRpaList_User();
            if (list_user.Count<=0) { return; }
            foreach (RpaInfo item in list_user)
            {
                item.DefaultStatus = "有效";
                list_new.Add(item);
            }
            //已经在GetRpaList_User()处理文件不存在问题，这里直接写入。
            using (StreamWriter writer = File.CreateText(MyPath.UserRpaList)) { writer.WriteLine(JsonConvert.SerializeObject(list_new, Formatting.Indented)); }

            ShowDataGridView();
        }

        private void 全部无效ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsThreadStop) { PleaseStopThread(); return; }
            List<RpaInfo> list_new = new List<RpaInfo>();
            List<RpaInfo> list_user = GetRpaList_User();
            if (list_user.Count <= 0) { return; }
            foreach (RpaInfo item in list_user)
            {
                item.DefaultStatus = "无效";
                list_new.Add(item);
            }
            //已经在GetRpaList_User()处理文件不存在问题，这里直接写入。
            using (StreamWriter writer = File.CreateText(MyPath.UserRpaList)) { writer.WriteLine(JsonConvert.SerializeObject(list_new, Formatting.Indented)); }

            ShowDataGridView();
        }

        private void 打开配置文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsThreadStop) { PleaseStopThread(); return; }
            int rowindex = dataGridView1.SelectedRows[0].Index;
            string DefualtPathStype = dataGridView1.Rows[rowindex].Cells[$"{Dvg1CN.PathStype}"].Value.ToString();
            string DefualtConfigPath = dataGridView1.Rows[rowindex].Cells[$"{Dvg1CN.ConfigPath}"].Value.ToString();
            string path = DefualtPathStype == "相对路径" ? MyPath.App + DefualtConfigPath : DefualtConfigPath;
            try { Process.Start(path); }
            catch (Exception ex) { MessageBox.Show(ex.Message,"错误"); }
        }

        #endregion

        # region 列表的双击相关操作
        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (!(e.RowIndex != -1 && e.ColumnIndex != -1)) { return; }
            if (!IsThreadStop) { PleaseStopThread(); return; }
            RpaInfo rpaInfo = new RpaInfo();
            rpaInfo.RpaClassName = dataGridView1.Rows[e.RowIndex].Cells[$"{Dvg1CN.ClassName}"].Value.ToString();
            rpaInfo.RpaName = dataGridView1.Rows[e.RowIndex].Cells[$"{Dvg1CN.Name}"].Value.ToString();
            rpaInfo.DefaultStatus = dataGridView1.Rows[e.RowIndex].Cells[$"{Dvg1CN.Status}"].Value.ToString();
            rpaInfo.DefaultRunTime1 = dataGridView1.Rows[e.RowIndex].Cells[$"{Dvg1CN.RunTime1}"].Value.ToString();
            rpaInfo.DefaultRunTime2 = dataGridView1.Rows[e.RowIndex].Cells[$"{Dvg1CN.RunTime2}"].Value.ToString();
            rpaInfo.DefaultPathStype = dataGridView1.Rows[e.RowIndex].Cells[$"{Dvg1CN.PathStype}"].Value.ToString();
            rpaInfo.DefaultConfigPath = dataGridView1.Rows[e.RowIndex].Cells[$"{Dvg1CN.ConfigPath}"].Value.ToString();
            RpaInfoAttr rpaInfoAttr = new RpaInfoAttr(rpaInfo);
            rpaInfoAttr.ShowDialog();
            if (rpaInfoAttr.DialogResult != DialogResult.OK) { return; }

            List<RpaInfo> list_new = new List<RpaInfo>();
            List <RpaInfo> list_user= GetRpaList_User();
            foreach (RpaInfo item in list_user)
            {
                RpaInfo temp= item;
                if (item.RpaClassName == rpaInfo.RpaClassName) { temp = rpaInfoAttr.MyRpaInfo; }
                list_new.Add(temp);
            }
            //已经在GetRpaList_User()处理文件不存在问题，这里直接写入。
            using (StreamWriter writer = File.CreateText(MyPath.UserRpaList)) { writer.WriteLine(JsonConvert.SerializeObject(list_new, Formatting.Indented)); }

            ShowDataGridView();

        }

        #endregion

        #region 其他菜单操作
        private void 重置我的配置RToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsThreadStop) { PleaseStopThread(); return; }
            try
            {
                if (File.Exists(MyPath.UserRpaList)) { File.Delete(MyPath.UserRpaList); }
                if (File.Exists(MyPath.RunToStart)) { File.Delete(MyPath.RunToStart); }
                LoadCheckBox3();
                SetRpaList_User();
                ShowDataGridView();
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }
           
        }

        private void 关于我们AToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsThreadStop) { PleaseStopThread(); return; }
            MessageBox.Show("KaizhongRPA 深圳市凯中精密技术股份有限公司 版权所有\r\n\r\n自主研发，商用必究。", "关于");
        }

        private void 关于AToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsThreadStop) { PleaseStopThread(); return; }
            MessageBox.Show("KaizhongRPA 深圳市凯中精密技术股份有限公司 版权所有\r\n\r\n自主研发，商用必究。", "关于");
        }

        private void 退出XToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AppExit(null);
        }

        private void 退出系统XToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AppExit(null);

        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            AppExit(e);
        }
        public int FClose = 0;
        public void AppExit(FormClosingEventArgs e)
        {
            if (FClose <= 0)
            {
                DialogResult result = MessageBox.Show("退出时将直接终止所有线程。", "您确认要退出吗?", MessageBoxButtons.YesNo);
                if (result == DialogResult.No)
                {
                    if (e != null) { e.Cancel = true; }
                    return;
                }
                else
                {
                    FClose += 1;
                    if (IsThreadStop = false&&MyGlobal.MyCts != null) 
                    { 
                        MyGlobal.MyCts.Cancel();
                    }                    
                    notifyIcon1.Dispose();
                    Application.Exit();
                }
            }
        }

        private void 查看本地日志LToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsThreadStop) { PleaseStopThread(); return; }
        }

        private async void 查看云端日志CToolStripMenuItem_Click(object sender, EventArgs e)
        {
             //await  publicClass.WechatPost("","");

            if (!IsThreadStop) { PleaseStopThread(); return; }
        }
                     

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SapGuiScripting sapGuiScripting = new SapGuiScripting();
            sapGuiScripting.ShowDialog();
        }

        #endregion
   
    
    }
}
