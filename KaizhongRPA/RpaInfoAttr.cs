using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace KaizhongRPA
{
    public partial class RpaInfoAttr : Form
    {
        PublicClass publicClass = new PublicClass();
        public RpaInfoAttr(RpaInfo rpaInfo)
        {
            InitializeComponent();
            MyRpaInfo=rpaInfo;//防NULL
            LoadInfo();

        }
        public RpaInfo MyRpaInfo { get; set; }
        private void LoadInfo()
        {
            lab_ClassName.Text = Dvg1CN.ClassName;
            lab_Name.Text = Dvg1CN.Name;
            lab_RunTime1.Text = Dvg1CN.RunTime1;
            lab_RunTime2.Text = Dvg1CN.RunTime2;
            lab_PathStype.Text=Dvg1CN.PathStype;
            lab_ConfigPath.Text = Dvg1CN.ConfigPath;

            txt_ClassName.Text = MyRpaInfo.RpaClassName;
            txt_Name.Text=MyRpaInfo.RpaName;
            txt_RunTime1.Text = MyRpaInfo.DefaultRunTime1;
            txt_RunTime2.Text = MyRpaInfo.DefaultRunTime2;
            txt_ConfigPath.Text=MyRpaInfo.DefaultConfigPath;           

            for (int i=0;i<cb_PathStype.Items.Count;i++ )
            {
                string itemText = cb_PathStype.Items[i].ToString();
                if (itemText == MyRpaInfo.DefaultPathStype)
                {                    
                    cb_PathStype.SelectedIndex = i;
                    break; 
                }
            }
            string Demo = "格式：4位年分-两位月-两位日 两位时:两位分:两位秒\r\n\r\n";
            Demo += "通配符*代表任意，举例子：\r\n\r\n";
            Demo += "   从：****-**-** 09:00:00 ";
            Demo += "   至：****-**-** 12:00:00 \r\n\r\n";
            Demo += "   表示：在任意的年月日的9点到12点的这段时间，运行这条RPA流程。";
            label1.Text = Demo;
            label2.Text = $"当前相对路径：{MyPath.App}";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            UpdInfo();
        }

        private void UpdInfo()
        {
            //校验时间格式
            string t1 = publicClass.ReplaceTime(txt_RunTime1.Text.Trim(), true);
            try { Convert.ToDateTime(t1); } catch { MessageBox.Show("运行条件(从)，时间格式错误！", "提示");return; }
            string t2 = publicClass.ReplaceTime(txt_RunTime2.Text.Trim(), false);
            try { Convert.ToDateTime(t2); } catch { MessageBox.Show("运行条件(至)，时间格式错误！", "提示"); return; }

            string PathStype = "相对路径";
            if (cb_PathStype.SelectedIndex != -1) { PathStype = cb_PathStype.SelectedItem.ToString(); }

            RpaInfo newInfo = new RpaInfo();
            newInfo.RpaClassName= MyRpaInfo.RpaClassName;
            newInfo.RpaName= MyRpaInfo.RpaName;
            newInfo.DefaultRunTime1 = txt_RunTime1.Text.Trim().Replace("：",":");
            newInfo.DefaultRunTime2 = txt_RunTime2.Text.Trim().Replace("：", ":");
            newInfo.DefaultStatus= MyRpaInfo.DefaultStatus;
            newInfo.DefaultPathStype = PathStype;
            newInfo.DefaultConfigPath = txt_ConfigPath.Text.Trim();

            MyRpaInfo = newInfo;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

    }
}
