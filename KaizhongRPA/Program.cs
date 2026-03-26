using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KaizhongRPA
{
    internal static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (!RunOnlyOne()) { return; }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        public static bool RunOnlyOne()
        {
            bool result = true;
            int processCount = 0;
            string myProcessName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
            System.Diagnostics.Process[] processArr = System.Diagnostics.Process.GetProcesses();
            foreach (System.Diagnostics.Process process in processArr)
            {
                if (process.ProcessName == myProcessName) { processCount += 1; }               
            }
            if (processCount > 1)
            {
                DialogResult dr = MessageBox.Show(myProcessName + "  程序已经在运行！", "退出程序", MessageBoxButtons.OK, MessageBoxIcon.Error);
                result= false;
            }
            return result;

        }





    }
}
