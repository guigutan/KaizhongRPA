using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace KaizhongRPA
{
    public class Rpa001 : InvokeCenter
    {
        public RpaInfo GetThisRpaInfo() 
        {
            RpaInfo rpaInfo = new RpaInfo();
            rpaInfo.RpaClassName = this.GetType().Name;
            rpaInfo.RpaName = "生产订单下达流程"; 
            rpaInfo.DefaultRunTime1 = "****-**-** **:**:**";
            rpaInfo.DefaultRunTime2 = "****-**-** **:**:**";
            rpaInfo.DefaultStatus = "无效";
            rpaInfo.DefaultPathStype = "相对路径";
            rpaInfo.DefaultConfigPath = @"config\RpaGroup\生产订单下达.xlsx";
            return rpaInfo;
        }

        PublicClass publicClass = new PublicClass();
        public async Task RpaMain(CancellationToken token, RpaInfo rpaInfo)
        {
            try
            {
                await publicClass.DisableScreen(token);
               
            }
            catch (Exception ex)
            {
               
            }
            finally
            {
                await publicClass.ExitSap(token);
            }
          
        }
    }




}
