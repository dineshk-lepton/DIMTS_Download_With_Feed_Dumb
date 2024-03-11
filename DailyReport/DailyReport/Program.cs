using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace DailyReport
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            //Uncomment when you are doing debugging
            #region
            //Service1 SERVICE = new Service1();
            //SERVICE.Start();
            //System.Threading.Thread.Sleep(5000000);
            #endregion

            //Uncomment when you want to publish for production
            #region
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
                {
                 new Service1()
                };
            ServiceBase.Run(ServicesToRun);
            #endregion
        }
    }
}
