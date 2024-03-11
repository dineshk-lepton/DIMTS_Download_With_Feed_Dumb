
using DailyReport.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace DailyReport
{
    public partial class Service1 : ServiceBase
    {
        DimtsReport railservice = null;
        static System.Timers.Timer process;
        public Service1()
        {
            InitializeComponent();
            railservice = new DimtsReport();
            process = new System.Timers.Timer();
        }

        public void  Start()
        {
            try
            {
                if (railservice == null)
                    railservice = new DimtsReport();
                railservice.StartVoidProcess();
            }
            catch(Exception ex) { }
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                process.Interval = 5000;
                process.Elapsed += new ElapsedEventHandler(_timer_Elapsed);
                process.Enabled = true;
                process.Start();
            }
            catch (Exception ex) { }

            
        }

        void _timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                process.Enabled = false;
                process.Stop();
                if (railservice == null)
                    railservice = new DimtsReport();
                railservice.StartVoidProcess();

            }
            catch { }
        }

        protected override void OnStop()
        {

        }
    }
}
