using System;
using System.Collections.Generic;
using System.Configuration.Install;
using System.Linq;
using System.Reflection;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace MailSending
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {


            // For Debug

            //var service = new MailSendingService();
            //service.onDebug();
            //System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);



            // FOR INSTALL

            //try
            //{
            //    if (Environment.UserInteractive)
            //    {
            //        string parameter = string.Concat("--uninstall");
            //        switch (parameter)
            //        {
            //            case "--install":
            //                ManagedInstallerClass.InstallHelper(new[] { Assembly.GetExecutingAssembly().Location });
            //                break;
            //            case "--uninstall":
            //                ManagedInstallerClass.InstallHelper(new[] { "/u", Assembly.GetExecutingAssembly().Location });
            //                break;
            //        }
            //    }
            //}
            //catch (Exception e)
            //{

            //}


            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new MailSendingService()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
