using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace MailSending
{
    public partial class MailSendingService : ServiceBase
    {
        String connectionString;
        OracleConnection oracleConnection;
        String serviceName;
        DateTime todayDate;
        bool EmailSended;
        ElapsedEventArgs e;
        string[] args;
        Timer timer;

        public MailSendingService()
        {
            serviceName = "MailSendingService";
            this.AutoLog = false;
            InitializeComponent();
            System.Diagnostics.EventLog.WriteEntry(serviceName, "MailService bit system = " + IntPtr.Size);
            connectionString = ConfigurationManager.AppSettings["ConnectionString"];
        }

        //public void onDebug()
        //{

        //    OnStart(args);
        //    OnTimer(connectionString, e);
        //}

        protected override void OnStart(string[] args)
        {
            todayDate = DateTime.Today;
            timer = new System.Timers.Timer();
            timer.Interval = 600 * 1000;
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();
            System.Diagnostics.EventLog.WriteEntry(serviceName, "On timer for MailService");
        }

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs e)
        {
            oracleConnection = new OracleConnection(connectionString);
            oracleConnection.Open();
            if (todayDate != DateTime.Today && (todayDate.Day != DateTime.Today.Day))
            {
                System.Diagnostics.EventLog.WriteEntry(serviceName, "new day began");
                todayDate = DateTime.Today;
                EmailSended = false;
            }

            DateTime currentTime = DateTime.Now;

            if ((currentTime.Hour == 16 && currentTime.Minute == 00) || (currentTime.Hour == 16 && currentTime.Minute > 00))
            {
                if (EmailSended != true)
                {
                    System.Diagnostics.EventLog.WriteEntry(serviceName, "run notification service");
                    Dictionary<String, Object> MailParams = new Dictionary<String, Object>();


                    // SET YOUR MAIL ADRESS
                    MailParams.Add("mailTo", "aliya.salimova.95@mail.ru");
                    EmailSending.SendMail(MailParams, oracleConnection);
                    EmailSended = true;
                }
            }

            oracleConnection.Close();
        }

        protected override void OnStop()
        {
        }
    }
}
