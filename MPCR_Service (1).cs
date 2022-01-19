using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using Oracle.ManagedDataAccess.Client;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using MPCR_Service.Models;
using System.IO;

namespace MPCR_Service
{
    public partial class MPCR_Service : ServiceBase
    {
        String connectionString;
        OracleConnection oracleConnection;
        DateTime todayDate;
        ElapsedEventArgs e;
        String serviceName;
        string[] args;
        Timer timer;
        bool EmailSended7;
        bool CancelEmailSended;
        bool TermEmailSended;

        public MPCR_Service()
        {
            serviceName = "MPCR Email and Update Service";
            this.AutoLog = true;
            InitializeComponent();
            System.Diagnostics.EventLog.WriteEntry(serviceName, "mpcr bit system = " + IntPtr.Size);
            string path = ConfigurationManager.AppSettings["LogFile"];
            if (!File.Exists(path))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("On timer for MPCR");
                }
            }
        }

        protected override void OnStart(string[] args)
        {
            todayDate = DateTime.Today;
            timer = new System.Timers.Timer();
            timer.Interval = 120 * 1000;
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();
            System.Diagnostics.EventLog.WriteEntry(serviceName, "On timer for MPCR");
            connectionString = ConfigurationManager.AppSettings["ConnectionString"];
        }

        //protected override void OnStop()
        //{

        //}

        //public void onDebug()
        //{

        //    OnStart(args);
        //    OnTimer(connectionString, e);
        //}

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs e)
        {

            oracleConnection = new OracleConnection(connectionString);
            try
            {
                oracleConnection.Open();
            }
            catch (Exception ex)
            {
                var path = ConfigurationManager.AppSettings["LogFile"];

                using (StreamWriter streamWriter = new StreamWriter(path, true))
                {
                    streamWriter.WriteLine(DateTime.Now.ToString() + " - oracleConnection.Open(); " + ex.Message);
                }
            }


            if (todayDate != DateTime.Today && (todayDate.Day != DateTime.Today.Day))
            {
                System.Diagnostics.EventLog.WriteEntry(serviceName, "new day began");
                todayDate = DateTime.Today;

                EmailSended7 = false;
                CancelEmailSended = false;
                TermEmailSended = false;
            }

            DateTime currentTime = DateTime.Now;

            // всем залогоДЕРЖАТЕЛЯМ за 7 дней ДО автопрекращения
            if ((currentTime.Hour == 14 && currentTime.Minute == 00) || (currentTime.Hour == 14 && currentTime.Minute > 00))
            {
                try
                {
                    if (EmailSended7 != true)
                    {
                        System.Diagnostics.EventLog.WriteEntry(serviceName, "run notification service for Secured creditor (owner) 7 days before");

                        // всем залогоДЕРЖАТЕЛЯМ за 7 дней ДО
                        var encumbrances = Encumbrances.OwnersList(oracleConnection);


                        if (encumbrances.HasRows)
                        {
                            while (encumbrances.Read())
                            {
                                Encumbrances encumbrance = new Encumbrances(encumbrances);
                                notificateUsers(encumbrance, "mailFor7DaysNotify");
                                EmailSended7 = true;
                            }
                        }
                    }
                }
                catch (Exception exception)
                {
                    var path = ConfigurationManager.AppSettings["LogFile"];

                    using (StreamWriter streamWriter = new StreamWriter(path, true))
                    {
                        streamWriter.WriteLine(DateTime.Now.ToString() + " - mail 7 days before - " + exception.Message);
                    }
                }
            }

            //автопрекращение и мэил
            else if ((currentTime.Hour == 1 && currentTime.Minute == 00) || (currentTime.Hour == 1 && currentTime.Minute > 00))
            {
                System.Diagnostics.EventLog.WriteEntry(serviceName, "run update auto termination");
                try
                {
                    var encumbrances = Encumbrances.EncumUsersList(oracleConnection);

                    var autoTerminations = Encumbrances.AutoTerminationList(oracleConnection);

                    if (autoTerminations.HasRows)
                    {
                        while (autoTerminations.Read())
                        {
                            Encumbrances encumbrance = new Encumbrances(autoTerminations);
                            if (encumbrance.UpperAccountEncumbranceId == 0 || encumbrance.UpperAccountEncumbranceId == null)
                            {
                                encumbrance.AutoTermination(oracleConnection);
                            }
                            else
                            {
                                encumbrance.UpperAutoTermination(oracleConnection);
                            }
                        }

                        if (TermEmailSended != true)
                        {
                            if (encumbrances.HasRows)
                            {
                                while (encumbrances.Read())
                                {
                                    Encumbrances encumbrance = new Encumbrances(encumbrances);

                                    notificateUsers(encumbrance, "mailForAutoTermination");
                                    TermEmailSended = true;
                                }
                            }
                        }
                    }
                }
                catch (Exception exception)
                {
                    var path = ConfigurationManager.AppSettings["LogFile"];

                    using (StreamWriter streamWriter = new StreamWriter(path, true))
                    {
                        streamWriter.WriteLine(DateTime.Now.ToString() + " - AutoTermination - " + exception.Message);
                    }
                }
            }

            //автоотмена и мэил
            else if ((currentTime.Hour == 00 && currentTime.Minute == 00) || (currentTime.Hour == 00 && currentTime.Minute > 00))
            {
                try
                {
                    System.Diagnostics.EventLog.WriteEntry(serviceName, "run update auto cancel and send mail");

                    var encumbrances = Encumbrances.AutoCanselUsersList(oracleConnection);

                    var autoCancels = Encumbrances.AutoCanselList(oracleConnection);

                    if (autoCancels.HasRows)
                    {
                        while (autoCancels.Read())
                        {
                            Encumbrances encumbrance = new Encumbrances(autoCancels);
                            encumbrance.AutoCancel(oracleConnection);
                        }

                        if (CancelEmailSended != true)
                        {
                            if (encumbrances.HasRows)
                            {
                                while (encumbrances.Read())
                                {
                                    Encumbrances encumbrance = new Encumbrances(encumbrances);

                                    notificateUsers(encumbrance, "mailForAutoCancel");
                                    CancelEmailSended = true;
                                }
                            }
                        }

                    }
                }
                catch (Exception exception)
                {
                    var path = ConfigurationManager.AppSettings["LogFile"];

                    using (StreamWriter streamWriter = new StreamWriter(path, true))
                    {
                        streamWriter.WriteLine(DateTime.Now.ToString() + " - AutoCancel - " + exception.Message);
                    }
                }
            }

            oracleConnection.Close();
        }

        private void notificateUsers(Encumbrances encumbrance, string type)
        {
            try
            {
                Dictionary<String, Object> MailParams = new Dictionary<String, Object>();

                if (IsValidEmail(encumbrance.CustomerEmail))
                {

                    if (encumbrance.EncumUserTypeId == 1)
                    {
                        var GivenPersonList = Encumbrances.GivenList(oracleConnection, encumbrance.AccountEncumbrencesId);
                        var DebtorPersonList = Encumbrances.DebtorList(oracleConnection, encumbrance.AccountEncumbrencesId);
                        List<string> GivenPerson = new List<string>();
                        List<string> DebtorPerson = new List<string>();

                        if (GivenPersonList.HasRows)
                        {
                            while (GivenPersonList.Read())
                            {
                                Encumbrances encum = new Encumbrances(GivenPersonList);
                                GivenPerson.Add(encum.Given);
                            }
                        }

                        if (DebtorPersonList.HasRows)
                        {
                            while (DebtorPersonList.Read())
                            {
                                Encumbrances encum = new Encumbrances(DebtorPersonList);
                                DebtorPerson.Add(encum.Debtor);
                            }
                        }

                        var GivenResult = string.Join(", ", GivenPerson.Select(s => string.Format("{0}", s)));
                        var DebtorResult = string.Join(", ", DebtorPerson.Select(s => string.Format("{0}", s)));

                        MailParams.Add("Givens", GivenResult);
                        MailParams.Add("Debtors", DebtorResult);
                    }

                    if (encumbrance.Lang != null)
                    {
                        MailParams.Add(encumbrance.Lang.ToUpper(), encumbrance.Lang.ToUpper());
                    }
                    else
                    {
                        if (encumbrance.UserTypeId == 2 || encumbrance.UserTypeId == 4)
                        {
                            MailParams.Add("AZ", "AZ");
                        }
                    }

                    MailParams.Add("mailTo", encumbrance.CustomerEmail);
                    MailParams.Add("Fullname", encumbrance.Fullname);
                    MailParams.Add("DocNo", encumbrance.AccountEncumbrencesId);
                    MailParams.Add("ReliabilityDate", encumbrance.ReliabilityDate.ToString("dd.MM.yyyy"));
                    EmailSending.SendMail(type, MailParams, oracleConnection);
                    System.Diagnostics.EventLog.WriteEntry(serviceName, "sent notification mail for " + encumbrance.Fullname);
                }
                else
                {
                    System.Diagnostics.EventLog.WriteEntry(serviceName, "the value of Email is null, empty or incorrect - " + encumbrance.CustomerEmail + " accountEncumbranceId - " + encumbrance.AccountEncumbrencesId);

                    var path = ConfigurationManager.AppSettings["LogFile"];
                    using (StreamWriter streamWriter = new StreamWriter(path, true))
                    {
                        streamWriter.WriteLine(DateTime.Now.ToString() + " - the value of Email is null, empty or incorrect - " + encumbrance.CustomerEmail + " accountEncumbranceId - " + encumbrance.AccountEncumbrencesId);
                    }
                }
            }
            catch (Exception ex)
            {
                var path = ConfigurationManager.AppSettings["LogFile"];
                using (StreamWriter streamWriter = new StreamWriter(path, true))
                {
                    streamWriter.WriteLine(DateTime.Now.ToString() + " - notificateUsers - " + ex.Message);
                }
            }
        }

        public bool IsValidEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return false;

            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }


    }
}
