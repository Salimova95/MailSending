using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using Oracle.ManagedDataAccess.Client;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailSending.Models 
{
    class Config
    {
        public String MailUsername { get; set; }

        public String MailPassword { get; set; }

        public String MailHost { get; set; }

        public int MailPort { get; set; }


        public static Config One(OracleConnection oracleConnection)
        {
            OracleCommand configQuery = new OracleCommand(@"SELECT * FROM ""Config"" WHERE ID = 1", oracleConnection);
            Config config = null;

            var configData = configQuery.ExecuteReader();

            if (configData.HasRows)
            {
                configData.Read();

                config = new Config();
                config.MailUsername = (String)configData["MAILUSERNAME"];
                config.MailPassword = (String)configData["MAILPASSWORD"];
                config.MailHost = (String)configData["MAILHOST"];
                config.MailPort = Convert.ToInt32(configData["MAILPORT"]);

            }

            configData.Close();

            return config;
        }
    }
}
