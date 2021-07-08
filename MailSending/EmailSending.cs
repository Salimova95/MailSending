using MailKit.Net.Smtp;
using MailSending.Models;
using MimeKit;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailSending
{
    class EmailSending
    {
        async public static void SendMail(Dictionary<String, Object> ContentParameters, OracleConnection oracleConnection)
        {
            Config config = Config.One(oracleConnection);
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("Mail.GTS", config.MailUsername));
            message.To.Add(new MailboxAddress(ContentParameters["mailTo"].ToString(), ContentParameters["mailTo"].ToString()));
            message.Subject = "Mail Notification";

            var footer = System.IO.File.ReadAllText(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MailTemplates/footer.html"));

            var header = System.IO.File.ReadAllText(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MailTemplates/header.html"));

            var appeal = System.IO.File.ReadAllText(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MailTemplates/Mail-info.html"));

            var checkout = System.IO.File.ReadAllText(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MailTemplates/Mail.html"));

            String mainContent = string.Format(checkout, header, appeal, footer);
            var bodyBuilder = new BodyBuilder();
            bodyBuilder.HtmlBody = mainContent;
            message.Body = bodyBuilder.ToMessageBody();


            using (var client = new SmtpClient())
            {
                try
                {
                   
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                    await client.ConnectAsync(config.MailHost, 25); //проверить порт в базе

                    client.AuthenticationMechanisms.Remove("XOAUTH2");
                    await client.AuthenticateAsync(config.MailUsername, config.MailPassword);


                    await client.SendAsync(message);

                    await client.DisconnectAsync(true);


                }
                catch (Exception e)
                {
                    System.Diagnostics.EventLog.WriteEntry("NOT_SENT", e.Message);

                }

            }

        }
    }
}
