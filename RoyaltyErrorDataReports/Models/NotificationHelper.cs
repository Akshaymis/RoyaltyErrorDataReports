using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;

namespace RoyaltyErrorDataReports.Models
{
    public static class NotificationHelper
    {


        public static bool SendMail(string toAddress, string subject, string mailContent, bool IsBodyHtml,  List<System.Net.Mail.Attachment> lstAttachment= null)
        {
            bool result = false;
 
            SmtpClient smtpClient = new SmtpClient();
            MailMessage message = new MailMessage();


            try
            {
                smtpClient = GetSMTPClientObject();
                //smtpClient.Host = EmailConfiguration.SMTPHost; 
                //smtpClient.Port =Convert.ToInt32( EmailConfiguration.SMTPPort); 
                //smtpClient.UseDefaultCredentials = EmailConfiguration.UseDefaultCredentials;
                //smtpClient.Timeout = 60000;
                //smtpClient.Credentials = new NetworkCredential(EmailConfiguration.SenderAddress, EmailConfiguration.EmailPassword);  
                //smtpClient.EnableSsl = Convert.ToBoolean(EmailConfiguration.SmtpEnableSsl);
                MailAddress senderAddress = new MailAddress("amt@bentex.com");
                message.From = new MailAddress("\"" + "" + "\" <" + senderAddress + ">");
                //message = GetEmails(toAddress, message);
                message.To.Add(toAddress);
                message.Subject = subject;
                message.IsBodyHtml = IsBodyHtml;
                message.Body = mailContent;

               
                message.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure | DeliveryNotificationOptions.OnSuccess;
                foreach (Attachment a in lstAttachment)
                {
                    message.Attachments.Add(a);

                }

                if (!String.IsNullOrEmpty(toAddress))
                {
                    smtpClient.Send(message);
                }
                else
                {

                    return false;

                }
                message.Dispose();
                result = true;


            }
            catch (Exception ex)
            {
                message.Dispose();
                ////Try to send using Backup SMTP In the case if SMTP failes
                //if (!SendUsingBackupSMTP(message))
                //{
                result = false;
                //    // GNF.LogFailedEmail(toAddress);
                //}

            }


            return result;
        }

        private static SmtpClient GetSMTPClientObject()
        {

            SmtpClient smtpClient = new SmtpClient();

            smtpClient.Host = "192.168.169.36";// ConfigurationManager.AppSettings["SMTPHostCabanaLife"];
            smtpClient.Port = 25;// Convert.ToInt32(ConfigurationManager.AppSettings["SMTPPortCabanaLife"]);
            smtpClient.UseDefaultCredentials = false;// EmailConfiguration.UseDefaultCredentials;
            smtpClient.Timeout = 60000;
            smtpClient.Credentials = new NetworkCredential("amt@bentex.com", "");
            smtpClient.EnableSsl = Convert.ToBoolean(EmailConfiguration.SSLCabanaLife);

            return smtpClient;
        }

    }

    public class EmailConfiguration
    {
        public static string EditorEmail = "";
        public static string SenderAddress = ConfigurationManager.AppSettings["SenderAddress"];
        public static string EmailPassword = ConfigurationManager.AppSettings["EmailPassword"];
        public static string SMTPHost = ConfigurationManager.AppSettings["SMTPHost"];
        public static string SMTPPort = ConfigurationManager.AppSettings["SMTPPort"];
        public static string EmailBlindCopy = "";// "akshaymishra.babu@gmail.com";
        public static string EmailCopy = "";
        public static string Get_LogoUrl = "";
        public static bool IsApplicationUnderDevelopment = false;
        public static object SmtpEnableSsl = true;
        public static object SSLCabanaLife = ConfigurationManager.AppSettings["SSLCabanaLife"];
        public static bool UseDefaultCredentials = false;

        public static string EmailReplyToCabana = "";// "<akshaymishra.babu@gmail.com>";
        public static string EmailReplyToRuby = "";//"<akshaymishra.babu@gmail.com>";

        public static string EmailErrorTo = "";// "akshaymishra.babu@gmail.com";
    }
}