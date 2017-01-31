using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceProcess;
using System.Management;
using System.Configuration;
using System.Net.Mail;

namespace CheckServiceStatus
{
    public class ServerUtility
    {
        string[] servers = ConfigurationSettings.AppSettings["Servers"].Split(',');
        StringBuilder sb = new StringBuilder();
        /// <summary>
        ///
        /// </summary>
        public void GetServiceStatusFromServers()
        {
            for (int i = 0; i < servers.Length; i++)
            {
                GetRemoteServiceStatus(servers[i].ToString());
            }

            SendReportEmail(sb.ToString());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="serverName"></param>
        public void GetRemoteServiceStatus(string serverName)
        {
            ServiceController sc = new ServiceController();
            ConnectionOptions op = new ConnectionOptions();
            op.Authority = "kerberos:redmond\\" + serverName;
            //op.Authority = "NTLMDOMAIN:DOMAIN";
            op.Username = ConfigurationSettings.AppSettings["UserName"].ToString();
            op.Password = ConfigurationSettings.AppSettings["Password"].ToString();
            op.Authentication = AuthenticationLevel.Default;
            op.EnablePrivileges = true;
            op.Impersonation = ImpersonationLevel.Impersonate;
            ManagementScope scope = new ManagementScope(@"\\" + serverName + "\\root\\cimv2", op);
            scope.Connect();
            ManagementPath path = new ManagementPath("Win32_Service");
            ManagementClass mgmtCls;
            mgmtCls = new ManagementClass(scope, path, null);

            //Run for no of services on that server
            string[] services = ConfigurationSettings.AppSettings["Services"].Split(',');
            Console.WriteLine("Server Name is {0}", serverName);

            sb.Append("<!DOCTYPE html><html lang='en' xmlns='http://www.w3.org/1999/xhtml'><head><meta charset='utf-8'/><style>table, th, td {border: 1px solid black;font-family:calibri;font-size:14px;}table {border-bottom:0;border-left:0;}td, th {border-top:0;border-right:0;} p{font-family:calibri;font-size:14px;}</style></head >");
            sb.Append("<body><table style='border:1px solid Black;width:50%'>");
            sb.Append("<tr style='border:1px solid Black'><td colspan='2' align='center' style='border:1px solid Black;background-color:darkslateblue;font-family:'Gill Sans', 'Gill Sans MT', Calibri, 'Trebuchet MS', sans-serif;font-size:smaller;color:white;font-weight:bold'>" + serverName + "</td></tr>");
            sb.Append("<tr style='border:1px solid Black'>");
            sb.Append("<td style='border:1px solid Black;background-color:darkslateblue;font-family:'Gill Sans', 'Gill Sans MT', Calibri, 'Trebuchet MS', sans-serif;font-size:smaller;color:white;font-weight:bold'>Service Name</td>");
            sb.Append("<td style='border:1px solid Black;background-color:darkslateblue;font-family:'Gill Sans', 'Gill Sans MT', Calibri, 'Trebuchet MS', sans-serif;font-size:smaller;color:white;font-weight:bold'>Service State</td>");
            sb.Append("</tr>");

            for (int j = 0; j < services.Length; j++)
            {
                string sql = string.Format("SELECT * From Win32_Service WHERE Name like '" + services[j].ToString() + "%'");

                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, new ObjectQuery()
                {
                    QueryString = sql
                });

                foreach (ManagementObject service in searcher.Get())
                {
                    Console.WriteLine("Service Name is {0} and State is {1}", service.GetPropertyValue("Name").ToString(), service.GetPropertyValue("State").ToString());
                    sb.Append("<tr style='border:1px solid Black>");
                    sb.Append("<td style='border: 1px solid Black; font-family:'Gill Sans','Gill Sans MT',Calibri,'Trebuchet MS',sans-serif;font-size:smaller; color:black'>" + service.GetPropertyValue("Name").ToString() + "</td>");
                    sb.Append("<td style='border: 1px solid Black; font-family:'Gill Sans','Gill Sans MT',Calibri,'Trebuchet MS',sans-serif;font-size:smaller; color:black'>" + service.GetPropertyValue("State").ToString() + "</td>");
                    sb.Append("</tr>");

                }
            }

            sb.Append("</table></body></html>");
            sb.Append("<br><br>");
            
        }

        public void CreateHtmlReport(string serverName)
        {
            
        }

        public void SendReportEmail(string htmlBody)
        {

            // Create a New System.Net.Mail message
            MailMessage email = new MailMessage();
            email.Body = null;
            email.From = new MailAddress("v-viban@microsoft.com");
            email.To.Add("v-viban@microsoft.com");
            email.Subject = "Service Status Report";

            // SMTP Client
            SmtpClient clientSmtp = new SmtpClient();
            clientSmtp.Host = "SMTPHOST";
            //clientSmtp.Port = 25;
            clientSmtp.EnableSsl = false;
            clientSmtp.UseDefaultCredentials = true;
            email.IsBodyHtml = true;
            email.Body = htmlBody;
            clientSmtp.Send(email);
            
        }

    }
}
