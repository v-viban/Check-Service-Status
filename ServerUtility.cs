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
        /// Loop through servers to check service status
        /// </summary>
        public void GetServiceStatusFromServers()
        {
            DateTime sysDate = DateTime.Now;
            TimeZoneInfo timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Indian Standard Time");
            DateTime newDateTime = TimeZoneInfo.ConvertTime(sysDate, timeZoneInfo);
            string currentDt = String.Format("{0:F}", newDateTime);
            sb.Append("<!DOCTYPE html><html lang='en' xmlns='http://www.w3.org/1999/xhtml'><head><meta charset='utf-8'/><style>table, th, td {border: 1px solid black;font-family:calibri;font-size:14px;}table {border-bottom:1;border-left:1;}td, th {border-top:1;border-right:1;} p{font-family:calibri;font-size:14px;}</style></head >");
            sb.Append("<body>");
            sb.AppendFormat("<p>Hi All, <br><br>Here is Jenkins & Wadi Service Status report for <b>{0}</b></p><br>", currentDt);

            for (int i = 0; i < servers.Length; i++)
            {
                GetRemoteServiceStatusReport(servers[i].ToString());
            }

            sb.Append("<br>");
            sb.Append("<p>Thanks,<br>Vivek");
            
            SendReportEmail(sb.ToString());
        }

        /// <summary>
        /// Get status of service state and create report
        /// ContinuePending The service has been paused and is about to continue.
        /// Paused The service is paused.
        /// PausePending The service is in the process of pausing.
        /// Running The service is running.
        /// StartPending The service is in the process of starting.
        /// Stopped The service is not running.
        /// StopPending The service is in the process of stopping.
        /// </summary>
        /// <param name="serverName"></param>
        public void GetRemoteServiceStatusReport(string serverName)
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
            //Console.WriteLine("Server Name is {0}", serverName);
           
            sb.Append("<table width='50%' cellspacing='0' cellpadding='0'>");
            sb.Append("<tr><td colspan='2' style='color:#FFFFFF;font-weight:bold' bgcolor='#191970' align='center'>" + serverName + "</td></tr>");
            sb.Append("<tr style='border:1px solid Black'>");
            sb.Append("<td style='color:#FFFFFF' bgcolor='#191970' align='center'>Service Name</td>");
            sb.Append("<td style='color:#FFFFFF' bgcolor='#191970' align='center'>Service State</td>");
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
                   // Console.WriteLine("Service Name is {0} and State is {1}", service.GetPropertyValue("Name").ToString(), service.GetPropertyValue("State").ToString());
                    sb.Append("<tr>");
                    sb.Append("<td>" + service.GetPropertyValue("Name").ToString() + "</td>");

                    // Assign color codes based on status
                    switch (service.GetPropertyValue("State").ToString())
                    {
                        case "ContinuePending":
                            sb.Append("<td style='color:#8B0000;font-weight:bold'>" + service.GetPropertyValue("State").ToString() + "</td>");
                            break;
                        case "Pause":
                            sb.Append("<td style='color:#8B0000;font-weight:bold'>" + service.GetPropertyValue("State").ToString() + "</td>");
                            break;
                        case "PausePending":
                            sb.Append("<td style='color:#8B0000;font-weight:bold'>" + service.GetPropertyValue("State").ToString() + "</td>");
                            break;
                        case "Running":
                            sb.Append("<td style='color:#008000;font-weight:bold'>" + service.GetPropertyValue("State").ToString() + "</td>");
                            break;
                        case "StartPending":
                            sb.Append("<td style='color:#8B0000;font-weight:bold'>" + service.GetPropertyValue("State").ToString() + "</td>");
                            break;
                        case "Stopped":
                            sb.Append("<td style='color:#8B0000;font-weight:bold'>" + service.GetPropertyValue("State").ToString() + "</td>");
                            break;
                        case "StopPending":
                            sb.Append("<td style='color:#8B0000;font-weight:bold'>" + service.GetPropertyValue("State").ToString() + "</td>");
                            break;
                        default:
                            sb.Append("<td style='color:#8B0000;font-weight:bold'>" + service.GetPropertyValue("State").ToString() + "</td>");
                            break;
                    }

                    sb.Append("</tr>");
                    
                }
            }

            sb.Append("</table></body></html>");
            sb.Append("<br><br>");
            
        }

        #region Send Report Email
        public void SendReportEmail(string htmlBody)
        {
            StringBuilder emailText = new StringBuilder();
            // Create a New System.Net.Mail message
            MailMessage email = new MailMessage();
            email.Body = null;
            email.From = new MailAddress(ConfigurationSettings.AppSettings["EmailFrom"].ToString());
            email.To.Add(ConfigurationSettings.AppSettings["EmailTo"].ToString());
            email.Subject = ConfigurationSettings.AppSettings["EmailSubject"].ToString();
            // SMTP Client
            SmtpClient clientSmtp = new SmtpClient();
            clientSmtp.Host = "SMTPHOST";
            clientSmtp.EnableSsl = false;
            clientSmtp.UseDefaultCredentials = true;
            email.IsBodyHtml = true;
            email.Body = htmlBody;
            clientSmtp.Send(email);
        }
        #endregion
    }
}



