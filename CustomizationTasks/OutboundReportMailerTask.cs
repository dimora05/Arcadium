using System;
using System.IO;
using System.Net.Mail;
using System.Text;
using Thermo.SampleManager.Library;
using Thermo.SampleManager.Server;

namespace Customization.Tasks
{
    [SampleManagerTask("OutboundReportMailerTask")]
    public class OutboundReportMailerTask : SampleManagerTask
    {
        protected override void SetupTask()
        {
            try
            {
                string to = "danny.loria@thermofisher.com";
                string subject = "Pruebatask";
                string body = "mensajetask";

                // Obtener configuración SMTP de Globals
                string host = Library.Environment.GetGlobalString("SMTP_HOST");
                int port = Library.Environment.GetGlobalInt("SMTP_PORT");
                bool enableSsl = Library.Environment.GetGlobalBoolean("SMTP_ENABLESSL");
                string username = Library.Environment.GetGlobalString("SMTP_USERNAME");
                string password = Library.Environment.GetGlobalString("SMTP_PASSWORD");
                string sender = Library.Environment.GetGlobalString("SMTP_SENDER");

                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(string.IsNullOrWhiteSpace(sender) ? "Sample.Manager@arcadiumlithium.com" : sender);
                mail.Sender = mail.From;
                mail.To.Add(to);
                mail.Subject = subject;
                mail.Body = body;
                mail.SubjectEncoding = Encoding.UTF8;
                mail.BodyEncoding = Encoding.UTF8;

                SmtpClient client = !string.IsNullOrEmpty(host) ? new SmtpClient(host, port) : new SmtpClient();
                if (!string.IsNullOrEmpty(username))
                {
                    client.UseDefaultCredentials = false;
                    client.Credentials = new System.Net.NetworkCredential(username, password);
                }
                client.EnableSsl = enableSsl || (!string.IsNullOrEmpty(host) && host.IndexOf("smtp2go.com", StringComparison.OrdinalIgnoreCase) >= 0);

                client.Send(mail);
                Library.Utils.FlashMessage("Correo enviado (OutboundReportMailerTask)", "OK");
            }
            catch (Exception ex)
            {
                Library.Utils.FlashMessage($"Error enviando correo: {ex.Message}", "Error");
            }
        }
    }
}
