using System;
using System.Net.Mail;
using System.Net.Mime;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text;
using System.Security.Cryptography;


namespace Automate_AR
{
    public enum EmailType
    {
        Reminder,
        Due
    }

    public class Mailer
    {
        private SmtpClient smtpClient;

        public Mailer()
        {
            // Setup SMTP Client.

            smtpClient = new SmtpClient();
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = new System.Net.NetworkCredential("accounting@augustmack.com", "dusty5Road*");
            smtpClient.Host = "smtp.office365.com";
            smtpClient.Port = 587;
            smtpClient.EnableSsl = true;
        }

        // Send an email with the given content.
        public void SendMail(string subject, string to, string body, byte[] attachment, List<String> ccs = null)
        {
            if (attachment == null)
                AutomateAR.eventLog.WriteEntry("Null", EventLogEntryType.Information, AutomateAR.eventId++);

            MemoryStream stream = null;
            Attachment data = null;

            if (attachment != null)
            {
                stream = new MemoryStream(attachment);
                data = new Attachment(stream, "Invoice.pdf");
            }

            MailMessage message = CreateEmailMessage(subject, to, body, data, ccs);
            Send(message);
        }

        public string GenerateEmail(EmailType type, string contact = "")
        {
            // Figure out which email template to load.
            string file = "";
            if (type == EmailType.Reminder)
                file = @"C:\AutomateAR\reminder.html";
            else
                file = @"C:\AutomateAR\due.html";

            StreamReader reader = File.OpenText(file);
            Regex regex = new Regex(@"\$\{.*\}");


            // Read template.

            string line;
            string contents = "";

            while ((line = reader.ReadLine()) != null)
            {
                // Match regex for custom html inline variable tags.
                Match match = regex.Match(line);

                if (match.Success)
                {
                    string m = match.Groups[0].Value;

                    if (m.Trim().Equals("${contact}", StringComparison.OrdinalIgnoreCase))
                    {
                        line = regex.Replace(line, contact);
                    }
                }

                contents += line + '\n';
            }

            contents = contents.Trim();

            return contents;
        }

        // Utility method to create the actual email message that will be delivered.
        private MailMessage CreateEmailMessage(string subject, string to, string body, Attachment attachment = null, List<String> ccs = null)
        {
            MailMessage mailMessage = new MailMessage();
            mailMessage.To.Add(to);
            mailMessage.From = new MailAddress("accounting@augustmack.com", "Accounting", System.Text.Encoding.UTF8);
            mailMessage.IsBodyHtml = true;
            mailMessage.Subject = subject;
            mailMessage.Body = body;
            mailMessage.Priority = MailPriority.Normal;
            mailMessage.BodyEncoding = System.Text.Encoding.UTF8;
            mailMessage.SubjectEncoding = System.Text.Encoding.UTF8;

            foreach (string c in ccs)
            {
                mailMessage.CC.Add(new MailAddress(c));
            }

            // Add attachments too.
            if (attachment != null)
                mailMessage.Attachments.Add(attachment);

            return mailMessage;
        }

        private void Send(MailMessage message)
        {
            try
            {
                smtpClient.Send(message);
            } catch (Exception ex)
            {
                AutomateAR.eventLog.WriteEntry("Unable to send mail: " + ex.Message, System.Diagnostics.EventLogEntryType.Error);
            }
        }
    }
}
