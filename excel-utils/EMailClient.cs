using excel_utils.Models;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text.RegularExpressions;

namespace excel_utils
{
    public class EMailClient
    {
        private char emailSeparator = ';';
        private string emailDomain = "@company.com";
        private string emailServer = "smtp.contoso.com";
        private string emailMatcher = @"^[a-z][a-z|0-9|]*([_][a-z|0-9]+)*([.][a-z|"
                                    + @"0-9]+([_][a-z|0-9]+)*)?@[a-z][a-z|0-9|]*\.([a-z]"
                                    + @"[a-z|0-9]*(\.[a-z][a-z|0-9]*)?)$";
        private string emailFileType = @"[a-zA-Z0-9]*.emsg";
        private MSGSetting msg;

        public EMailClient(MSGSetting msg)
        {
            this.msg = msg;

            if (!string.IsNullOrEmpty(msg.To))
            {
                if (msg.From != null && msg.From.ToLower().Equals(Environment.UserName.ToLower()))
                {
                    msg.From = msg.From + emailDomain;
                }
            }
        }

        public bool SendEMail()
        {
            try
            {
                var smtp = new SmtpClient
                {
                    Host = emailServer,
                    Port = 25,
                    EnableSsl = false,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = true,
                    Credentials = new NetworkCredential()
                };
                using (var message = new MailMessage())
                {
                    Match match;

                    match = Regex.Match(msg.From, emailMatcher, RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        message.From = new MailAddress(msg.From);
                    }

                    foreach (string idTo in msg.To.Split(emailSeparator))
                    {
                        match = Regex.Match(idTo, emailMatcher, RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            message.To.Add(idTo);
                        }
                        else
                        {
                            string tempMailid = GetMailId(idTo);
                            foreach (string idTof in tempMailid.Split(emailSeparator))
                            {
                                message.To.Add(idTof);
                            }
                        }
                    }
                    if (msg.Cc != null && !msg.Cc.Equals(string.Empty))
                    {
                        foreach (string idCc in msg.Cc.Split(emailSeparator))
                        {
                            match = Regex.Match(idCc, emailMatcher, RegexOptions.IgnoreCase);
                            if (match.Success)
                            {
                                message.CC.Add(idCc);
                            }
                            else
                            {
                                string tempMailid = GetMailId(idCc);
                                foreach (string idCCf in tempMailid.Split(emailSeparator))
                                {
                                    message.CC.Add(idCCf);
                                }
                            }
                        }
                    }
                    message.Subject = msg.Subject;
                    if (Regex.Match(msg.Body, emailFileType, RegexOptions.IgnoreCase).Success)
                    {
                        StreamReader file = new StreamReader(msg.Body);
                        msg.Body = file.ReadToEnd();
                    }

                    if (msg.Attch != null && !msg.Attch.Equals(string.Empty))
                    {
                        foreach (string file in msg.Attch.Split(emailSeparator))
                        {
                            Attachment oAttch = new Attachment(file);
                            message.Attachments.Add(oAttch);
                        }
                    }
                    if (!string.IsNullOrEmpty(msg.Importance))
                    {
                        message.Priority = (MailPriority)Enum.Parse(typeof(MailPriority), msg.Importance);
                    }

                    //body = "<font face='Arial' size='10'>" + body + "</font>";

                    message.IsBodyHtml = true;
                    ContentType contentType = new ContentType(MediaTypeNames.Text.Html);
                    using (AlternateView altView = AlternateView.CreateAlternateViewFromString(msg.Body, contentType))
                    {
                        message.AlternateViews.Add(altView);
                        smtp.Send(message);
                    }
                }
                return true;
            }
            catch (Exception)
            {
                //Console.WriteLine(ex.Message);
                throw;
            }
        }

        private string GetMailId(string fileName)
        {
            string mailId = string.Empty;
            string line;

            // Read the file and display it line by line.
            StreamReader file = new StreamReader(fileName);

            while ((line = file.ReadLine()) != null)
            {
                mailId = mailId + line + emailSeparator;
            }

            mailId = mailId.Remove(mailId.Length - 1, 1);
            file.Close();

            return mailId;
        }
    }
}
