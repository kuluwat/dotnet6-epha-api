
using Microsoft.Exchange.WebServices.Data;

namespace Class
{
    public class mailtest
    {
        public class sendEmailModel
        {
            public string mail_from { get; set; }
            public string mail_to { get; set; }
            public string mail_cc { get; set; }
            public string mail_subject { get; set; }
            public string mail_body { get; set; }
            public string mail_attachments { get; set; }

        }
        private bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            return redirectionUrl.ToLower().StartsWith("https://");
        }
        public string sendMail(sendEmailModel value)
        {
            String s_mail_to = value.mail_to + "";
            String s_mail_cc = value.mail_cc + "";
            String s_subject = value.mail_subject + "";
            String s_mail_body = value.mail_body + "";
            String s_mail_attachments = value.mail_attachments + "";

            String msg_mail = "";
            String msg_mail_file = "";
            Boolean SendAndSaveCopy = false;

            string mail_font = "";
            string mail_fontsize = "";

            //"MailSMTPServer": "smtp-tsr.thaioil.localnet",
            //"MailFrom": "xxx@thaioilgroup.com",
            //"MailPassword": "xxx",
            //"MailTest": "xxx@thaioilgroup.com;",
            string mail_server = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailSMTPServer"];
            string mail_from = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailFrom"];
            string mail_password = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailPassword"];
            string mail_test = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailTest"];

            if (mail_test != "")
            {
                s_mail_body += "</br></br>ข้อมูล mail to: " + s_mail_to + "</br></br>ข้อมูล mail cc: " + s_mail_cc;

                s_mail_to = mail_test;
                s_mail_cc = mail_test;
            }

            ExchangeService service = new ExchangeService();
            service.Credentials = new WebCredentials(mail_from, mail_password);
            service.TraceEnabled = true;

            // Look up the user's EWS endpoint by using Autodiscover.  
            EmailMessage email = new EmailMessage(service);
            service.AutodiscoverUrl(mail_from, RedirectionUrlValidationCallback);
            email.From = new EmailAddress("Mail Display ใส่ไม่มีผล", mail_from);
            //email.From.Name = "Car Service TSR";
            //email.From.Address = MailFrom;

            var email_to = s_mail_to.Split(';');
            for (int i = 0; i < email_to.Length; i++)
            {
                string _mail = (email_to[i].ToString()).Trim();
                if (_mail != "")
                {
                    // Mail To จะต้องใช้วิธี Loop และห้ามใส่ ; ต่อท้าย
                    email.ToRecipients.Add(_mail);
                }
            }
            var email_cc = s_mail_cc.Split(';');
            for (int i = 0; i < email_cc.Length; i++)
            {
                string _mail = (email_cc[i].ToString()).Trim();
                if (_mail != "")
                {
                    //Mail CC จะต้องใช้วิธี Loop และห้ามใส่ ; ต่อท้าย
                    email.CcRecipients.Add(_mail);
                }
            }

            //Subject
            if (mail_test != "") { s_subject = "(DEV)" + s_subject; }
            email.Subject = s_subject;

            //Body
            //เพิ่ม กำหนด font  
            if (mail_font == "") { mail_font = "Cordia New"; }
            if (mail_fontsize == "") { mail_fontsize = "18"; }
            s_mail_body = "<html><div style='font-size:" + mail_fontsize + "px; font-family:" + mail_font + ";'>" + s_mail_body + "</div></html>";
            email.Body = new MessageBody(BodyType.HTML, s_mail_body);

            try
            {
                msg_mail_file = "";
                //Attachments
                //string filePath = Path.Combine(Server.MapPath("~/temp"), "EMPLOYEE LETTER_TES_Mr._Luck_Saraya_170521012548" + ".docx");
                string filePath = s_mail_attachments;
                if ((s_mail_attachments + "") != "")
                {
                    string[] xsplit_attachments = s_mail_attachments.Split(new char[] { '|', '|' }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < xsplit_attachments.Length; i++)
                    {
                        string templateFile = xsplit_attachments[i];
                        email.Attachments.AddFileAttachment(templateFile);
                    }
                }
            }
            catch (Exception ex)
            {
                msg_mail_file = ex.ToString();
            }

            try
            {
                if (SendAndSaveCopy == true)
                {
                    //จะมีใน send box item
                    email.SendAndSaveCopy();
                }
                else
                {
                    //ไม่เก็บใน send box item
                    email.Send();
                }
                msg_mail = "";
            }
            catch (Exception ex)
            {
                msg_mail = ex.ToString();
            }

            return msg_mail;
        }

    }


}
