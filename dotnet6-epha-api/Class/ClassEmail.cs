using Aspose.Cells;
using dotnet6_epha_api.Class;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using Microsoft.Exchange.WebServices.Data;
using Model;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Buffers.Text;
using System.Data;
using System.Drawing;
using System.Net;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Security.Policy;

namespace Class
{
    //<!-- Mail Server />-->
    //<add key = "Server" value="PRO" />
    //<add key = "MailFrom" value="application@thaioilgroup.com" /> 
    //<add key = "emai_test" value="zkul-uwat@thaioilgroup.com;zmiyukis@thaioilgroup.com" />
    //<add key = "MailSMTPServer" value="smtp-tsr.thaioil.localnet" />



    public class ClassEmail
    {
        //string server_url = "WebServer_ePHA_Index";// @"https://localhost:7096/hazop/Index?";
        string server_url = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["WebServer_ePHA_Index"];


        string sqlstr = "";
        string jsper = "";
        ClassFunctions cls = new ClassFunctions();
        ClassJSON cls_json = new ClassJSON();
        ClassConnectionDb cls_conn = new ClassConnectionDb();

        DataSet dsData;
        DataTable dt, dtcopy, dtcheck;


        private bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            return redirectionUrl.ToLower().StartsWith("https://");
        }
        public class sendEmailModel
        {
            public string mail_from { get; set; }
            public string mail_to { get; set; }
            public string mail_cc { get; set; }
            public string mail_subject { get; set; }
            public string mail_body { get; set; }
            public string mail_attachments { get; set; }

        }
        private static string EncryptDataWithAes(string plainText, string keyBase64, out string vectorBase64)
        {
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.Key = Convert.FromBase64String(keyBase64);
                aesAlgorithm.GenerateIV();
                Console.WriteLine($"Aes Cipher Mode : {aesAlgorithm.Mode}");
                Console.WriteLine($"Aes Padding Mode: {aesAlgorithm.Padding}");
                Console.WriteLine($"Aes Key Size : {aesAlgorithm.KeySize}");

                //set the parameters with out keyword
                vectorBase64 = Convert.ToBase64String(aesAlgorithm.IV);

                // Create encryptor object
                ICryptoTransform encryptor = aesAlgorithm.CreateEncryptor();

                byte[] encryptedData;

                //Encryption will be done in a memory stream through a CryptoStream object
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter sw = new StreamWriter(cs))
                        {
                            sw.Write(plainText);
                        }
                        encryptedData = ms.ToArray();
                    }
                }

                return Convert.ToBase64String(encryptedData);
            }
        }
        private static string DecryptDataWithAes(string cipherText, string keyBase64, string vectorBase64)
        {
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.Key = Convert.FromBase64String(keyBase64);
                aesAlgorithm.IV = Convert.FromBase64String(vectorBase64);

                Console.WriteLine($"Aes Cipher Mode : {aesAlgorithm.Mode}");
                Console.WriteLine($"Aes Padding Mode: {aesAlgorithm.Padding}");
                Console.WriteLine($"Aes Key Size : {aesAlgorithm.KeySize}");
                Console.WriteLine($"Aes Block Size : {aesAlgorithm.BlockSize}");


                // Create decryptor object
                ICryptoTransform decryptor = aesAlgorithm.CreateDecryptor();

                byte[] cipher = Convert.FromBase64String(cipherText);

                //Decryption will be done in a memory stream through a CryptoStream object
                using (MemoryStream ms = new MemoryStream(cipher))
                {
                    using (CryptoStream cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader sr = new StreamReader(cs))
                        {
                            return sr.ReadToEnd();
                        }
                    }
                }
            }
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

            String mail_server = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailSMTPServer"];
            String mail_from = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailFrom"];
            String mail_test = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailTest"];

            string mail_font = "";
            string mail_fontsize = "";

            string mail_user = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailUser"];
            string mail_password = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailPassword"];
            mail_user = "zkuluwat@thaioilgroup.com";
            mail_password = "Initial1;Q4";

            //mail_user = "kuluwat@adb-thailand.com";
            //mail_password = "Initial1;d";


            if (mail_test != "")
            {
                s_mail_body += "</br></br>ข้อมูล mail to: " + s_mail_to + "</br></br>ข้อมูล mail cc: " + s_mail_cc;

                s_mail_to = mail_test;
                s_mail_cc = mail_test + ";kuluwat@adb-thailand.com";
            }

            ExchangeService service = new ExchangeService();
            service.Credentials = new WebCredentials(mail_user, mail_password);
            service.TraceEnabled = true;

            // Look up the user's EWS endpoint by using Autodiscover.  
            EmailMessage email = new EmailMessage(service);
            service.AutodiscoverUrl(mail_user, RedirectionUrlValidationCallback);
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

        string s_subject = "";
        string s_body = "";

        public string get_mail_admin_group()
        {

            string _mail = "";

            ClassLogin cls_login = new ClassLogin();
            sqlstr = cls_login.QueryAdminUser_Role("");

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (i > 0) { _mail += ";"; }
                    _mail += (dt.Rows[i]["user_email"] + "");
                }
            }
            return _mail;
        }
        public string QueryActionOwner(string seq, string responder_user_name, string sub_software)
        {
            cls = new ClassFunctions();
            sqlstr = @" select h.pha_status, h.pha_no, g.pha_request_name as pha_name, empre.user_email as request_email
                             , a.responder_user_name, emp.user_displayname, emp.user_email
                             , count(1) as total
                             , count(case when lower(a.action_status) = 'open' then 1 else null end) 'open'
                             , count(case when lower(a.action_status) = 'closed' then 1 else null end) 'closed' 
                             , g.reference_moc
                             from EPHA_F_HEADER h
                             inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha) 
                             left join EPHA_T_NODE_WORKSHEET a on lower(h.id) = lower(a.id_pha) 
                             left join EPHA_PERSON_DETAILS emp on lower(a.responder_user_name) = lower(emp.user_name)  
                             left join EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name)  
                             where a.responder_user_name is not null and h.id = " + seq;
            if (responder_user_name != "") { sqlstr += " and lower(a.responder_user_name) = lower(" + cls.ChkSqlStr(responder_user_name, 50) + ") "; }
            sqlstr += "  group by h.pha_status, h.pha_no, g.pha_request_name, empre.user_email, a.responder_user_name, emp.user_displayname, emp.user_email, a.action_status, g.reference_moc";
            return sqlstr;
        }
        public string MailToPHAConduct(string seq, string sub_software)
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";
            string url = "";
            string step_text = "PHA Conduct.";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string meeting_date = "";
            DataTable dt = new DataTable();

            if (sub_software == "hazop")
            {
                sqlstr = @" select distinct h.pha_status, h.pha_no as pha_no,g.pha_request_name as pha_name,empre.user_email as request_email
                        , b.no, format(a.meeting_date, 'dd MMM yyyy') +'('+ replace(a.meeting_start_time,'1/1/1970 ','') +'-'+ replace(a.meeting_end_time,'1/1/1970 ','') +')' as meeting_date
                        ,emp.user_displayname, emp.user_email, g.reference_moc
                        from EPHA_F_HEADER h
                        inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha) 
                        inner join EPHA_T_SESSION a on lower(h.id) = lower(a.id_pha) 
                        inner join EPHA_T_MEMBER_TEAM b on lower(a.id) = lower(b.id_session) 
                        inner join EPHA_PERSON_DETAILS emp on lower(b.user_name) = lower(emp.user_name) 
                        inner join EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name) 
                        where h.id =" + seq;
                sqlstr += " order by b.no";
            }
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            #region mail to
            if (dt.Rows.Count > 0)
            {
                doc_no = (dt.Rows[0]["pha_no"] + "");
                doc_name = (dt.Rows[0]["pha_name"] + "");
                reference_moc = (dt.Rows[0]["reference_moc"] + "");
                meeting_date = (dt.Rows[0]["meeting_date"] + "");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (i > 0) { s_mail_to += ";"; }
                    s_mail_to += (dt.Rows[i]["user_email"] + "");
                }
            }
            #endregion mail to

            #region mail cc 
            if (dt.Rows.Count > 0)
            {
                //cc to pha_request_email
                s_mail_cc = (dt.Rows[0]["request_email"] + "");
            }
            #endregion mail cc

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                //insert keyBase64 to db 
                string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=2";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                //string x = DecryptDataWithAes(cipherText, keyBase64, vectorBase64);

                url = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;
            }
            #endregion url 


            s_subject = "ePHA Online System : " + doc_no + (doc_name == "" ? "" : "") + ",Please be invited to meeting to conduct of PHA.";

            s_body = "<html><body><font face='tahoma' size='2'>";
            s_body += "Dear " + to_displayname + ",";

            s_body += "<br/><br/><b>Step</b> : " + step_text;
            s_body += "<br/><b>Reference MOC</b> : " + reference_moc;
            s_body += "<br/><b>Project Name</b> : " + doc_name;

            s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please be invited to meeting to conduct of PHA No." + doc_no;
            s_body += "<br/>To see the detailed infromation,<font color='red'> please click <a href='" + url + "'>here</a></font>";

            s_body += "<br/><br/>Best Regards,";
            s_body += "<br/>ePHA Online System ";
            s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
            s_body += "</font></body></html>";

            sendEmailModel data = new sendEmailModel();
            data.mail_subject = s_subject;
            data.mail_body = s_body;
            data.mail_to = s_mail_to;
            data.mail_cc = s_mail_cc;
            data.mail_from = s_mail_from;

            return sendMail(data);


        }
        public string MailToActionOwner(string seq, string sub_software )
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";

            string url = "";
            string step_text = "PHA Follow up Item";

            string to_displayname = "";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = ""; 

            string meeting_date = "";
            DataTable dt = new DataTable();
             

            if (sub_software == "hazop")
            {  
                sqlstr = QueryActionOwner(seq, "", sub_software);
            }
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                //insert keyBase64 to db 
                string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                //string x = DecryptDataWithAes(cipherText, keyBase64, vectorBase64);

                url = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;
            }
            #endregion url 


            #region mail to
            string msg = "";
            if (dt.Rows.Count > 0)
            {
                string xbefor = "";
                string xafter = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    xbefor = (dt.Rows[i]["user_displayname"] + "");
                    if (xbefor != xafter)
                    {
                        xafter = xbefor;
                    }
                    else { if (i != dt.Rows.Count - 1) { continue; } }

                    //cc to pha_request_email
                    s_mail_cc = (dt.Rows[i]["request_email"] + "");
                    s_mail_to = (dt.Rows[i]["user_email"] + "");

                    doc_no = (dt.Rows[0]["pha_no"] + "");
                    doc_name = (dt.Rows[0]["pha_name"] + "");
                    reference_moc = (dt.Rows[0]["reference_moc"] + "");
                    to_displayname = (dt.Rows[i]["user_displayname"] + "");

                    int iTotal = 0; int iOpen = 0; int iClosed = 0;
                    iTotal = Convert.ToInt32(dt.Rows[i]["total"] + "");
                    iOpen = Convert.ToInt32(dt.Rows[i]["open"] + "");
                    iClosed = Convert.ToInt32(dt.Rows[i]["closed"] + "");

                    s_subject = "ePHA Online System : " + doc_no + (doc_name == "" ? "" : "")
                              + ",Please follow up item and update action.";

                    s_body = "<html><body><font face='tahoma' size='2'>";
                    s_body += "Dear " + to_displayname + ",";

                    s_body += "<br/><br/><b>Step</b> : " + step_text;
                    s_body += "<br/><b>Reference MOC</b> : " + reference_moc;
                    s_body += "<br/><b>Project Name</b> : " + doc_name;
                    s_body += "<br/><br/>Items Status Total: " + iTotal + ", Open: " + iOpen + ", Closed: " + iClosed + " ";

                    s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please review data of PHA No." + doc_no;
                    s_body += "<br/>To see the detailed infromation,<font color='red'> please click <a href='" + url + "'>here</a></font>";

                    s_body += "<br/><br/>Best Regards,";
                    s_body += "<br/>ePHA Online System ";
                    s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
                    s_body += "</font></body></html>";


                    sendEmailModel data = new sendEmailModel();
                    data.mail_subject = s_subject;
                    data.mail_body = s_body;
                    data.mail_to = s_mail_to;
                    data.mail_cc = s_mail_cc;
                    data.mail_from = s_mail_from;

                    msg = sendMail(data);
                    if (msg != "") { }
                }
            }
            #endregion mail to

            return "";


        }
        public string MailToApproverReview(string seq, string sub_software)
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";

            string url = "";
            string url_approver = "";
            string url_reject = "";
            string step_text = "Approver TA2 Review.";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            DataTable dt = new DataTable();

            if (sub_software == "hazop")
            {
                sqlstr = @"  select h.approver_user_name,h.pha_status, h.pha_no, g.pha_request_name, emp.user_displayname, emp.user_email 
                             , h.approve_action_type, h.approve_status, h.approve_comment, g.reference_moc
                             from EPHA_F_HEADER h
                             inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha)  
                             left join EPHA_PERSON_DETAILS emp on lower(h.approver_user_name) = lower(emp.user_name)   
                             where h.approver_user_name is not null and h.id =" + seq;

            }
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            #region mail to
            if (dt.Rows.Count > 0)
            {
                doc_no = (dt.Rows[0]["pha_no"] + "");
                doc_name = (dt.Rows[0]["pha_name"] + "");
                reference_moc = (dt.Rows[0]["reference_moc"] + "");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (i > 0) { s_mail_to += ";"; }
                    s_mail_to += (dt.Rows[i]["user_email"] + "");
                }
            }
            #endregion mail to

            #region mail cc 
            if (dt.Rows.Count > 0)
            {
                //cc to pha_request_email
                s_mail_cc = (dt.Rows[0]["request_email"] + "");
            }
            #endregion mail cc

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                //insert keyBase64 to db 
                string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                url = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                //reject 
                url_reject = url;

                //approve
                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=approve";
                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                url_reject = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;

            }
            #endregion url 


            s_subject = "ePHA Online System : " + doc_no + (doc_name == "" ? "" : "")
                + ",Please review data.";

            s_body = "<html><body><font face='tahoma' size='2'>";
            s_body += "Dear " + to_displayname + ",>";

            s_body += "<br/><br/><b>Step</b> : " + step_text;
            s_body += "<br/><b>Reference MOC</b> : " + reference_moc;
            s_body += "<br/><b>Project Name</b> : " + doc_name;

            s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please review data of PHA No." + doc_no;
            s_body += "<br/>To see the detailed infromation,<font color='red'> please click <a href='" + url + "'>here</a></font>";

            s_body += "<br/><br/> <font color='blue'><a href='" + url_approver + "'>Approve</a></font> or <font color='red'><a href='" + url_reject + "'>Send back with Comment</a></font>";


            s_body += "<br/><br/>Best Regards,";
            s_body += "<br/>ePHA Online System ";
            s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
            s_body += "</font></body></html>";

            sendEmailModel data = new sendEmailModel();
            data.mail_subject = s_subject;
            data.mail_body = s_body;
            data.mail_to = s_mail_to;
            data.mail_cc = s_mail_cc;
            data.mail_from = s_mail_from;

            return sendMail(data);
        }
        public string MailApprovByApprover(string seq, string sub_software)
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";
            string comment = "";
            string approver_displayname = "XXXXX (TOP-XX)";

            string url = "";
            string url_approver = "";
            string url_reject = "";
            string step_text = "Approver TA2 Approve.";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string mail_admin_group = get_mail_admin_group();
            DataTable dt = new DataTable();

            if (sub_software == "hazop")
            {
                sqlstr = @"  select h.approver_user_name,h.pha_status, h.pha_no, g.pha_request_name, emp.user_displayname, emp.user_email 
                             , h.approve_action_type, h.approve_status, h.approve_comment, g.reference_moc
                             from EPHA_F_HEADER h
                             inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha)  
                             left join EPHA_PERSON_DETAILS emp on lower(h.approver_user_name) = lower(emp.user_name)   
                             where h.approver_user_name is not null and h.id =" + seq;

            }
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            #region mail to
            if (dt.Rows.Count > 0)
            {
                doc_no = (dt.Rows[0]["pha_no"] + "");
                doc_name = (dt.Rows[0]["pha_name"] + "");
                reference_moc = (dt.Rows[0]["reference_moc"] + "");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //to pha_request_email, admin
                    if (i > 0) { s_mail_to += ";"; }
                    s_mail_to += (dt.Rows[i]["request_email"] + "");
                }
                s_mail_to += mail_admin_group ;
            }
            #endregion mail to

            #region mail cc 
            if (dt.Rows.Count > 0)
            {
                //cc approver ta2
              s_mail_cc += (dt.Rows[0]["user_email"] + "");
            }
            #endregion mail cc

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                //insert keyBase64 to db 
                string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=3";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                url = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                //reject 
                url_reject = url;

                //approve
                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=approve";
                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                url_reject = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;

            }
            #endregion url 


            s_subject = "ePHA Online System : " + doc_no + (doc_name == "" ? "" : "")
                            + ",Please follow up item and update action.";

            s_body = "<html><body><font face='tahoma' size='2'>";
            s_body += "Dear " + to_displayname + ",";

            s_body += "<br/><br/><b>Step</b> : " + step_text;
            s_body += "<br/><b>Reference MOC</b> : " + reference_moc;
            s_body += "<br/><b>Project Name</b> : " + doc_name;
            
            s_body += "<br/><b>" + approver_displayname + ", has approved the conduct of PHA</b>";
            if (comment != "") { s_body += "<br/><b>Comment: " + comment; }

            s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please review data of PHA No." + doc_no;
            s_body += "<br/>To see the detailed infromation,<font color='red'> please click <a href='" + url + "'>here</a></font>";

            s_body += "<br/><br/>Best Regards,";
            s_body += "<br/>ePHA Online System ";
            s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
            s_body += "</font></body></html>";

            sendEmailModel data = new sendEmailModel();
            data.mail_subject = s_subject;
            data.mail_body = s_body;
            data.mail_to = s_mail_to;
            data.mail_cc = s_mail_cc;
            data.mail_from = s_mail_from;

            return sendMail(data);

        }
        public string MailRejectByApprover(string seq, string sub_software)
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";
            string comment = "";
            string approver_displayname = "XXXXX (TOP-XX)";

            string url = "";
            string url_approver = "";
            string url_reject = "";
            string step_text = "ApproverTA2 Send back with Comment.";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            DataTable dt = new DataTable();

            if (sub_software == "hazop")
            {
                sqlstr = @"  select h.approver_user_name,h.pha_status, h.pha_no, g.pha_request_name, emp.user_displayname, emp.user_email 
                             , h.approve_action_type, h.approve_status, h.approve_comment, g.reference_moc
                             from EPHA_F_HEADER h
                             inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha)  
                             left join EPHA_PERSON_DETAILS emp on lower(h.approver_user_name) = lower(emp.user_name)   
                             where h.approver_user_name is not null and h.id =" + seq;

            }
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            #region mail to
            if (dt.Rows.Count > 0)
            {
                doc_no = (dt.Rows[0]["pha_no"] + "");
                doc_name = (dt.Rows[0]["pha_name"] + "");
                reference_moc = (dt.Rows[0]["reference_moc"] + "");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (i > 0) { s_mail_to += ";"; }
                    s_mail_to += (dt.Rows[i]["user_email"] + "");
                }
            }
            #endregion mail to

            #region mail cc 
            if (dt.Rows.Count > 0)
            {
                //cc to pha_request_email
                s_mail_cc = (dt.Rows[0]["request_email"] + "");
            }
            #endregion mail cc

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                //insert keyBase64 to db 
                string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                url = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                //reject 
                url_reject = url;

                //approve
                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=approve";
                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                url_reject = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;

            }
            #endregion url 


            s_subject = "ePHA Online System : " + doc_no + (doc_name == "" ? "" : "")
                + ",Please be invited to meeting to conduct of PHA.";

            s_body = "<html><body><font face='tahoma' size='2'>";
            s_body += "Dear " + to_displayname + ",";

            s_body += "<br/><br/><b>Step</b> : " + step_text;
            s_body += "<br/><b>Reference MOC</b> : " + reference_moc;
            s_body += "<br/><b>Project Name</b> : " + doc_name;

            s_body += "<br/><b>"+ approver_displayname + ",  Send back with Comment</b>"  ;
            s_body += "<br/><b>Comment: "+ comment;

            s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please review data of PHA No." + doc_no;
            s_body += "<br/>To see the detailed infromation,<font color='red'> please click <a href='" + url + "'>here</a></font>";
             
            s_body += "<br/><br/>Best Regards,";
            s_body += "<br/>ePHA Online System ";
            s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
            s_body += "</font></body></html>";

            sendEmailModel data = new sendEmailModel();
            data.mail_subject = s_subject;
            data.mail_body = s_body;
            data.mail_to = s_mail_to;
            data.mail_cc = s_mail_cc;
            data.mail_from = s_mail_from;

            return sendMail(data);

        }

        public string MailNotificationToAdminReviewByOwner(string seq, string responder_user_name, string sub_software)
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";
            string user_displayname = "";

            string url = "";
            string step_text = "Notification";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string meeting_date = "";
            DataTable dt = new DataTable();

            string mail_admin_group = get_mail_admin_group();

            if (sub_software == "hazop")
            {
                sqlstr = QueryActionOwner(seq, responder_user_name, sub_software);

            }
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                //insert keyBase64 to db 
                string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                //string x = DecryptDataWithAes(cipherText, keyBase64, vectorBase64);

                url = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;
            }
            #endregion url 


            #region mail to
            s_mail_to = mail_admin_group;

            string msg = "";
            if (dt.Rows.Count > 0)
            {
                string xbefor = "";
                string xafter = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    xbefor = (dt.Rows[i]["user_displayname"] + "");
                    if (xbefor != xafter)
                    {
                        xafter = xbefor;
                    }
                    else { if (i != dt.Rows.Count - 1) { continue; } }

                    //cc to  action owner
                    s_mail_cc = (dt.Rows[i]["user_email"] + "");

                    doc_no = (dt.Rows[0]["pha_no"] + "");
                    doc_name = (dt.Rows[0]["pha_name"] + "");
                    reference_moc = (dt.Rows[0]["reference_moc"] + "");
                    user_displayname = (dt.Rows[i]["user_displayname"] + "");

                    int iTotal = 0; int iOpen = 0; int iClosed = 0;
                    iTotal = Convert.ToInt32(dt.Rows[i]["total"] + "");
                    iOpen = Convert.ToInt32(dt.Rows[i]["open"] + "");
                    iClosed = Convert.ToInt32(dt.Rows[i]["closed"] + "");

                    s_subject = "ePHA Online System : " + doc_no + (doc_name == "" ? "" : "")
                              + ",Please follow up item and update action.";

                    s_body = "<html><body><font face='tahoma' size='2'>";
                    s_body += "Dear " + to_displayname + ",";

                    s_body += "<br/><br/><b>Step</b> : " + step_text;
                    s_body += "<br/><b>Reference MOC</b> : " + reference_moc;
                    s_body += "<br/><b>Project Name</b> : " + doc_name;

                    s_body += "<br/><br/>" + user_displayname + " has updated the statuses of all tasks.";
                    s_body += "<br/>Items Status Total: " + iTotal + ", Open: " + iOpen + ", Closed: " + iClosed + " ";

                    s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please review data of PHA No." + doc_no;
                    s_body += "<br/>To see the detailed infromation,<font color='red'> please click <a href='" + url + "'>here</a></font>";

                    s_body += "<br/><br/>Best Regards,";
                    s_body += "<br/>ePHA Online System ";
                    s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
                    s_body += "</font></body></html>";


                    sendEmailModel data = new sendEmailModel();
                    data.mail_subject = s_subject;
                    data.mail_body = s_body;
                    data.mail_to = s_mail_to;
                    data.mail_cc = s_mail_cc;
                    data.mail_from = s_mail_from;

                    msg = sendMail(data);
                    if (msg != "") { }
                }
            }
            #endregion mail to

            return msg;


        }
        public string MailNotificationToAdminOwnerUpdateAction(string seq, string responder_user_name, string sub_software)
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";
            string user_displayname = "";

            string url = "";
            string step_text = "Notification Closed Item";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string meeting_date = "";
            DataTable dt = new DataTable();

            string mail_admin_group = get_mail_admin_group();

            if (sub_software == "hazop")
            {
                sqlstr = QueryActionOwner(seq, responder_user_name, sub_software);

            }
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                //insert keyBase64 to db 
                string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                //string x = DecryptDataWithAes(cipherText, keyBase64, vectorBase64);

                url = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;
            }
            #endregion url 


            #region mail to
            s_mail_to = mail_admin_group;

            string msg = "";
            if (dt.Rows.Count > 0)
            {
                string xbefor = "";
                string xafter = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    xbefor = (dt.Rows[i]["user_displayname"] + "");
                    if (xbefor != xafter)
                    {
                        xafter = xbefor;
                    }
                    else { if (i != dt.Rows.Count - 1) { continue; } }


                    doc_no = (dt.Rows[0]["pha_no"] + "");
                    doc_name = (dt.Rows[0]["pha_name"] + "");
                    reference_moc = (dt.Rows[0]["reference_moc"] + "");
                    user_displayname = (dt.Rows[i]["user_displayname"] + "");

                    int iTotal = 0; int iOpen = 0; int iClosed = 0;
                    iTotal = Convert.ToInt32(dt.Rows[i]["total"] + "");
                    iOpen = Convert.ToInt32(dt.Rows[i]["open"] + "");
                    iClosed = Convert.ToInt32(dt.Rows[i]["closed"] + "");

                    s_subject = "ePHA Online System : " + doc_no + (doc_name == "" ? "" : "")
                              + ",The Responder has updated the action status.";

                    s_body = "<html><body><font face='tahoma' size='2'>";
                    s_body += "Dear " + to_displayname + ",";

                    s_body += "<br/><br/><b>Step</b> : " + step_text;
                    s_body += "<br/><b>Reference MOC</b> : " + reference_moc;
                    s_body += "<br/><b>Project Name</b> : " + doc_name;

                    s_body += "<br/><br/>" + user_displayname + " has updated the action status.";
                    s_body += "<br/>Items Status Total: " + iTotal + ", Open: " + iOpen + ", Closed: " + iClosed + " ";

                    s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please review data of PHA No." + doc_no;
                    s_body += "<br/>To see the detailed infromation,<font color='red'> please click <a href='" + url + "'>here</a></font>";

                    s_body += "<br/><br/>Best Regards,";
                    s_body += "<br/>ePHA Online System ";
                    s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
                    s_body += "</font></body></html>";


                    sendEmailModel data = new sendEmailModel();
                    data.mail_subject = s_subject;
                    data.mail_body = s_body;
                    data.mail_to = s_mail_to;
                    data.mail_cc = s_mail_cc;
                    data.mail_from = s_mail_from;

                    msg = sendMail(data);
                    if (msg != "") { }
                }
            }
            #endregion mail to

            return "";


        }

        public string MailToAdminReviewAll(string seq, string sub_software)
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";
            string user_displayname = "";

            string url = "";
            string step_text = "Notification Follow Up";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            string meeting_date = "";
            DataTable dt = new DataTable();

            string mail_admin_group = get_mail_admin_group();

            if (sub_software == "hazop")
            {
                sqlstr = @"  select h.approver_user_name,h.pha_status, h.pha_no, g.pha_request_name as pha_name, emp.user_displayname, emp.user_email 
                             , h.approve_action_type, h.approve_status, h.approve_comment, g.reference_moc
                             from EPHA_F_HEADER h
                             inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha)  
                             left join EPHA_PERSON_DETAILS emp on lower(h.approver_user_name) = lower(emp.user_name)   
                             where h.id =" + seq; 
            }
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];


            #region mail to
            s_mail_to = mail_admin_group;

            string msg = "";
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                { 
                    //cc to  action owner
                    s_mail_cc = (dt.Rows[i]["user_email"] + "");

                    doc_no = (dt.Rows[0]["pha_no"] + "");
                    doc_name = (dt.Rows[0]["pha_name"] + "");
                    reference_moc = (dt.Rows[0]["reference_moc"] + "");
                    user_displayname = "Responder";

                    #region url 

                    using (Aes aesAlgorithm = Aes.Create())
                    {
                        aesAlgorithm.KeySize = 256;
                        aesAlgorithm.GenerateKey();
                        string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                        //insert keyBase64 to db 
                        string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4";
                        string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                        //string x = DecryptDataWithAes(cipherText, keyBase64, vectorBase64);

                        url = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                    }
                    #endregion url 
                     
                    s_subject = "ePHA Online System : " + doc_no + (doc_name == "" ? "" : "")
                              + ",Please follow up item and update action.";

                    s_body = "<html><body><font face='tahoma' size='2'>";
                    s_body += "Dear " + to_displayname + ",";

                    s_body += "<br/><br/><b>Step</b> : " + step_text;
                    s_body += "<br/><b>Reference MOC</b> : " + reference_moc;
                    s_body += "<br/><b>Project Name</b> : " + doc_name;

                    s_body += "<br/><br/>" + user_displayname + " has updated the statuses of all tasks."; 

                    s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please review data of PHA No." + doc_no;
                    s_body += "<br/>To see the detailed infromation,<font color='red'> please click <a href='" + url + "'>here</a></font>";

                    s_body += "<br/><br/>Best Regards,";
                    s_body += "<br/>ePHA Online System ";
                    s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
                    s_body += "</font></body></html>";


                    sendEmailModel data = new sendEmailModel();
                    data.mail_subject = s_subject;
                    data.mail_body = s_body;
                    data.mail_to = s_mail_to;
                    data.mail_cc = s_mail_cc;
                    data.mail_from = s_mail_from;

                    msg = sendMail(data);
                    if (msg != "") { }
                }
            }
            #endregion mail to

            return msg;


        }
        public string MailToAllUserClosed(string seq, string sub_software)
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";

            string url = "";
            string url_approver = "";
            string url_reject = "";
            string step_text = "Admin Closed PHA.";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            DataTable dt = new DataTable();
            DataTable dtAction = new DataTable();

            if (sub_software == "hazop")
            {

                sqlstr = @"  select h.approver_user_name,h.pha_status, h.pha_no, g.pha_request_name as pha_name, emp.user_displayname, emp.user_email 
                             , emp2.user_email as request_email
                             , h.approve_action_type, h.approve_status, h.approve_comment, g.reference_moc
                             from EPHA_F_HEADER h
                             inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha)  
                             left join EPHA_PERSON_DETAILS emp on lower(h.approver_user_name) = lower(emp.user_name)   
                             left join EPHA_PERSON_DETAILS emp2 on lower(h.approver_user_name) = lower(emp2.user_name)   
                             where h.id =" + seq;

            }
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            #region mail to
            if (dt.Rows.Count > 0)
            {
                doc_no = (dt.Rows[0]["pha_no"] + "");
                doc_name = (dt.Rows[0]["pha_name"] + "");
                reference_moc = (dt.Rows[0]["reference_moc"] + "");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (i > 0) { s_mail_to += ";"; }
                    s_mail_to += (dt.Rows[i]["user_email"] + "");
                }
            }
            #endregion mail to

            #region mail cc 
            if (dt.Rows.Count > 0)
            {
                //cc to pha_request_email
                s_mail_cc = (dt.Rows[0]["request_email"] + "");
            }
            #endregion mail cc

            #region url  
            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                //insert keyBase64 to db 
                string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=5";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                url = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                //reject 
                url_reject = url;

                //approve
                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=approve";
                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                url_reject = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;

            }
            #endregion url 


            s_subject = "ePHA Online System : " + doc_no + (doc_name == "" ? "" : "")
                + ",Please review data.";

            s_body = "<html><body><font face='tahoma' size='2'>";
            s_body += "Dear " + to_displayname + ",";

            s_body += "<br/><br/><b>Step</b> : " + step_text;
            s_body += "<br/><b>Reference MOC</b> : " + reference_moc;
            s_body += "<br/><b>Project Name</b> : " + doc_name;

            s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please review data of PHA No." + doc_no;
            s_body += "<br/>To see the detailed infromation,<font color='red'> please click <a href='" + url + "'>here</a></font>";
             
            s_body += "<br/><br/>Best Regards,";
            s_body += "<br/>ePHA Online System ";
            s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
            s_body += "</font></body></html>";

            sendEmailModel data = new sendEmailModel();
            data.mail_subject = s_subject;
            data.mail_body = s_body;
            data.mail_to = s_mail_to;
            data.mail_cc = s_mail_cc;
            data.mail_from = s_mail_from;

            return sendMail(data);


        }
     

    }
}
