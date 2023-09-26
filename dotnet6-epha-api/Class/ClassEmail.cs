﻿using Aspose.Cells;
using dotnet6_epha_api.Class;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using Microsoft.Exchange.WebServices.Data;
using Model;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Org.BouncyCastle.Ocsp;
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

            //"MailSMTPServer": "smtp-tsr.thaioil.localnet",
            //"MailFrom": "zkuluwat@thaioilgroup.com",
            //"MailTest": "zkuluwat@thaioilgroup.com;",
            String mail_server = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailSMTPServer"];
            String mail_from = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailFrom"];
            String mail_test = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailTest"];
             
            #region mail test
            sqlstr = @"select email, email as user_email from EPHA_M_CONFIGMAIL where active_type = 1 ";
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count > 0)
            {
                mail_test = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (i > 0) { mail_test += ";"; }
                    mail_test += (dt.Rows[i]["user_email"] + "");
                }
            }
            #endregion mail test

            string mail_font = "";
            string mail_fontsize = "";

            string mail_user = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailUser"];
            string mail_password = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("MailConfig")["MailPassword"];
            mail_user = "zkuluwat@thaioilgroup.com";
            mail_password = "Initial1;Q4";
             
            if (mail_test != "")
            {
                s_mail_body += "</br></br>ข้อมูล mail to: " + s_mail_to + "</br></br>ข้อมูล mail cc: " + s_mail_cc;

                s_mail_to = mail_test;
                s_mail_cc = mail_test;
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

        #region mail workflow
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
            sqlstr = @" select h.pha_status,h.pha_sub_software, h.pha_no, g.pha_request_name as pha_name, empre.user_email as request_email
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
            sqlstr += "  group by h.pha_status,h.pha_sub_software, h.pha_no, g.pha_request_name, empre.user_email, a.responder_user_name, emp.user_displayname, emp.user_email, a.action_status, g.reference_moc";
            return sqlstr;
        }
        public string QueryFollowActionOwner(string seq, Boolean orderby)
        {
            cls = new ClassFunctions();
            sqlstr = @"  select h.id, h.pha_status, h.pha_sub_software, h.pha_no, g.pha_request_name as pha_name, g.reference_moc, empre.user_email as request_email
                             , emp.user_displayname, emp.user_email
							 , nw.consequences as recommendations_no, n.no as node_no, nw.causes_no, nw.recommendations, nw.action_status
							 , format(nw.estimated_start_date,'dd MMM yyyy') as estimated_start_date
							 , format(nw.estimated_end_date,'dd MMM yyyy') as estimated_end_date 
                             , case when nw.action_status = 'closed' then 0 else isnull(datediff(day,case when nw.estimated_end_date > getdate() then getdate() else nw.estimated_end_date end,getdate()),0) end over_due
                             , case when isnull(datediff(day,getdate(),nw.estimated_end_date),0) < 0 then 0 else isnull(datediff(day,getdate(),nw.estimated_end_date),0) end remaining
                             , nw.responder_user_name 
                             from EPHA_F_HEADER h
                             inner join EPHA_T_GENERAL g on lower(h.id) = lower(g.id_pha) 
                             left join EPHA_T_NODE n on lower(h.id) = lower(n.id_pha)  
                             left join EPHA_T_NODE_WORKSHEET nw on lower(h.id) = lower(nw.id_pha) and nw.id_node = n.id
                             left join EPHA_PERSON_DETAILS emp on lower(nw.responder_user_name) = lower(emp.user_name)  
                             left join EPHA_PERSON_DETAILS empre on lower(h.pha_request_by) = lower(empre.user_name)  
                             where nw.responder_user_name is not null and h.id = " + seq;
            if (orderby == true)
            {
                sqlstr += " order by h.pha_sub_software, h.pha_no, convert(int, n.no), convert(int, nw.no)";
            }
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
        public string MailToActionOwner(string seq, string sub_software)
        {

            #region call function  export excel 
            string file_ResponseSheet = "";
            //{"sub_software":"hazop","user_name":"' + user_name + '","seq":"' + seq + '","export_type":"' + data_type + '"}
            ReportModel param = new ReportModel();
            param.seq = seq;
            param.export_type = "pdf";
            param.user_name = "";

            ClassHazopSet classHazopSet = new ClassHazopSet();
            string jsper = classHazopSet.export_hazop_recommendation(param);

            DataTable dtReport = new DataTable();
            cls_json = new ClassJSON();
            dtReport = cls_json.ConvertJSONresult(jsper);
            if (dtReport.Rows.Count > 0)
            {
                file_ResponseSheet = (Path.Combine(Directory.GetCurrentDirectory(), "") + @"/wwwroot" + (dtReport.Rows[0]["ATTACHED_FILE_PATH"] + "").Replace("~", "").Replace(".xlsx", "." + param.export_type));
            }
            #endregion call function  export excel 


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

                    //file excel Response Sheet 

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

                    if (file_ResponseSheet != "")
                    {
                        if (Directory.Exists(file_ResponseSheet))
                        {
                            data.mail_attachments = file_ResponseSheet;
                        }
                    }

                    msg = sendMail(data);
                    if (msg != "") { }
                }
            }
            #endregion mail to

            return "";


        }
        public string MailToApproverReview(string seq, string sub_software)
        {
            #region call function  export excel 
            string file_ResponseSheet = "";
            //{"sub_software":"hazop","user_name":"' + user_name + '","seq":"' + seq + '","export_type":"' + data_type + '"}
            ReportModel param = new ReportModel();
            param.seq = seq;
            param.export_type = "pdf";
            param.user_name = "";

            ClassHazopSet classHazopSet = new ClassHazopSet();
            string jsper = classHazopSet.export_hazop_report(param);

            DataTable dtReport = new DataTable();
            cls_json = new ClassJSON();
            dtReport = cls_json.ConvertJSONresult(jsper);
            if (dtReport.Rows.Count > 0)
            {
                file_ResponseSheet = (Path.Combine(Directory.GetCurrentDirectory(), "") + @"/wwwroot" + (dtReport.Rows[0]["ATTACHED_FILE_PATH"] + "").Replace("~", "").Replace(".xlsx", "." + param.export_type));
            }
            #endregion call function  export excel 


            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";

            string url = "";
            string url_approver = "";
            string url_reject_no_comment = "";
            string url_reject_comment = "";
            string step_text = "Approver TA2 Review.";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";


            DataTable dtOwner = new DataTable();
            cls_conn = new ClassConnectionDb();
            dtOwner = new DataTable();
            dtOwner = cls_conn.ExecuteAdapterSQL(QueryActionOwner(seq, "", sub_software)).Tables[0];

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
                url_reject_comment = url;

                //reject no comment
                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=reject";
                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                url_approver = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                //approve
                plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=4" + "&approver_type=approve";
                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);
                url_approver = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;

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

            s_body += @"<br/><br/>
                        <table style ='border-collapse: collapse;font-family: Tahoma, Geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>
	                        <thead>
		                        <tr>
			                        <td style ='padding: 15px;' rowspan='2'>SUB-SOFTWARE</td>
			                        <td style ='padding: 15px;' rowspan='2'>PHA NO.</td>
			                        <td style ='padding: 15px;' rowspan='2'>RESPONDER</td>
			                        <td style ='padding: 15px; text-align: center;' colspan='3'>ITEMS STATUS</td> 
		                        </tr>
                                <tr>
                                    <td style ='padding: 15px;'>TOTAL</td>
                                    <td style ='padding: 15px;'>OPEN</td>
                                    <td style ='padding: 15px;'>CLOSE</td>		
                                </tr>
	                        </thead> ";

            s_body += "<tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'>";
            for (int o = 0; o < dtOwner.Rows.Count; o++)
            {
                s_body += @"<tr>";
                s_body += "<td style ='padding: 15px;'>" + sub_software.ToUpper() + "</td>";
                s_body += "<td style ='padding: 15px;'>" + dtOwner.Rows[o]["pha_no"] + "</td>";
                s_body += "<td style ='padding: 15px;'>" + dtOwner.Rows[o]["user_displayname"] + "</td>";
                s_body += "<td style ='padding: 15px;'>" + dtOwner.Rows[o]["total"] + "</td>";
                s_body += "<td style ='padding: 15px; color: red'>" + dtOwner.Rows[o]["open"] + "</td>";
                s_body += "<td style ='padding: 15px;'>" + dtOwner.Rows[o]["closed"] + "</td>";
                s_body += "</tr>";
            }
            s_body += " </tbody>";
            s_body += "</table>";

            s_body += "<br/><br/>Reply :";
            s_body += "<a style='border: none;background-color: #25b003; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; ' href='" + url_approver + "'>Approve</a>";
            s_body += "<a style='border: none;background-color: #d90476; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;'  '" + url_reject_no_comment + "'>Send back No Comment</a>";
            s_body += "<a style='border: none;background-color: #f64a8a; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;'  '" + url_reject_comment + "'>Send back with Comment</a>";


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

            if (file_ResponseSheet != "")
            {
                if (Directory.Exists(file_ResponseSheet))
                {
                    data.mail_attachments = file_ResponseSheet;
                }
            }


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
                s_mail_to += mail_admin_group;
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

            s_body += "<br/><b>" + approver_displayname + ",  Send back with Comment</b>";
            s_body += "<br/><b>Comment: " + comment;

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
        public string MailToAdminCaseStudy(string seq, string sub_software)
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";

            string url = "";
            string url_approver = "";
            string url_reject = "";
            string step_text = "Original Closed PHA.";

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

        public string MailClosedAll(string seq, string sub_software)
        {

            #region call function  export excel 
            string file_ResponseSheet = "";
            //{"sub_software":"hazop","user_name":"' + user_name + '","seq":"' + seq + '","export_type":"' + data_type + '"}
            ReportModel param = new ReportModel();
            param.seq = seq;
            param.export_type = "pdf";
            param.user_name = "";

            ClassHazopSet classHazopSet = new ClassHazopSet();
            string jsper = classHazopSet.export_hazop_report(param);

            DataTable dtReport = new DataTable();
            cls_json = new ClassJSON();
            dtReport = cls_json.ConvertJSONresult(jsper);
            if (dtReport.Rows.Count > 0)
            {
                file_ResponseSheet = (Path.Combine(Directory.GetCurrentDirectory(), "") + @"/wwwroot" + (dtReport.Rows[0]["ATTACHED_FILE_PATH"] + "").Replace("~", "").Replace(".xlsx", "." + param.export_type));
            }
            #endregion call function  export excel 


            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";
            string user_displayname = "";

            string url = "";
            string step_text = "Notification Closed";

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
                        string plainText = "seq=" + seq + "&pha_no=" + doc_no + "&step=5";
                        string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);
                        //string x = DecryptDataWithAes(cipherText, keyBase64, vectorBase64);

                        url = server_url + cipherText + "&" + keyBase64 + "&" + vectorBase64;
                    }
                    #endregion url 

                    s_subject = "ePHA Online System : " + doc_no + (doc_name == "" ? "" : "")
                              + ",Updated the statuses of all tasks.";

                    s_body = "<html><body><font face='tahoma' size='2'>";
                    s_body += "Dear " + to_displayname + ",";

                    s_body += "<br/><br/><b>Step</b> : " + step_text;
                    s_body += "<br/><b>Reference MOC</b> : " + reference_moc;
                    s_body += "<br/><b>Project Name</b> : " + doc_name;

                    s_body += "<br/><br/>Updated the statuses of all tasks.";

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

                    if (file_ResponseSheet != "")
                    {
                        if (Directory.Exists(file_ResponseSheet))
                        {
                            data.mail_attachments = file_ResponseSheet;
                        }
                    }

                    msg = sendMail(data);
                    if (msg != "") { }
                }
            }
            #endregion mail to

            return msg;


        }


        #endregion mail workflow

        #region mail noti
        public string MailNotificationFollowUpItemToOwner(string seq)
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

            sqlstr = QueryFollowActionOwner(seq, false);
            sqlstr = "select distinct pha_sub_software, pha_no, responder_user_name, user_displayname, user_email from (" + sqlstr + ")t order by t.pha_sub_software, t.pha_no ";
            cls_conn = new ClassConnectionDb();
            DataTable dtOwner = new DataTable();
            dtOwner = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];


            sqlstr = QueryFollowActionOwner(seq, false);
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            #region mail to
            s_mail_cc = mail_admin_group;

            string msg = "";
            if (dt.Rows.Count > 0)
            {

                for (int iOwner = 0; iOwner < dtOwner.Rows.Count; iOwner++)
                {
                    to_displayname = (dtOwner.Rows[iOwner]["user_displayname"] + "");
                    s_mail_to = (dtOwner.Rows[iOwner]["user_email"] + "");
                    string responder_user_name = (dtOwner.Rows[iOwner]["responder_user_name"] + "");
                    int iNo = 1;

                    s_subject = "ePHA Online System : Notification follow up item and update action.";

                    s_body = "<html><body><font face='tahoma' size='2'>";
                    s_body += "Dear " + to_displayname + ",";

                    s_body += "<br/><br/>Your have the following document(s) for action.";

                    s_body += @"<br/><br/><table style ='border-collapse: collapse;font-family: tahoma, geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>   <thead>    
                                <tr>
                                    <td style ='padding: 15px;' rowspan='2'>No</td>
                                    <td style ='padding: 15px;' rowspan='2'>Sub-Software</td>
                                    <td style ='padding: 15px;' rowspan='2'>PHA no.</td>
                                    <td style ='padding: 15px;' rowspan='2'>Project name/e-moc request number</td>
                                    <td style ='padding: 15px;' rowspan='2'>Responder</td>
                                    <td style ='padding: 15px;' rowspan='2'>Recommendations no </td>
                                    <td style ='padding: 15px;' rowspan='2'>Node no.</td>
                                    <td style ='padding: 15px;' rowspan='2'>Casuse no.</td>
                                    <td style ='padding: 15px;' rowspan='2'>Recommendations</td>
                                    <td style ='padding: 15px;' rowspan='2'>Actions status</td>
                                    <td style ='padding: 15px; text-align: center;' colspan='2'>Estimated date</td>
                                    <td style ='padding: 15px;' rowspan='2'>Over due (day)</td>
                                    <td style ='padding: 15px;' rowspan='2'>Remaining (day)</td>
                                <tr>
                                    <td style ='padding: 15px;'>start date</td>
                                    <td style ='padding: 15px;'>end date</td>
                                </tr>
                                </thead>
                                <tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'> ";

                    for (int a = 0; a < dt.Rows.Count; a++)
                    {
                        if (!(responder_user_name == (dt.Rows[a]["responder_user_name"] + ""))) { continue; }

                        doc_no = (dt.Rows[a]["pha_no"] + "");


                        string background_color = "white";
                        int iRemaining = 0;
                        Boolean action_status_close = (dt.Rows[a]["remaining"] + "").ToLower() == "closed";
                        try
                        {
                            iRemaining = Convert.ToInt32(dt.Rows[a]["remaining"] + "");
                            if (iRemaining > 3)
                            {
                                background_color = "green";
                            }
                            else if ((iRemaining > 0 && iRemaining < 3) && action_status_close == false)
                            {
                                background_color = "yellow";
                            }
                            else if (iRemaining <= 0 && action_status_close == false)
                            { background_color = "red"; }
                        }
                        catch { }

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

                        s_body += "<tr>";
                        s_body += "<td style ='padding: 15px;'>" + (iNo) + "</td>";//1</td>";
                        s_body += "<td style ='padding: 15px;'>" + dt.Rows[a]["pha_sub_software"] + "</td>";//hazop</td>";
                        s_body += "<td style ='padding: 15px;'><a href='" + url + "'>" + dt.Rows[a]["pha_no"] + "</a></td>";//hazop-2023-0000023</td>";
                        s_body += "<td style ='padding: 15px;'><a href='" + url + "'>" + dt.Rows[a]["pha_name"] + "</a></td>";//xxmoc0003</td>";
                        s_body += "<td style ='padding: 15px;'>" + dt.Rows[a]["user_displayname"] + "</td>";//zkul</td>";
                        s_body += "<td style ='padding: 15px;'>" + dt.Rows[a]["recommendations_no"] + "</td>";//recommendations no</td>";
                        s_body += "<td style ='padding: 15px;'>" + dt.Rows[a]["node_no"] + "</td>";//node</td>";
                        s_body += "<td style ='padding: 15px;'>" + dt.Rows[a]["causes_no"] + "</td>";//casuse no</td> ";
                        s_body += "<td style ='padding: 15px;'>" + dt.Rows[a]["recommendations"] + "</td>";//recommendations</td>";
                        s_body += "<td style ='padding: 15px;'>" + dt.Rows[a]["action_status"] + "</td>";//open</td>";
                        s_body += "<td style ='padding: 15px;'>" + dt.Rows[a]["estimated_start_date"] + "</td>";//19-01-2023</td>";
                        s_body += "<td style ='padding: 15px;'>" + dt.Rows[a]["estimated_end_date"] + "</td>";//19-09-2023</td>";
                        s_body += "<td style ='padding: 15px;'>" + dt.Rows[a]["over_due"] + "</td>";// 6</td>";
                        s_body += "<td style ='padding: 15px; background-color:" + background_color + "; '>" + dt.Rows[a]["remaining"] + "</td>";// 6</td>";
                        s_body += "</tr>";
                        iNo += 1;
                    }

                    s_body += "</tbody>";
                    s_body += "</table>";

                    s_body += "<br/><br/>The message of color assignment is as follow:";
                    s_body += "<br/><label style='width: 120px;padding:4px;background-color:green; color:red'>Green Color</label> : &gt; 3 days; this document has more than 3 days to complete your task";
                    s_body += "<br/><label style='width: 120px;padding:4px;background-color:yellow;'>Yellow Color</label> : &lt; 3 days; this document has less than 3 days to complete your task";
                    s_body += "<br/><label style='width: 130px;padding:4px;background-color:Red; color : white'>Red Color</label> : &lt;= 0 days; this document <label style='color:red'>is overdue, please urgent action</label>";

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

        #endregion mail noti

        #region login
        public string MailToAdminRegisterAccount(string _user_displayname, string _user_email, string _user_password, string _user_password_confirm)
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";
            string user_displayname = _user_displayname;

            string urlAccept = ""; string urlNotAccept = "";
            string step_text = "Register Account";

            string to_displayname = "All";
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            DataTable dt = new DataTable();

            string mail_admin_group = get_mail_admin_group();

            string msg = "";

            #region mail to
            s_mail_to = mail_admin_group;

            #region url 

            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                //insert keyBase64 to db 
                string plainText = "user_email=" + _user_email + "&accept_status=1";
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                urlAccept = server_url.Replace("index", "RegisterAccount") + cipherText + "&" + keyBase64 + "&" + vectorBase64;

                //insert keyBase64 to db 
                plainText = "user_email=" + _user_email + "&accept_status=0";
                cipherText = EncryptDataWithAes(plainText, keyBase64, out vectorBase64);

                urlNotAccept = server_url.Replace("index", "RegisterAccount") + cipherText + "&" + keyBase64 + "&" + vectorBase64;
            }
            #endregion url 

            s_subject = "ePHA Online System : Staff or Contractor register account.";

            s_body = "<html><body><font face='tahoma' size='2'>";
            s_body += "Dear " + to_displayname + ",";

            s_body += "<br/><br/>" + user_displayname + " register account.";
            s_body += "<br/>Email address: " + _user_email + " ";
            s_body += "<br/>Password: " + _user_password + " ";
            s_body += "<br/>Confirm Password: " + _user_password_confirm + " ";

            s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please check your registration to use the system.";
            s_body += "<br/><font color='blue'><a href='" + urlAccept + "'>Accept</a></font>";
            s_body += ",<font color='red'><a font color='red' href='" + urlNotAccept + "'>Not Accept</a></font>";

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


            #endregion mail to

            return msg;


        }
        public string MailToUserRegisterAccount(string _user_displayname, string _user_email, string _user_password, string _user_password_confirm, string _accept_status)
        {
            string doc_no = "";
            string doc_name = "";
            string reference_moc = "";
            string user_displayname = _user_displayname;

            string url = "";
            string step_text = "Register Account";

            string to_displayname = _user_displayname;
            string s_mail_to = "";
            string s_mail_cc = "";
            string s_mail_from = "";

            DataTable dt = new DataTable();

            string mail_admin_group = get_mail_admin_group();

            string msg = "";

            #region mail to
            s_mail_to = mail_admin_group;

            #region url 

            using (Aes aesAlgorithm = Aes.Create())
            {
                aesAlgorithm.KeySize = 256;
                aesAlgorithm.GenerateKey();
                string keyBase64 = Convert.ToBase64String(aesAlgorithm.Key);

                //insert keyBase64 to db 
                string plainText = "user_email=" + _user_email;
                string cipherText = EncryptDataWithAes(plainText, keyBase64, out string vectorBase64);

                url = server_url.Replace("hazop", "login") + cipherText + "&" + keyBase64 + "&" + vectorBase64;

            }
            #endregion url 

            s_subject = "ePHA Online System : Staff or Contractor register account.";

            s_body = "<html><body><font face='tahoma' size='2'>";
            s_body += "Dear " + to_displayname + ",";

            s_body += "<br/><br/>Register account.";
            s_body += "<br/><br/>Name: " + user_displayname + " ";
            s_body += "<br/>Email address: " + _user_email + " ";
            s_body += "<br/>Password: " + _user_password + " ";
            s_body += "<br/>Confirm Password: " + _user_password_confirm + " ";

            s_body += "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Admin " + (_accept_status == "1" ? "accept" : "not accept") + " registration account.";
            s_body += "<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;System Administrator has " + (_accept_status == "1" ? "confirmed" : "not confirmed") + " your system registration.";
            if (_accept_status == "1")
            {
                s_body += "<br/><font color='red'>You can now click <a href='" + url + "'>the link</a> to access the system.,</font>";
            }

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


            #endregion mail to

            return msg;


        }
        #endregion login 
        public string MailTest()
        {

            DataTable dt = new DataTable();

            string msg = "";

            #region mail to
            string s_mail_to = "zkuluwat@thaioilgroup.com";

            s_subject = "ePHA Online System : Staff or Contractor register account.";

            s_body = "<html><body><font face='tahoma' size='2'>";
            s_body += "Dear xxx,";

            s_body += "<br/><br/>Register account.";
            s_body += "<br/><br/><button style='border: none;background-color: inherit; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block;' type='button'>Click Me!</button>";

            s_body += @"<br/><br/>
                        <table style ='border-collapse: collapse;font-family: Tahoma, Geneva, sans-serif;background-color: #215289;color: #ffffff;font-weight: bold;font-size: 13px;border: 1px solid #54585d;'>
	                        <thead>
		                        <tr>
			                        <td style ='padding: 15px;' rowspan='2'>SUB-SOFTWARE</td>
			                        <td style ='padding: 15px;' rowspan='2'>PHA NO.</td>
			                        <td style ='padding: 15px;' rowspan='2'>RESPONDER</td>
			                        <td style ='padding: 15px; text-align: center;' colspan='3'>ITEMS STATUS</td> 
		                        </tr>
                                <tr>
                                    <td style ='padding: 15px;'>TOTAL</td>
                                    <td style ='padding: 15px;'>OPEN</td>
                                    <td style ='padding: 15px;'>CLOSE</td>		
                                </tr>
	                        </thead>
	                        <tbody style='color: #636363;background-color: #ffffff;border: 1px solid #dddfe1;'>
		                        <tr>
			                        <td style ='padding: 15px;'>HAZOP</td>
			                        <td style ='padding: 15px;'>HAZOP-2023-0000001</td>
			                        <td style ='padding: 15px;'>zKuluwat S.</td>
			                        <td style ='padding: 15px;'>30</td>
			                        <td style ='padding: 15px; color: red'>20</td>
			                        <td style ='padding: 15px;'>10</td>
		                        </tr> 
	                        </tbody>
                        </table>";

            s_body += "<a  style='border: none;background-color: #25b003; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; '  href='https://localhost:7052/hazop/Index'>Approve</a>";
            s_body += "<a  style='border: none;background-color: #d90476; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;'  href='https://localhost:7052/hazop/Index'>Send back No Comment</a>";
            s_body += "<a  style='border: none;background-color: #f64a8a; padding: 14px 28px;font-size: 14px;cursor: pointer;display: inline-block; margin-left: 30px;'  href='https://localhost:7052/hazop/Index'>Send back with Comment</a>";


            s_body += "<br/><br/>Best Regards,";
            s_body += "<br/>ePHA Online System ";
            s_body += "<br/><br/><br/>Note that this message was automatically sent by ePHA Online System.";
            s_body += "</font></body></html>";


            sendEmailModel data = new sendEmailModel();
            data.mail_subject = s_subject;
            data.mail_body = s_body;
            data.mail_to = s_mail_to;

            msg = sendMail(data);
            if (msg != "") { }


            #endregion mail to

            return msg;


        }
    }
}
