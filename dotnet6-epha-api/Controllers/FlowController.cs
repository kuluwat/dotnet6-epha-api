using Class;
using Microsoft.AspNetCore.Mvc;
using Model;
using Newtonsoft.Json;
using System.Data;
using static Class.ClassEmail;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FlowController : ControllerBase
    {
        [HttpPost("ClearDataTableTransactions", Name = "ClearDataTableTransactions")]
        public string ClearDataTableTransactions()
        {
            string ret = "";
            ClassConnectionDb cls_conn = new ClassConnectionDb();
            System.Data.DataTable dt = new System.Data.DataTable();

            string sqlstr = "SELECT name FROM SYSOBJECTS where lower(name) like 'epha_t%' ";
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                sqlstr = "delete from " + dt.Rows[i]["name"];

                cls_conn = new ClassConnectionDb();
                cls_conn.OpenConnection();
                ret = cls_conn.ExecuteNonQuery(sqlstr);
                cls_conn.CloseConnection();
                if (ret.ToLower() != "true") { break; }
            }

            return (ret.ToLower() != "true" ? ret + "-sqlstr:" + sqlstr : ret);
        }


        [HttpPost("ConnectionSting", Name = "ConnectionSting")]
        public string ConnectionSting(string param)
        {
            String ConnStrSQL = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ConnectionConfig")["ConnString"];

            return ConnStrSQL;

        }

        [HttpPost("ExQueryString", Name = "ExQueryString")]
        public string ExQueryString(string param)
        {
            ClassConnectionDb cls_conn = new ClassConnectionDb();
            System.Data.DataTable dt = new System.Data.DataTable();
            dt = cls_conn.ExecuteAdapterSQL(param).Tables[0];

            string json = Newtonsoft.Json.JsonConvert.SerializeObject(dt, Newtonsoft.Json.Formatting.Indented);

            return json;
        }

        [HttpPost("config_email_test", Name = "config_email_test")]
        public string config_email_test(EmailConfigModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.config_email_test(param);

        }

        [HttpPost("uploadfile_data", Name = "uploadfile_data")]
        public string uploadfile_data([FromForm] uploadFile param)
        {
            string sub_software = "hazop";
            try { sub_software = param.sub_software; } catch { }

            ClassHazopSet cls = new ClassHazopSet();
            return cls.uploadfile_data(param, sub_software);

        }
        [HttpPost("uploadfile_data_followup", Name = "uploadfile_data_followup")]
        public string uploadfile_data_followup([FromForm] uploadFile param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.uploadfile_data(param, "followup");

        }
        [HttpPost("get_hazop_details", Name = "get_hazop_details")]
        public string get_hazop_details(LoadDocModel param)
        {
            param.sub_software = "hazop";
            ClassHazop cls = new ClassHazop();
            return cls.get_details(param);

        }
        [HttpPost("get_jsea_details", Name = "get_jsea_details")]
        public string get_jsea_details(LoadDocModel param)
        {
            param.sub_software = "jsea";
            ClassHazop cls = new ClassHazop();
            return cls.get_details(param);
        }
        [HttpPost("get_whatif_details", Name = "get_whatif_details")]
        public string get_whatif_details(LoadDocModel param)
        {
            param.sub_software = "whatif";
            ClassHazop cls = new ClassHazop();
            return cls.get_details(param);
        }


        [HttpPost("load_hazop_details", Name = "load_hazop_details")]
        public string load_hazop_details(LoadDocModel param)
        {
            ClassHazop cls = new ClassHazop();
            return cls.get_hazop_search(param);

        }

        [HttpPost("set_hazop", Name = "set_hazop")]
        public string set_hazop(SetDocHazopModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.set_hazop(param);
        }
        [HttpPost("set_jsea", Name = "set_jsea")]
        public string set_jsea(SetDocHazopModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.set_jsea(param);
        }
        
        [HttpPost("set_whatif", Name = "set_whatif")]
        public string set_whatif(SetDocHazopModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.set_whatif(param);
        }


        [HttpPost("set_master_ram", Name = "set_master_ram")]
        public string set_master_ram(SetDocHazopModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();

            return cls.set_master_ram(param);
        }
        [HttpPost("set_approve", Name = "set_approve")]
        public string set_approve(SetDocApproveModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.set_approve(param);
        }
        [HttpPost("set_email_member_review", Name = "set_email_member_review")]
        public string set_email_member_review(SetDocHazopModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.set_email_member_review(param);
        }
        [HttpPost("set_email_member_review_stamp", Name = "set_email_member_review_stamp")]
        public string set_email_member_review_stamp(SetDocHazopModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.set_email_member_review_stamp(param);
        }

        #region mail test
        [HttpPost("MailToPHAConduct", Name = "MailToPHAConduct")]
        public string MailToPHAConduct(string seq, string sub_software)
        {
            ClassEmail cls = new ClassEmail();
            return cls.MailToPHAConduct(seq, sub_software);
        }
        [HttpPost("MailToActionOwner", Name = "MailToActionOwner")]
        public string MailToActionOwner(string seq, string sub_software)
        {
            ClassEmail cls = new ClassEmail();
            return cls.MailToActionOwner(seq, sub_software);
        }
        #endregion mail test

        #region follow up  
        [HttpPost("load_follow_up", Name = "load_follow_up")]
        public string load_follow_up(LoadDocModel param)
        {
            ClassHazop cls = new ClassHazop();
            return cls.get_followup(param);

        }
        [HttpPost("load_follow_up_details", Name = "load_follow_up_details")]
        public string load_follow_up_details(LoadDocFollowModel param)
        {
            ClassHazop cls = new ClassHazop();
            return cls.get_followup_detail(param);

        }
        [HttpPost("set_follow_up", Name = "set_follow_up")]
        public string set_follow_up(SetDocHazopModel param)
        { 
            ClassHazopSet cls = new ClassHazopSet();
            return cls.set_follow_up(param);
        }
        [HttpPost("set_follow_up_review", Name = "set_follow_up_review")]
        public string set_follow_up_review(SetDocHazopModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.set_follow_up_review(param);
        }
        #endregion follow up 


        #region export hazop
        [HttpPost("export_hazop_report", Name = "export_hazop_report")]
        public string export_hazop_report(ReportModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.export_hazop_report(param);
        }

        [HttpPost("export_hazop_worksheet", Name = "export_hazop_worksheet")]
        public string export_hazop_worksheet(ReportModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.export_hazop_worksheet(param);
        }

        [HttpPost("export_hazop_recommendation", Name = "export_hazop_recommendation")]
        public string export_hazop_recommendation(ReportModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.export_hazop_recommendation(param);
        }

        [HttpPost("export_hazop_ram", Name = "export_hazop_ram")]
        public string export_hazop_ram(ReportModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.export_hazop_ram(param);
        }
        [HttpPost("export_hazop_guidewords", Name = "export_hazop_guidewords")]
        public string export_hazop_guidewords(ReportModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.export_hazop_guidewords(param);
        }


        #endregion export hazop


        #region export jsea

        [HttpPost("export_template_data", Name = "export_template_data")]
        public string export_template_data(ReportModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.export_template_data(param);
        }
        #endregion export jsea
        //[HttpPost("MailNotificationFollowUpItemToOwner", Name = "MailNotificationFollowUpItemToOwner")]
        //public string MailNotificationFollowUpItemToOwner(string seq)
        //{
        //    ClassEmail classEmail = new ClassEmail();
        //    return classEmail.MailNotificationFollowUpItemToOwner(seq);
        //}
        [HttpPost("MailNotificationDaily", Name = "MailNotificationDaily")]
        public string MailNotificationDaily(LoadDocModel param)
        {
            string user_name = (param.user_name + "");
            string seq = (param.token_doc + "");
            string sub_software = (param.sub_software + "");

            ClassEmail classEmail = new ClassEmail();
            return classEmail.MailNotificationDaily(user_name, seq, sub_software);
        }



        [HttpPost("copy_pdf_file", Name = "copy_pdf_file")]
        public string copy_pdf_file(CopyFileModel param)
        {
            ClassHazopSet cls = new ClassHazopSet();
            return cls.copy_pdf_file(param);
        }
        [HttpPost("mail_attachments", Name = "mail_attachments")]
        public string mail_attachments(string seq)
        {
            //{"sub_software":"hazop","user_name":"' + user_name + '","seq":"' + seq + '","export_type":"' + data_type + '"}
            ReportModel param = new ReportModel();
            param.seq = seq;
            param.export_type = "pdf";
            param.user_name = "";

            ClassHazopSet classHazopSet = new ClassHazopSet();
            string jsper = classHazopSet.export_hazop_recommendation(param);

            string file_ResponseSheet = "";
            DataTable dtReport = new DataTable();
            ClassJSON cls_json = new ClassJSON();
            dtReport = cls_json.ConvertJSONresult(jsper);
            if (dtReport.Rows.Count > 0)
            {
                file_ResponseSheet = (Path.Combine(Directory.GetCurrentDirectory(), "") + @"/wwwroot" + (dtReport.Rows[0]["ATTACHED_FILE_PATH"] + "").Replace("~", "").Replace(".xlsx", "." + param.export_type));
            }

            sendEmailModel data = new sendEmailModel();
            data.mail_subject = "test send file";
            data.mail_body = "xx";
            data.mail_to = "zkuluwat@thaioilgroup.com";
            data.mail_from = "zkuluwat@thaioilgroup.com";

            if (file_ResponseSheet != "")
            {
                data.mail_attachments = file_ResponseSheet;
            }
            if (data.mail_attachments == null) { return "not file attachments"; }

            ClassEmail classEmail = new ClassEmail();
            return classEmail.sendMail(data);
        }

        [HttpPost("MailTest", Name = "MailTest")]
        public string MailTest()
        {
            ClassEmail classEmail = new ClassEmail();
            return classEmail.MailTest();
        }
        [HttpPost("MailMS365Test", Name = "MailMS365Test")]
        public string MailMS365Test()
        {
            ClassEmail classEmail = new ClassEmail();
            return classEmail.MailMS365Test();
        }

    }
}
