using Aspose.Cells.Charts;
using dotnet6_epha_api.Class;
using Model;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SkiaSharp;
using System.Data;
using System.Xml.Linq;

namespace Class
{

    public class ClassNoti
    {
        string sqlstr = "";
        string jsper = "";
        ClassFunctions cls = new ClassFunctions();
        ClassJSON cls_json = new ClassJSON();
        ClassConnectionDb cls_conn = new ClassConnectionDb();

        DataSet dsData;
        DataTable dt, dtcopy, dtcheck;

        string[] sMonth = ("JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC").Split(',');

        private static DataTable refMsg(string status, string remark)
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            return dtMsg;
        }
        public static string Base64Encode(string text)
        {
            var textBytes = System.Text.Encoding.UTF8.GetBytes(text);
            return System.Convert.ToBase64String(textBytes);
        }
        public static string Base64Decode(string base64)
        {
            var base64Bytes = System.Convert.FromBase64String(base64);
            return System.Text.Encoding.UTF8.GetString(base64Bytes);
        }
        private int get_max(string table_name)
        {
            DataTable _dt = new DataTable();
            cls = new ClassFunctions();

            sqlstr = @" select coalesce(max(id),0)+1 as id from " + table_name;

            cls_conn = new ClassConnectionDb();
            _dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            return Convert.ToInt32(_dt.Rows[0]["id"].ToString() + "");
        }

        public void DataNotification_Daily(LoadDocModel param)
        {
            string user_name = (param.user_name + "").Trim();
            string token_doc = (param.token_doc + "").Trim();
            string sub_software = (param.sub_software + "").Trim();
            string type_doc = (param.type_doc + "").Trim();//review_document
            string seq = token_doc;

            DataDailyByActionRequired(user_name, seq, sub_software, false);
        }
        public DataTable DataDailyByActionRequired(string user_name, string seq, string sub_software, Boolean group_by_user)
        {
            dt = new DataTable();
            cls = new ClassFunctions();

            #region get data 
            //Recommendation Closing   
            sqlstr = "";
            sqlstr += @"select a.id as id_pha, a.pha_status
                     , isnull(nw.responder_user_name,'') as user_name, emp.user_displayname, emp.user_email
                     , isnull(a.pha_request_by,'') as user_name_ori 
                     , nw.seq as id_action, nw.responder_action_date as user_action_date
                     , 1 as action_sort
                     , 0 as task, upper(a.pha_sub_software) as pha_type
                     , 'Recommendation Closing' as action_required, a.pha_no as document_number, g.pha_request_name as document_title
                     , nw.recommendations_no as rev, emp_ori.user_displayname as originator
                     , format(nw.responder_receivesd_date,'dd MMM yyyy') as receivesd, format(nw.estimated_start_date,'dd MMM yyyy') as due_date 
                     , nw.responder_action_date as action_date
                     , isnull(datediff(day, getdate(),case when nw.estimated_start_date >= getdate() then nw.estimated_start_date else getdate() end),0) as remaining
                     , emp_conso.user_displayname as consolidator 
                     from EPHA_F_HEADER a  
                     inner join EPHA_T_GENERAL g on a.id = g.id_pha   
                     inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha ) s on a.id = s.id_pha 
                     inner join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha ) t2 on a.id = t2.id_pha and s.id_session = t2.id_session 
                     inner join EPHA_T_NODE_WORKSHEET nw on a.id = nw.id_pha   
                     left join VW_EPHA_PERSON_DETAILS emp on lower(nw.responder_user_name) = lower(emp.user_name)
                     left join VW_EPHA_PERSON_DETAILS emp_ori on lower(a.pha_request_by) = lower(emp_ori.user_name)
                     left join VW_EPHA_PERSON_DETAILS emp_conso on lower(nw.responder_user_name) = lower(emp_conso.user_name)
                     where nw.responder_user_name is not null and nw.estimated_start_date is not null and nw.responder_action_date is null ";

            //Review -> member team 
            sqlstr += @" union ";
            sqlstr += @"select a.id as id_pha, a.pha_status 
                     , isnull(t.user_name,'') as user_name, emp.user_displayname, emp.user_email
                     , isnull(a.pha_request_by,'') as user_name_ori 
                     , t.seq as id_action, t.date_review as user_action_date
                     , 2 as action_sort
                     , 0 as task, upper(a.pha_sub_software) as pha_type
                     , 'Review' as action_required, a.pha_no as document_number, g.pha_request_name as document_title
                     , t.no as rev, emp_ori.user_displayname as originator
                     , format(s.date_to_review,'dd MMM yyyy') as receivesd, format(dateadd(day,5,s.date_to_review) ,'dd MMM yyyy') as due_date
                     , t.date_review as action_date
                     , isnull(datediff(day, getdate(),case when (dateadd(day,5,s.date_to_review)) >= getdate() then dateadd(day,5,s.date_to_review) else getdate() end),0) as remaining
                     , emp_conso.user_displayname as consolidator 
                     from EPHA_F_HEADER a  
                     inner join EPHA_T_GENERAL g on a.id = g.id_pha   
                     inner join EPHA_T_SESSION s on a.id = s.id_pha  
                     inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha ) s2 on a.id = s2.id_pha 
                     inner join EPHA_T_MEMBER_TEAM t on a.id = t.id_pha and s2.id_session = t.id_session
                     inner join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha ) t2 on a.id = t2.id_pha and s2.id_session = t2.id_session 
                     left join VW_EPHA_PERSON_DETAILS emp on lower(t.user_name) = lower(emp.user_name) 
                     left join VW_EPHA_PERSON_DETAILS emp_ori on lower(a.pha_request_by) = lower(emp_ori.user_name)
                     left join VW_EPHA_PERSON_DETAILS emp_conso on lower(t.user_name) = lower(emp_conso.user_name)
                     where t.user_name is not null and s.date_to_review is not null and t.date_review is null
                     and s.action_to_review = 2  ";

            //Approve -> Originator review action owner followup
            sqlstr += @" union ";
            sqlstr += @" select a.id as id_pha, a.pha_status
                     , isnull(a.pha_request_by,'') as user_name, emp.user_displayname, emp.user_email
                     , isnull(a.pha_request_by,'') as user_name_ori 
                     , nw.seq as id_action, null as user_action_date
                     , 3 as action_sort
                     , 0 as task, upper(a.pha_sub_software) as pha_type
                     , 'Approve' as action_required, a.pha_no as document_number, g.pha_request_name as document_title
                     , nw.recommendations_no as rev, emp_ori.user_displayname as originator
                     , format(nw.responder_action_date,'dd MMM yyyy') as receivesd, format(g.target_end_date,'dd MMM yyyy') as due_date 
                     , null as action_date
                     , isnull(datediff(day, getdate(),case when g.target_end_date >= getdate() then g.target_end_date else getdate() end),0) as remaining
                     , emp_conso.user_displayname as consolidator 
                     from EPHA_F_HEADER a  
                     inner join EPHA_T_GENERAL g on a.id = g.id_pha   
                     inner join EPHA_T_SESSION s on a.id = s.id_pha 
                     inner join (select max(id) as id_session, id_pha from EPHA_T_SESSION group by id_pha ) s2 on a.id = s2.id_pha and s.id = s2.id_session 
                     inner join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha ) t2 on a.id = t2.id_pha and s2.id_session = t2.id_session 
                     inner join EPHA_T_NODE_WORKSHEET nw on a.id = nw.id_pha   
                     left join VW_EPHA_PERSON_DETAILS emp on lower(a.pha_request_by) = lower(emp.user_name)
                     left join VW_EPHA_PERSON_DETAILS emp_ori on lower(a.pha_request_by) = lower(emp_ori.user_name)
                     left join VW_EPHA_PERSON_DETAILS emp_conso on lower(nw.responder_user_name) = lower(emp_conso.user_name)
                     where a.pha_status in (12,13) and nw.responder_user_name is not null and nw.responder_action_date is not null and lower(nw.action_status) not in ('closed') ";


            if (group_by_user == true)
            {
                sqlstr = "select distinct t.user_name,t.user_displayname, t.user_email from (" + sqlstr + ")t where t.user_name is not null order by t.user_name";
            }
            else
            {
                sqlstr = "select t.* from (" + sqlstr + ")t where t.pha_status in (12,13)  ";
                if (user_name != "") { sqlstr += " and lower(t.user_name) = lower(" + cls.ChkSqlStr(user_name, 50) + ")"; }
                if (seq != "") { sqlstr += " and lower(t.id_pha) = lower(" + cls.ChkSqlNum(seq, "N") + ") "; }
                sqlstr += "  order by t.user_name, t.action_sort, t.document_number, t.rev";
            }


            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            #endregion get data 

            return dt;
        }



    }
}
