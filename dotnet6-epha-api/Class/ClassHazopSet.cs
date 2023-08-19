using dotnet6_epha_api.Class;
using Model;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Data;
using System.Xml.Linq;

namespace Class
{

    public class ClassHazopSet
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
        private string MapPathFiles(string _folder)
        {
            return (Path.Combine(Directory.GetCurrentDirectory(), "") + _folder.Replace("~", ""));
        }
        public string uploadfile_data(uploadFile uploadFile)
        {
            DataTable dtdef = new DataTable();
            IFormFileCollection files = uploadFile.file_obj;
            var file_seq = uploadFile.file_seq;
            var file_name = uploadFile.file_name;

            var file_FullName = "";
            var file_FullPath = "";

            string _Folder = "/wwwroot/AttachedFileTemp/FollowUp/";
            string _DownloadPath = "/AttachedFileTemplate/FollowUp/";
            string _Path = MapPathFiles(_Folder);

            #region Determine whether the directory exists.
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.AcceptChanges();
            dtdef = dt.Clone(); dtdef.AcceptChanges();

            string msg_error = "";

            try
            {
                DataRow dr = dt.NewRow();
                if (!Directory.Exists(_Path))
                {
                    Directory.CreateDirectory(_Path);
                }

                //delete all files and folders in a directory
                System.IO.DirectoryInfo di = new DirectoryInfo(_Path);

                //foreach (FileInfo file in di.GetFiles())
                //{ 
                //    file.Delete();
                //}

                for (int i = 0; i < files.Count; i++)
                {
                    //*** ต้องเปลี่ยนวิธี
                    IFormFile file = files[i];
                    file_FullName = file.FileName;
                    file_FullPath = _Path + file_FullName; //MapPathFiles("~/AttachedFile/Plan/" + file.FileName);
                    using (Stream fileStream = new FileStream(file_FullPath, FileMode.Create))
                    {
                        file.CopyTo(fileStream);
                    }
                    dr["ATTACHED_FILE_NAME"] = file.FileName;
                    dr["ATTACHED_FILE_PATH"] = _DownloadPath + file.FileName;
                }

                dt.Rows.Add(dr);
                dt.AcceptChanges();
                dtdef = dt.Copy(); dtdef.AcceptChanges();
            }
            catch (Exception ex) { msg_error = ex.Message.ToString(); }

            #endregion Determine whether the directory exists.

            try
            {
                dtdef.Rows.Add(dtdef.NewRow()); dtdef.AcceptChanges();
                dtdef.Rows[dtdef.Rows.Count - 1]["IMPORT_DATA_MSG"] = msg_error;
                dtdef.AcceptChanges();
            }
            catch (Exception ex) { ex.Message.ToString(); }

            return cls_json.SetJSONresult(dtdef);
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
        private void ConvertJSONresultToDataSet(ref string msg, ref string ret, ref DataSet dsData, SetDocHazopModel param, string pha_status)
        {
            #region ConvertJSONresult

            jsper = param.json_header + "";
            if (jsper.Trim() == "") { msg = "No Data."; ret = "Error"; return; }
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);

                dt.TableName = "header";
                dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }


            jsper = param.json_general + "";
            if (jsper.Trim() == "") { msg = "No Data."; ret = "Error"; return; }
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "general";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_functional_audition + "";
            try
            {
                if (jsper.Trim() != "")
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "functional_audition";
                        dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                    }
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }


            jsper = param.json_session + "";
            if (jsper.Trim() == "") { msg = "No Data."; ret = "Error"; return; }
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "session";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_memberteam + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "memberteam";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_drawing + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "drawing";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_node + "";
            try
            {
                if (jsper.Trim() != "")
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "node";
                        dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                    }
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_nodedrawing + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "nodedrawing";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_nodeguidwords + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "nodeguidwords";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            if (pha_status == "11") { goto Next_Line_Data; }

            jsper = param.json_nodeworksheet + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "nodeworksheet";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            jsper = param.json_managerecom + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "managerecom";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

        Next_Line_Data:;
            #endregion ConvertJSONresult

        }
        public string set_hazop(SetDocHazopModel param)
        {
            string msg = "";
            string ret = "";
            cls_json = new ClassJSON();

            DataSet dsData = new DataSet();
            string seq_header = (param.token_doc + "");
            string pha_status = (param.pha_status + "");
            string pha_version = (param.pha_version + "");
            string user_name = (param.user_name + "");
            string seq = (param.token_doc + "");


            //$scope.flow_role_type = "admin";//admin,request,responder,approver
            string role_type = ("admin");
            Boolean bOwnerAction = true;//เป็น Owner Action ด้วยหรือป่าว

            //11	DF	Draft
            //12	WP	PHA Conduct 
            //21	WA	Waiting Approve Review
            //22	AR	Approve Reject
            //13	WF	Waiting Follow Up
            //14	WR	Waiting Review Follow Up
            //91	CL	Closed
            //81	CN	Cancle

            ConvertJSONresultToDataSet(ref msg, ref ret, ref dsData, param, pha_status);
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }

            //action type = insert,update,delete,old_data 
            string year_header_now = DateTime.Now.ToString("YYYY");
            string seq_header_now = get_max("EPHA_F_HEADER").ToString();

            if (pha_status == "22")
            {
                //ตรวจสอบว่า seq นี้เป็น version ล่าสุดหรือไม่
                Boolean pha_new_version = false;
                sqlstr = @" select max(a.pha_version) as pha_version from EPHA_F_HEADER a where lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";
                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["pha_version"].ToString() != pha_version) { pha_new_version = true; }
                }
                if (pha_new_version == true)
                {
                    //ต้อง copy เป็น version ใหม่
                    dsData.Tables["header"].Rows[0]["seq"] = seq_header_now;
                    dsData.Tables["header"].Rows[0]["id"] = seq_header_now;
                    dsData.Tables["header"].Rows[0]["pha_version"] = (Convert.ToInt32(dsData.Tables["header"].Rows[0]["pha_version"] + "") + 1);
                    dsData.Tables["header"].Rows[0]["action_by"] = "insert";

                    dsData.Tables["general"].Rows[0]["id_pha"] = seq_header_now;
                    dsData.Tables["general"].Rows[0]["action_by"] = "insert";

                    for (int i = 0; i < dsData.Tables["node"].Rows.Count; i++) { dsData.Tables["node"].Rows[0]["id_pha"] = seq_header_now; dsData.Tables["node"].Rows[0]["action_by"] = "insert"; }
                    for (int i = 0; i < dsData.Tables["nodeworksheet"].Rows.Count; i++) { dsData.Tables["nodeworksheet"].Rows[0]["id_pha"] = seq_header_now; dsData.Tables["nodeworksheet"].Rows[0]["action_by"] = "insert"; }
                    for (int i = 0; i < dsData.Tables["managerecom"].Rows.Count; i++) { dsData.Tables["managerecom"].Rows[0]["id_pha"] = seq_header_now; dsData.Tables["managerecom"].Rows[0]["action_by"] = "insert"; }
                    dsData.AcceptChanges();
                }
            }
            else
            {
                seq_header_now = seq;
            }

            ClassHazop cls_old = new ClassHazop();
            DataSet dsDataOld = new DataSet();

            if (dsData.Tables["header"].Rows.Count > 0)
            {
                #region connection transaction
                cls = new ClassFunctions();
                ClassConnectionDb cls_conn_header = new ClassConnectionDb();
                ClassConnectionDb cls_conn_node = new ClassConnectionDb();
                ClassConnectionDb cls_conn_worksheet = new ClassConnectionDb();
                ClassConnectionDb cls_conn_managerecom = new ClassConnectionDb();

                cls_conn = new ClassConnectionDb();
                cls_conn_header = new ClassConnectionDb();
                cls_conn_node = new ClassConnectionDb();
                cls_conn_worksheet = new ClassConnectionDb();
                cls_conn_managerecom = new ClassConnectionDb();

                cls_conn.OpenConnection();
                cls_conn_header.OpenConnection();
                cls_conn_node.OpenConnection();
                cls_conn_worksheet.OpenConnection();
                cls_conn_managerecom.OpenConnection();

                cls_conn.BeginTransaction();
                cls_conn_header.BeginTransaction();
                cls_conn_node.BeginTransaction();
                cls_conn_worksheet.BeginTransaction();
                cls_conn_managerecom.BeginTransaction();

                #endregion connection transaction
                try
                {
                    if (pha_status == "11")
                    {
                        ret = set_hazop_header(ref dsData, ref cls_conn_header, seq_header_now);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }
                    }

                    if (pha_status == "11" || pha_status == "22")
                    {
                        ret = set_hazop_parti(ref dsData, ref cls_conn_header, seq_header_now, dsDataOld);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }

                        ret = set_hazop_partii(ref dsData, ref cls_conn_node, seq_header_now);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }
                    }

                    if (pha_status == "12" || pha_status == "22")
                    {
                        ret = set_hazop_partii(ref dsData, ref cls_conn_node, seq_header_now);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }

                        ret = set_hazop_partiii(ref dsData, ref cls_conn_worksheet, seq_header_now);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }

                        ret = set_hazop_partiv(ref dsData, ref cls_conn_managerecom, seq_header_now);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }
                    }

                }
                catch (Exception ex) { ret = ex.Message.ToString(); goto Next_Line; }

            Next_Line:;

                #region connection transaction
                if (ret == "") { ret = "true"; }
                if (ret == "true")
                {
                    cls_conn_header.CommitTransaction();
                    cls_conn_node.CommitTransaction();
                    cls_conn_worksheet.CommitTransaction();
                    cls_conn_managerecom.CommitTransaction();

                    cls_conn.CommitTransaction();
                }
                else
                {
                    cls_conn_header.RollbackTransaction();
                    cls_conn_node.RollbackTransaction();
                    cls_conn_worksheet.RollbackTransaction();
                    cls_conn_managerecom.RollbackTransaction();

                    cls_conn.RollbackTransaction();
                }
                cls_conn_header.CloseConnection();
                cls_conn_node.CloseConnection();
                cls_conn_worksheet.CloseConnection();
                cls_conn_managerecom.CloseConnection();

                cls_conn.CloseConnection();
                #endregion connection transaction

                #region  flow action  submit  
                if (ret == "true")
                {
                    //11	DF	Draft
                    //12	WP	PHA Conduct 
                    //21	WA	Waiting Approve Review
                    //22	AR	Approve Reject
                    //13	WF	Waiting Follow Up
                    //14	WR	Waiting Review Follow Up
                    //91	CL	Closed
                    //81	CN	Cancle

                    if (param.flow_action == "submit")
                    {
                        ClassEmail clsmail = new ClassEmail();
                        if (pha_status == "11")
                        {

                            cls = new ClassFunctions();
                            cls_conn = new ClassConnectionDb();
                            cls_conn.OpenConnection();
                            cls_conn.BeginTransaction();

                            int i = 0;
                            dt = new DataTable();
                            dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();

                            string pha_status_new = "12";

                            #region update
                            sqlstr = "update  EPHA_F_HEADER set ";
                            sqlstr += " PHA_STATUS = " + cls.ChkSqlNum((pha_status_new).ToString(), "N");

                            sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                            sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                            sqlstr += " and YEAR = " + cls.ChkSqlNum((dt.Rows[i]["YEAR"] + "").ToString(), "N");
                            sqlstr += " and PHA_NO = " + cls.ChkSqlStr((dt.Rows[i]["PHA_NO"] + "").ToString(), 200);

                            #endregion update

                            ret = cls_conn.ExecuteNonQuery(sqlstr);
                            if (ret == "") { ret = "true"; }
                            if (ret == "true")
                            {
                                cls_conn.CommitTransaction();
                            }
                            else
                            {
                                cls_conn.RollbackTransaction();
                            }
                            cls_conn.CloseConnection();

                            clsmail = new ClassEmail();
                            clsmail.MailToPHAConduct((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");

                        }
                        else if (pha_status == "12")
                        {
                            //12	WP	PHA Conduct 
                            cls = new ClassFunctions();
                            cls_conn = new ClassConnectionDb();
                            cls_conn.OpenConnection();
                            cls_conn.BeginTransaction();

                            int i = 0;
                            dt = new DataTable();
                            dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();

                            //13	WF	Waiting Follow Up
                            string pha_status_new = "13";
                            if (dt.Rows[0]["request_approver"].ToString() == "1")
                            {
                                //21	WA	Waiting Approve Review
                                pha_status_new = "21";
                            }

                            #region update
                            sqlstr = "update  EPHA_F_HEADER set ";
                            sqlstr += " PHA_STATUS = " + cls.ChkSqlNum((pha_status_new).ToString(), "N");

                            sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                            sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                            sqlstr += " and YEAR = " + cls.ChkSqlNum((dt.Rows[i]["YEAR"] + "").ToString(), "N");
                            sqlstr += " and PHA_NO = " + cls.ChkSqlStr((dt.Rows[i]["PHA_NO"] + "").ToString(), 200);

                            #endregion update

                            ret = cls_conn.ExecuteNonQuery(sqlstr);
                            if (ret == "") { ret = "true"; }
                            if (ret == "true")
                            {
                                cls_conn.CommitTransaction();
                            }
                            else
                            {
                                cls_conn.RollbackTransaction();
                            }
                            cls_conn.CloseConnection();


                            if (pha_status_new == "13")
                            {
                                //13	WF	Waiting Follow Up
                                clsmail = new ClassEmail();
                                clsmail.MailToActionOwner((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");
                            }
                            else
                            {
                                //21	WA	Waiting Approve Review
                                clsmail = new ClassEmail();
                                clsmail.MailToApproverReview((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");
                            }

                        }
                        else if (pha_status == "21")
                        {
                            //ต้อง copy เป็น version ใหม่
                            cls = new ClassFunctions();
                            cls_conn = new ClassConnectionDb();
                            cls_conn.OpenConnection();
                            cls_conn.BeginTransaction();

                            int i = 0;
                            dt = new DataTable();
                            dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();

                            string pha_status_new = "22";

                            #region update
                            sqlstr = "update  EPHA_F_HEADER set ";
                            sqlstr += " PHA_STATUS = " + cls.ChkSqlNum((pha_status_new).ToString(), "N");

                            sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                            sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                            sqlstr += " and YEAR = " + cls.ChkSqlNum((dt.Rows[i]["YEAR"] + "").ToString(), "N");
                            sqlstr += " and PHA_NO = " + cls.ChkSqlStr((dt.Rows[i]["PHA_NO"] + "").ToString(), 200);

                            #endregion update

                            ret = cls_conn.ExecuteNonQuery(sqlstr);
                            if (ret == "") { ret = "true"; }
                            if (ret == "true")
                            {
                                cls_conn.CommitTransaction();
                            }
                            else
                            {
                                cls_conn.RollbackTransaction();
                            }
                            cls_conn.CloseConnection();

                            //13	WF	Waiting Follow Up
                            if (pha_status_new == "13")
                            {
                                clsmail = new ClassEmail();
                                clsmail.MailToActionOwner((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");
                            }
                            else
                            {
                                //22	AR	Approve Reject
                                clsmail = new ClassEmail();
                                clsmail.MailRejectByApprover((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");
                            }

                        }
                        else if (pha_status == "22")
                        {
                            //ต้อง copy เป็น version ใหม่
                            cls = new ClassFunctions();
                            cls_conn = new ClassConnectionDb();
                            cls_conn.OpenConnection();
                            cls_conn.BeginTransaction();

                            int i = 0;
                            dt = new DataTable();
                            dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();

                            string pha_status_new = "21";

                            #region update
                            sqlstr = "update  EPHA_F_HEADER set ";
                            sqlstr += " PHA_STATUS = " + cls.ChkSqlNum((pha_status_new).ToString(), "N");

                            sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                            sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                            sqlstr += " and YEAR = " + cls.ChkSqlNum((dt.Rows[i]["YEAR"] + "").ToString(), "N");
                            sqlstr += " and PHA_NO = " + cls.ChkSqlStr((dt.Rows[i]["PHA_NO"] + "").ToString(), 200);

                            #endregion update

                            ret = cls_conn.ExecuteNonQuery(sqlstr);
                            if (ret == "") { ret = "true"; }
                            if (ret == "true")
                            {
                                cls_conn.CommitTransaction();
                            }
                            else
                            {
                                cls_conn.RollbackTransaction();
                            }
                            cls_conn.CloseConnection();

                            //21	WA	Waiting Approve Review
                            clsmail = new ClassEmail();
                            clsmail.MailToApproverReview((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");
                        }
                        else if (pha_status == "13" && false)
                        {

                            int i = 0;
                            dt = new DataTable();
                            dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();

                            string pha_status_new = "14";


                            #region check status follow up -> update status all 

                            DataTable dtaction = new DataTable();
                            ClassConnectionDb cls_conn = new ClassConnectionDb();
                            sqlstr = @" select count(1) as total, count(case when lower(a.action_status) = 'open' then 1 else null end) 'open'
                                        from EPHA_T_NODE_WORKSHEET a where a.id_pha = " + (dt.Rows[i]["SEQ"] + "").ToString();

                            dtaction = new DataTable();
                            dtaction = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
                            #endregion check status follow up -> update status all 

                            if (dtaction.Rows.Count > 0)
                            {
                                if ((dtaction.Rows[0]["total"] + "") != "0")
                                {
                                    if ((dtaction.Rows[0]["open"] + "") == "0")
                                    {
                                        //ต้อง copy เป็น version ใหม่
                                        cls = new ClassFunctions();
                                        cls_conn = new ClassConnectionDb();
                                        cls_conn.OpenConnection();
                                        cls_conn.BeginTransaction();

                                        #region update
                                        sqlstr = "update  EPHA_F_HEADER set ";
                                        sqlstr += " PHA_STATUS = " + cls.ChkSqlNum((pha_status_new).ToString(), "N");

                                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                                        sqlstr += " and YEAR = " + cls.ChkSqlNum((dt.Rows[i]["YEAR"] + "").ToString(), "N");
                                        sqlstr += " and PHA_NO = " + cls.ChkSqlStr((dt.Rows[i]["PHA_NO"] + "").ToString(), 200);

                                        #endregion update

                                        ret = cls_conn.ExecuteNonQuery(sqlstr);
                                        if (ret == "") { ret = "true"; }
                                        if (ret == "true")
                                        {
                                            cls_conn.CommitTransaction();
                                        }
                                        else
                                        {
                                            cls_conn.RollbackTransaction();
                                        }
                                        cls_conn.CloseConnection();


                                        //14	WR	Waiting Review Follow Up
                                        clsmail = new ClassEmail();
                                        clsmail.MailNotificationToAdminOwnerUpdateAction((dt.Rows[i]["SEQ"] + "").ToString(), "", "hazop");
                                    }
                                    else
                                    {
                                        if (role_type != "admin" || bOwnerAction)
                                        {
                                            //Check by Action Owner  
                                            cls_conn = new ClassConnectionDb();
                                            sqlstr = @" select count(1) as total, count(case when lower(a.action_status) = 'open' then 1 else null end) 'open'
                                                    , emp.user_displayname, emp.user_email,
                                                     from EPHA_T_NODE_WORKSHEET a 
                                                     left join EPHA_PERSON_DETAILS emp on lower(a.responder_user_name) = lower(emp.user_name)  
                                                     where a.id_pha = " + (dt.Rows[i]["SEQ"] + "").ToString();
                                            sqlstr += @" and lower(a.responder_user_name)  = lower('" + user_name + "')";
                                            sqlstr += @" group by emp.user_displayname, emp.user_email";

                                            dtaction = new DataTable();
                                            dtaction = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

                                            if ((dtaction.Rows[0]["open"] + "") == "0")
                                            {
                                                //mail not admin กรณีที่ Action Owner Update status Closed All 
                                                //Notification: Mr. Kun has updated the statuses of all tasks.
                                                clsmail = new ClassEmail();
                                                clsmail.MailNotificationToAdminReviewByOwner((dt.Rows[i]["SEQ"] + "").ToString(), user_name, "hazop");
                                            }
                                        }

                                    }

                                }
                            }

                        }
                        else if (pha_status == "14")
                        {
                            //ต้อง copy เป็น version ใหม่
                            cls = new ClassFunctions();
                            cls_conn = new ClassConnectionDb();
                            cls_conn.OpenConnection();
                            cls_conn.BeginTransaction();

                            int i = 0;
                            dt = new DataTable();
                            dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();

                            string pha_status_new = "91";

                            #region update
                            sqlstr = "update  EPHA_F_HEADER set ";
                            sqlstr += " PHA_STATUS = " + cls.ChkSqlNum((pha_status_new).ToString(), "N");

                            sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                            sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                            sqlstr += " and YEAR = " + cls.ChkSqlNum((dt.Rows[i]["YEAR"] + "").ToString(), "N");
                            sqlstr += " and PHA_NO = " + cls.ChkSqlStr((dt.Rows[i]["PHA_NO"] + "").ToString(), 200);

                            #endregion update

                            ret = cls_conn.ExecuteNonQuery(sqlstr);
                            if (ret == "") { ret = "true"; }
                            if (ret == "true")
                            {
                                cls_conn.CommitTransaction();
                            }
                            else
                            {
                                cls_conn.RollbackTransaction();
                            }
                            cls_conn.CloseConnection();


                            //91	CL	Closed
                            clsmail = new ClassEmail();
                            clsmail.MailToAllUserClosed((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");
                        }

                    }
                }
                #endregion  flow action  submit 

            }

        Next_Line_Convert:;
            return cls_json.SetJSONresult(refMsg(ret, msg));
        }

        private string SeqTypeDelete(DataTable dt, DataTable dtOld)
        {
            //data type delete
            string seq_delete = "";
            Boolean bDataNow = (dt.Rows.Count > 0 ? true : false);

            if (bDataNow == true)
            {
                for (int n = 0; n < dtOld.Rows.Count; n++)
                {
                    string seq_def = (dtOld.Rows[n]["seq"] + "").ToString();
                    for (int m = 0; m < dt.Rows.Count; m++)
                    {
                        if (seq_def == (dt.Rows[m]["seq"] + "").ToString())
                        {
                            continue;
                        }
                        if (seq_delete != "") { seq_delete += ","; }
                        seq_delete = seq_def;
                    }
                }
            }
            else
            {
                for (int n = 0; n < dtOld.Rows.Count; n++)
                {
                    string seq_def = (dtOld.Rows[n]["seq"] + "").ToString();
                    if (seq_delete != "") { seq_delete += ","; }
                    seq_delete = seq_def;
                }
            }
            return seq_delete;
        }
        public string set_hazop_header(ref DataSet dsData, ref ClassConnectionDb cls_conn, string seq_header_now)
        {
            string ret = "";

            #region update data header
            dt = new DataTable();
            dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                if (action_type == "insert")
                {
                    string pha_version = (dt.Rows[i]["PHA_VERSION"] + "").ToString();
                    if (pha_version == "0") { pha_version = "1"; }

                    #region insert
                    //SEQ Auto running
                    sqlstr = "insert into EPHA_F_HEADER(SEQ,ID,YEAR,PHA_NO,PHA_VERSION,PHA_STATUS,PHA_REQUEST_BY,PHA_SUB_SOFTWARE" +
                        ",REQUEST_APPROVER,APPROVER_USER_NAME,APPROVER_USER_DISPLAYNAME,APPROVE_ACTION_TYPE,APPROVE_STATUS,APPROVE_COMMENT" +
                        ",REQUEST_USER_NAME,REQUEST_USER_DISPLAYNAME" +
                        ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                        ") values ";
                    sqlstr += " ( ";
                    sqlstr += " " + cls.ChkSqlNum(seq_header_now, "N");
                    sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");
                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["YEAR"] + "").ToString(), "N");
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["PHA_NO"] + "").ToString(), 200);
                    sqlstr += " ," + cls.ChkSqlNum(pha_version, "N");
                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["PHA_STATUS"] + "").ToString(), "N");
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["PHA_REQUEST_BY"] + "").ToString(), 200);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["PHA_SUB_SOFTWARE"] + "").ToString(), 200);

                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["REQUEST_APPROVER"] + "").ToString(), "N");
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["APPROVER_USER_NAME"] + "").ToString(), 50);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["APPROVER_USER_DISPLAYNAME"] + "").ToString(), 4000);
                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["APPROVE_ACTION_TYPE"] + "").ToString(), "N");
                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["APPROVE_STATUS"] + "").ToString(), "N");
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["APPROVE_COMMENT"] + "").ToString(), 200);

                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["REQUEST_USER_NAME"] + "").ToString(), 50);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["REQUEST_USER_DISPLAYNAME"] + "").ToString(), 4000);

                    sqlstr += " ,getdate()";
                    sqlstr += " ,null";
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CREATE_BY"] + "").ToString(), 50);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                    sqlstr += ")";
                    #endregion insert
                }
                else if (action_type == "update")
                {
                    seq_header_now = (dt.Rows[i]["seq"] + "").ToString();

                    #region update

                    sqlstr = "update  EPHA_F_HEADER set ";
                    sqlstr += " PHA_VERSION = " + cls.ChkSqlNum((dt.Rows[i]["PHA_VERSION"] + "").ToString(), "N");
                    sqlstr += " ,PHA_STATUS = " + cls.ChkSqlNum((dt.Rows[i]["PHA_STATUS"] + "").ToString(), "N");
                    sqlstr += " ,PHA_REQUEST_BY = " + cls.ChkSqlStr((dt.Rows[i]["PHA_REQUEST_BY"] + "").ToString(), 200);
                    sqlstr += " ,PHA_SUB_SOFTWARE = " + cls.ChkSqlStr((dt.Rows[i]["PHA_SUB_SOFTWARE"] + "").ToString(), 200);

                    if ((dt.Rows[i]["REQUEST_APPROVER"] + "").ToString() == "1")
                    {
                        sqlstr += " ,REQUEST_APPROVER = " + cls.ChkSqlNum((dt.Rows[i]["REQUEST_APPROVER"] + "").ToString(), "N");
                        sqlstr += " ,APPROVER_USER_NAME = " + cls.ChkSqlStr((dt.Rows[i]["APPROVER_USER_NAME"] + "").ToString(), 50);
                        sqlstr += " ,APPROVER_USER_DISPLAYNAME = " + cls.ChkSqlStr((dt.Rows[i]["APPROVER_USER_DISPLAYNAME"] + "").ToString(), 4000);
                        sqlstr += " ,APPROVE_ACTION_TYPE = " + cls.ChkSqlNum((dt.Rows[i]["APPROVE_ACTION_TYPE"] + "").ToString(), "N");
                        sqlstr += " ,APPROVE_STATUS = " + cls.ChkSqlNum((dt.Rows[i]["APPROVE_STATUS"] + "").ToString(), "N");
                        sqlstr += " ,APPROVE_COMMENT = " + cls.ChkSqlStr((dt.Rows[i]["APPROVE_COMMENT"] + "").ToString(), 200);
                    }

                    sqlstr += " ,REQUEST_USER_NAME = " + cls.ChkSqlStr((dt.Rows[i]["REQUEST_USER_NAME"] + "").ToString(), 50);
                    sqlstr += " ,REQUEST_USER_DISPLAYNAME = " + cls.ChkSqlStr((dt.Rows[i]["REQUEST_USER_DISPLAYNAME"] + "").ToString(), 4000);


                    sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                    sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                    sqlstr += " and YEAR = " + cls.ChkSqlNum((dt.Rows[i]["YEAR"] + "").ToString(), "N");
                    sqlstr += " and PHA_NO = " + cls.ChkSqlStr((dt.Rows[i]["PHA_NO"] + "").ToString(), 200);

                    #endregion update
                }
                else if (action_type == "delete")
                {
                    #region delete
                    sqlstr = "delete from EPHA_F_HEADER ";

                    sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                    sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                    sqlstr += " and YEAR = " + cls.ChkSqlNum((dt.Rows[i]["YEAR"] + "").ToString(), "N");
                    sqlstr += " and PHA_NO = " + cls.ChkSqlStr((dt.Rows[i]["PHA_NO"] + "").ToString(), 200);
                    #endregion delete
                }

                if (action_type != "")
                {
                    ret = cls_conn.ExecuteNonQuery(sqlstr);
                    if (ret != "true") { break; }
                }
            }

            #endregion update data header

            return ret;
        }
        public string set_hazop_parti(ref DataSet dsData, ref ClassConnectionDb cls_conn, string seq_header_now, DataSet dsDataOld)
        {
            DataTable dtMainDelete = new DataTable();
            dtMainDelete.Columns.Add("SEQ", typeof(string));
            dtMainDelete.Columns.Add("ID", typeof(string));
            dtMainDelete.Columns.Add("ID_PHA", typeof(string));
            dtMainDelete.Columns.Add("ID_SESSION", typeof(string));

            string ret = "";

            #region update data general
            dt = new DataTable();
            dt = dsData.Tables["general"].Copy(); dt.AcceptChanges();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                if (action_type == "insert")
                {
                    #region insert
                    //SEQ Auto running
                    sqlstr = "insert into EPHA_T_GENERAL (" +
                        "SEQ,ID,ID_PHA,ID_RAM,EXPENSE_TYPE,SUB_EXPENSE_TYPE,REFERENCE_MOC  " +
                        ",ID_AREA,ID_APU,ID_BUSINESS_UNIT,ID_UNIT_NO,OTHER_AREA,OTHER_APU,OTHER_BUSINESS_UNIT,OTHER_UNIT_NO,FUNCTIONAL_LOCATION  " +
                        ",PHA_REQUEST_NAME,TARGET_START_DATE,TARGET_END_DATE,ACTUAL_START_DATE,ACTUAL_END_DATE  " +
                        ",DESCRIPTIONS" +
                        ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                        ") values ";
                    sqlstr += " ( ";
                    sqlstr += " " + cls.ChkSqlNum(seq_header_now, "N");
                    sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");
                    sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");
                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_RAM"] + "").ToString(), "N");
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["EXPENSE_TYPE"] + "").ToString(), 50);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["SUB_EXPENSE_TYPE"] + "").ToString(), 50);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["REFERENCE_MOC"] + "").ToString(), 4000);

                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_AREA"] + "").ToString(), "N");
                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_APU"] + "").ToString(), "N");
                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_BUSINESS_UNIT"] + "").ToString(), "N");
                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_UNIT_NO"] + "").ToString(), "N");

                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["OTHER_AREA"] + "").ToString(), 4000);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["OTHER_APU"] + "").ToString(), 4000);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["OTHER_BUSINESS_UNIT"] + "").ToString(), 4000);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["OTHER_UNIT_NO"] + "").ToString(), 4000);

                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["FUNCTIONAL_LOCATION"] + "").ToString(), 4000);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["PHA_REQUEST_NAME"] + "").ToString(), 4000);
                    sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["TARGET_START_DATE"] + "").ToString());
                    sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["TARGET_END_DATE"] + "").ToString());
                    sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ACTUAL_START_DATE"] + "").ToString());
                    sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ACTUAL_END_DATE"] + "").ToString());

                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DESCRIPTIONS"] + "").ToString(), 4000);

                    sqlstr += " ,getdate()";
                    sqlstr += " ,null";
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CREATE_BY"] + "").ToString(), 50);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                    sqlstr += ")";
                    #endregion insert
                }
                else if (action_type == "update")
                {
                    #region update

                    sqlstr = "update  EPHA_T_GENERAL set ";

                    sqlstr += " ID_RAM = " + cls.ChkSqlNum((dt.Rows[i]["ID_RAM"] + "").ToString(), "N");
                    sqlstr += " ,EXPENSE_TYPE = " + cls.ChkSqlStr((dt.Rows[i]["EXPENSE_TYPE"] + "").ToString(), 50);
                    sqlstr += " ,SUB_EXPENSE_TYPE = " + cls.ChkSqlStr((dt.Rows[i]["SUB_EXPENSE_TYPE"] + "").ToString(), 50);
                    sqlstr += " ,REFERENCE_MOC = " + cls.ChkSqlStr((dt.Rows[i]["REFERENCE_MOC"] + "").ToString(), 4000);

                    sqlstr += " ,ID_AREA = " + cls.ChkSqlNum((dt.Rows[i]["ID_AREA"] + "").ToString(), "N");
                    sqlstr += " ,ID_APU = " + cls.ChkSqlNum((dt.Rows[i]["ID_APU"] + "").ToString(), "N");
                    sqlstr += " ,ID_BUSINESS_UNIT = " + cls.ChkSqlNum((dt.Rows[i]["ID_BUSINESS_UNIT"] + "").ToString(), "N");
                    sqlstr += " ,ID_UNIT_NO = " + cls.ChkSqlNum((dt.Rows[i]["ID_UNIT_NO"] + "").ToString(), "N");

                    sqlstr += " ,OTHER_AREA = " + cls.ChkSqlStr((dt.Rows[i]["OTHER_AREA"] + "").ToString(), 4000);
                    sqlstr += " ,OTHER_APU = " + cls.ChkSqlStr((dt.Rows[i]["OTHER_APU"] + "").ToString(), 4000);
                    sqlstr += " ,OTHER_BUSINESS_UNIT = " + cls.ChkSqlStr((dt.Rows[i]["OTHER_BUSINESS_UNIT"] + "").ToString(), 4000);
                    sqlstr += " ,OTHER_UNIT_NO = " + cls.ChkSqlStr((dt.Rows[i]["OTHER_UNIT_NO"] + "").ToString(), 4000);

                    sqlstr += " ,FUNCTIONAL_LOCATION = " + cls.ChkSqlStr((dt.Rows[i]["FUNCTIONAL_LOCATION"] + "").ToString(), 4000);
                    sqlstr += " ,PHA_REQUEST_NAME = " + cls.ChkSqlStr((dt.Rows[i]["PHA_REQUEST_NAME"] + "").ToString(), 4000);
                    sqlstr += " ,TARGET_START_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["TARGET_START_DATE"] + "").ToString());
                    sqlstr += " ,TARGET_END_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["TARGET_END_DATE"] + "").ToString());
                    sqlstr += " ,ACTUAL_START_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ACTUAL_START_DATE"] + "").ToString());
                    sqlstr += " ,ACTUAL_END_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ACTUAL_END_DATE"] + "").ToString());

                    sqlstr += " ,DESCRIPTIONS = " + cls.ChkSqlStr((dt.Rows[i]["DESCRIPTIONS"] + "").ToString(), 4000);

                    sqlstr += " ,UPDATE_DATE = getdate()";
                    sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);


                    sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                    sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");

                    #endregion update
                }
                else if (action_type == "delete")
                {
                    #region delete
                    sqlstr = "delete from EPHA_T_GENERAL ";

                    sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                    sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                    sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                    #endregion delete
                }

                if (action_type != "")
                {
                    ret = cls_conn.ExecuteNonQuery(sqlstr);
                    if (ret == "") { ret = "true"; }
                    if (ret != "true") { break; }
                }
            }

            if (ret == "") { ret = "true"; }
            if (ret != "true") { return ret; }
            #endregion update data general

            #region update data functional audition
            if (dsData.Tables["functional_audition"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["functional_audition"].Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    string seq_functional_audition = (dt.Rows[i]["seq"] + "").ToString();

                    if (action_type == "insert")
                    {
                        #region insert
                        //SEQ Auto running
                        sqlstr = "insert into EPHA_T_FUNCTIONAL_AUDITION (" +
                            "SEQ,ID,ID_PHA,FUNCTIONAL_LOCATION" +
                            ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";
                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum(seq_functional_audition, "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_functional_audition, "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");

                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["FUNCTIONAL_LOCATION"] + "").ToString(), 4000);

                        sqlstr += " ,getdate()";
                        sqlstr += " ,null";
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CREATE_BY"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += ")";
                        #endregion insert

                    }
                    else if (action_type == "update")
                    {
                        #region update

                        sqlstr = "update EPHA_T_FUNCTIONAL_AUDITION set ";

                        sqlstr += " FUNCTIONAL_LOCATION = " + cls.ChkSqlStr((dt.Rows[i]["FUNCTIONAL_LOCATION"] + "").ToString(), 4000);

                        sqlstr += " ,UPDATE_DATE = getdate()";
                        sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");

                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "delete from EPHA_T_FUNCTIONAL_AUDITION ";

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ret = cls_conn.ExecuteNonQuery(sqlstr);
                        if (ret != "true") { break; }
                    }
                }

                if (ret == "") { ret = "true"; }
                if (ret != "true") { return ret; }
            }
            #endregion update data functional audition

            #region update data session
            dt = new DataTable();
            dt = dsData.Tables["session"].Copy(); dt.AcceptChanges();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string action_type = (dt.Rows[i]["action_type"] + "").ToString();

                if (action_type == "insert")
                {

                    #region insert
                    //SEQ Auto running
                    sqlstr = "insert into EPHA_T_SESSION (" +
                        "SEQ,ID,ID_PHA,NO,MEETING_DATE,MEETING_START_TIME,MEETING_END_TIME" +
                        ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                        ") values ";
                    sqlstr += " ( ";
                    sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                    sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");

                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");

                    sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["MEETING_DATE"] + "").ToString());
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["MEETING_START_TIME"] + "").ToString(), 100);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["MEETING_END_TIME"] + "").ToString(), 100);

                    sqlstr += " ,getdate()";
                    sqlstr += " ,null";
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CREATE_BY"] + "").ToString(), 50);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                    sqlstr += ")";
                    #endregion insert 
                }
                else if (action_type == "update")
                {
                    #region update

                    sqlstr = "update EPHA_T_SESSION set ";

                    sqlstr += " MEETING_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["MEETING_DATE"] + "").ToString());
                    sqlstr += " ,MEETING_START_TIME = " + cls.ChkSqlStr((dt.Rows[i]["MEETING_START_TIME"] + "").ToString(), 100);
                    sqlstr += " ,MEETING_END_TIME = " + cls.ChkSqlStr((dt.Rows[i]["MEETING_END_TIME"] + "").ToString(), 100);

                    sqlstr += " ,UPDATE_DATE = getdate()";
                    sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                    sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                    sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                    sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");

                    #endregion update
                }
                else if (action_type == "delete")
                {
                    #region delete
                    sqlstr = "delete from EPHA_T_SESSION ";

                    sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                    sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                    sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                    #endregion delete
                }

                if (action_type != "")
                {
                    ret = cls_conn.ExecuteNonQuery(sqlstr);
                    if (ret != "true") { break; }
                }
            }
            if (ret == "") { ret = "true"; }
            if (ret != "true") { return ret; }
            #endregion update data session 

            #region update data memberteam
            if (dsData.Tables["memberteam"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["memberteam"].Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    if (action_type == "insert")
                    {
                        #region insert
                        //SEQ Auto running
                        sqlstr = "insert into EPHA_T_MEMBER_TEAM (" +
                            "SEQ,ID,ID_SESSION,ID_PHA,NO,USER_NAME,USER_DISPLAYNAME" +
                            ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";
                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_SESSION"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["USER_NAME"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["USER_DISPLAYNAME"] + "").ToString(), 4000);

                        sqlstr += " ,getdate()";
                        sqlstr += " ,null";
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CREATE_BY"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += ")";
                        #endregion insert 
                    }
                    else if (action_type == "update")
                    {
                        #region update

                        sqlstr = "update EPHA_T_MEMBER_TEAM set ";

                        sqlstr += " NO = " + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ,USER_NAME = " + cls.ChkSqlStr((dt.Rows[i]["USER_NAME"] + "").ToString(), 50);
                        sqlstr += " ,USER_DISPLAYNAME = " + cls.ChkSqlStr((dt.Rows[i]["USER_DISPLAYNAME"] + "").ToString(), 4000);

                        sqlstr += " ,UPDATE_DATE = getdate()";
                        sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and ID_SESSION = " + cls.ChkSqlNum((dt.Rows[i]["ID_SESSION"] + "").ToString(), "N");

                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "delete from EPHA_T_MEMBER_TEAM ";

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and ID_SESSION = " + cls.ChkSqlNum((dt.Rows[i]["ID_SESSION"] + "").ToString(), "N");
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ret = cls_conn.ExecuteNonQuery(sqlstr);
                        if (ret != "true") { break; }
                    }
                }

                if (ret == "") { ret = "true"; }
                if (ret != "true") { return ret; }
            }
            #endregion update data memberteam

            #region update data drawing
            if (dsData.Tables["drawing"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["drawing"].Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    if (action_type == "insert")
                    {
                        #region insert
                        //SEQ Auto running
                        sqlstr = "insert into EPHA_T_DRAWING (" +
                            "SEQ,ID,ID_PHA,NO,DOCUMENT_NAME,DOCUMENT_NO,DOCUMENT_FILE_NAME,DOCUMENT_FILE_PATH,DESCRIPTIONS" +
                            ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";
                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_NAME"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_NO"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_NAME"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_PATH"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DESCRIPTIONS"] + "").ToString(), 4000);

                        sqlstr += " ,getdate()";
                        sqlstr += " ,null";
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CREATE_BY"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += ")";
                        #endregion insert 
                    }
                    else if (action_type == "update")
                    {
                        #region update

                        sqlstr = "update EPHA_T_DRAWING set ";

                        sqlstr += " NO = " + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ,DOCUMENT_NAME = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_NAME"] + "").ToString(), 4000);
                        sqlstr += " ,DOCUMENT_NO = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_NO"] + "").ToString(), 4000);
                        sqlstr += " ,DOCUMENT_FILE_NAME = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_NAME"] + "").ToString(), 4000);
                        sqlstr += " ,DOCUMENT_FILE_PATH = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_PATH"] + "").ToString(), 4000);
                        sqlstr += " ,DESCRIPTIONS = " + cls.ChkSqlStr((dt.Rows[i]["DESCRIPTIONS"] + "").ToString(), 4000);

                        sqlstr += " ,UPDATE_DATE = getdate()";
                        sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");

                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "delete from EPHA_T_DRAWING ";

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ret = cls_conn.ExecuteNonQuery(sqlstr);
                        if (ret != "true") { break; }
                    }
                }
                if (ret == "") { ret = "true"; }
                if (ret != "true") { return ret; }
            }
            #endregion update data drawing

            return ret;
        }
        public string set_hazop_partii(ref DataSet dsData, ref ClassConnectionDb cls_conn, string seq_header_now)
        {
            string ret = "";

            #region update data node
            if (dsData.Tables["node"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["node"].Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    if (action_type == "insert")
                    {
                        #region insert
                        //SEQ Auto running
                        sqlstr = "insert into EPHA_T_NODE (" +
                            "SEQ,ID,ID_PHA,NO,NODE,DESIGN_INTENT,DESIGN_CONDITIOINS,OPERATING_CONDITIOINS,NODE_BOUNDARY,DESCRIPTIONS" +
                            ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";
                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["NODE"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DESIGN_INTENT"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DESIGN_CONDITIOINS"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["OPERATING_CONDITIOINS"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["NODE_BOUNDARY"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DESCRIPTIONS"] + "").ToString(), 4000);

                        sqlstr += " ,getdate()";
                        sqlstr += " ,null";
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CREATE_BY"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += ")";
                        #endregion insert

                    }
                    else if (action_type == "update")
                    {
                        #region update

                        sqlstr = "update EPHA_T_NODE set ";

                        sqlstr += " NO = " + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ,NODE = " + cls.ChkSqlStr((dt.Rows[i]["NODE"] + "").ToString(), 4000);
                        sqlstr += " ,DESIGN_INTENT = " + cls.ChkSqlStr((dt.Rows[i]["DESIGN_INTENT"] + "").ToString(), 4000);
                        sqlstr += " ,DESIGN_CONDITIOINS = " + cls.ChkSqlStr((dt.Rows[i]["DESIGN_CONDITIOINS"] + "").ToString(), 4000);
                        sqlstr += " ,OPERATING_CONDITIOINS = " + cls.ChkSqlStr((dt.Rows[i]["OPERATING_CONDITIOINS"] + "").ToString(), 4000);
                        sqlstr += " ,NODE_BOUNDARY = " + cls.ChkSqlStr((dt.Rows[i]["NODE_BOUNDARY"] + "").ToString(), 4000);
                        sqlstr += " ,DESCRIPTIONS = " + cls.ChkSqlStr((dt.Rows[i]["DESCRIPTIONS"] + "").ToString(), 4000);

                        sqlstr += " ,UPDATE_DATE = getdate()";
                        sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");

                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "delete from EPHA_T_NODE ";

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ret = cls_conn.ExecuteNonQuery(sqlstr);
                        if (ret != "true") { break; }
                    }
                }
                if (ret == "") { ret = "true"; }
                if (ret != "true") { return ret; }
            }
            #endregion update data node

            #region update data nodedrawing
            if (dsData.Tables["nodedrawing"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["nodedrawing"].Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    if (action_type == "insert")
                    {
                        #region insert
                        //SEQ Auto running
                        sqlstr = "insert into EPHA_T_NODE_DRAWING (" +
                            "SEQ,ID,ID_PHA,ID_NODE,ID_DRAWING,NO,PAGE_START_FIRST,PAGE_END_FIRST,PAGE_START_SECOND,PAGE_END_SECOND,PAGE_START_THIRD,PAGE_END_THIRD,DESCRIPTIONS" +
                            ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";
                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_NODE"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_DRAWING"] + "").ToString(), "N");

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["PAGE_START_FIRST"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["PAGE_END_FIRST"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["PAGE_START_SECOND"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["PAGE_END_SECOND"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["PAGE_START_THIRD"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["PAGE_END_THIRD"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DESCRIPTIONS"] + "").ToString(), 4000);

                        sqlstr += " ,getdate()";
                        sqlstr += " ,null";
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CREATE_BY"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += ")";
                        #endregion insert

                    }
                    else if (action_type == "update")
                    {
                        #region update

                        sqlstr = "update EPHA_T_NODE_DRAWING set ";

                        sqlstr += " NO = " + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ,ID_DRAWING = " + cls.ChkSqlNum((dt.Rows[i]["ID_DRAWING"] + "").ToString(), "N");
                        sqlstr += " ,PAGE_START_FIRST = " + cls.ChkSqlNum((dt.Rows[i]["PAGE_START_FIRST"] + "").ToString(), "N");
                        sqlstr += " ,PAGE_END_FIRST = " + cls.ChkSqlNum((dt.Rows[i]["PAGE_END_FIRST"] + "").ToString(), "N");
                        sqlstr += " ,PAGE_START_SECOND = " + cls.ChkSqlNum((dt.Rows[i]["PAGE_START_SECOND"] + "").ToString(), "N");
                        sqlstr += " ,PAGE_END_SECOND = " + cls.ChkSqlNum((dt.Rows[i]["PAGE_END_SECOND"] + "").ToString(), "N");
                        sqlstr += " ,PAGE_START_THIRD = " + cls.ChkSqlNum((dt.Rows[i]["PAGE_START_THIRD"] + "").ToString(), "N");
                        sqlstr += " ,PAGE_END_THIRD = " + cls.ChkSqlNum((dt.Rows[i]["PAGE_END_THIRD"] + "").ToString(), "N");
                        sqlstr += " ,DESCRIPTIONS = " + cls.ChkSqlStr((dt.Rows[i]["DESCRIPTIONS"] + "").ToString(), 4000);

                        sqlstr += " ,UPDATE_DATE = getdate()";
                        sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and ID_NODE = " + cls.ChkSqlNum((dt.Rows[i]["ID_NODE"] + "").ToString(), "N");

                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "delete from EPHA_T_NODE_DRAWING ";

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and ID_NODE = " + cls.ChkSqlNum((dt.Rows[i]["ID_NODE"] + "").ToString(), "N");
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ret = cls_conn.ExecuteNonQuery(sqlstr);
                        if (ret != "true") { break; }
                    }
                }
                if (ret == "") { ret = "true"; }
                if (ret != "true") { return ret; }
            }
            #endregion update data nodedrawing

            #region update data nodeguidwords
            if (dsData.Tables["nodeguidwords"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["nodeguidwords"].Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    if (action_type == "insert")
                    {
                        #region insert
                        //SEQ Auto running
                        sqlstr = "insert into EPHA_T_NODE_GUIDE_WORDS (" +
                            "SEQ,ID,ID_PHA,ID_NODE,ID_GUIDE_WORD,NO " +
                            ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";
                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_NODE"] + "").ToString(), "N");

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_GUIDE_WORD"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");

                        sqlstr += " ,getdate()";
                        sqlstr += " ,null";
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CREATE_BY"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += ")";
                        #endregion insert

                    }
                    else if (action_type == "update")
                    {
                        #region update

                        sqlstr = "update EPHA_T_NODE_GUIDE_WORDS set ";

                        sqlstr += " NO = " + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ,ID_GUIDE_WORD = " + cls.ChkSqlNum((dt.Rows[i]["ID_GUIDE_WORD"] + "").ToString(), "N");

                        sqlstr += " ,UPDATE_DATE = getdate()";
                        sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and ID_NODE = " + cls.ChkSqlNum((dt.Rows[i]["ID_NODE"] + "").ToString(), "N");

                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "delete from EPHA_T_NODE_GUIDE_WORDS ";

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and ID_NODE = " + cls.ChkSqlNum((dt.Rows[i]["ID_NODE"] + "").ToString(), "N");
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ret = cls_conn.ExecuteNonQuery(sqlstr);
                        if (ret != "true") { break; }
                    }
                }
                if (ret == "") { ret = "true"; }
                if (ret != "true") { return ret; }
            }
            #endregion update data nodeguidwords

            return ret;

        }
        public string set_hazop_partiii(ref DataSet dsData, ref ClassConnectionDb cls_conn, string seq_header_now)
        {
            string ret = "";
            #region update data nodeworksheet
            if (dsData.Tables["nodeworksheet"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["nodeworksheet"].Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    if (action_type == "insert")
                    {
                        #region insert
                        //SEQ Auto running
                        sqlstr = "insert into EPHA_T_NODE_WORKSHEET (" +
                            "SEQ,ID,ID_PHA,ROW_TYPE,ID_NODE,ID_GUIDE_WORD,NO,CAUSES_NO,CAUSES,CONSEQUENCES_NO,CONSEQUENCES" +
                            ",CATEGORY_NO,CATEGORY_TYPE,RAM_BEFOR_SECURITY,RAM_BEFOR_LIKELIHOOD,RAM_BEFOR_RISK,MAJOR_ACCIDENT_EVENT,EXISTING_SAFEGUARDS" +
                            ",RAM_AFTER_SECURITY,RAM_AFTER_LIKELIHOOD,RAM_AFTER_RISK,RECOMMENDATIONS,RESPONDER_USER_NAME,RESPONDER_USER_DISPLAYNAME" +
                            ",ACTION_STATUS" +
                            ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";
                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");

                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["ROW_TYPE"] + "").ToString(), 50);

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_NODE"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_GUIDE_WORD"] + "").ToString(), "N");

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["CAUSES_NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CAUSES"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["CONSEQUENCES_NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CONSEQUENCES"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["CATEGORY_NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CATEGORY_TYPE"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_BEFOR_SECURITY"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_BEFOR_LIKELIHOOD"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_BEFOR_RISK"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["MAJOR_ACCIDENT_EVENT"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["EXISTING_SAFEGUARDS"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_SECURITY"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_LIKELIHOOD"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_RISK"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RECOMMENDATIONS"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_NAME"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_DISPLAYNAME"] + "").ToString(), 4000);

                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["ACTION_STATUS"] + "").ToString(), 50);

                        sqlstr += " ,getdate()";
                        sqlstr += " ,null";
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CREATE_BY"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += ")";
                        #endregion insert

                    }
                    else if (action_type == "update")
                    {
                        #region update

                        sqlstr = "update EPHA_T_NODE_WORKSHEET set ";

                        sqlstr += " NO = " + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ,CAUSES_NO = " + cls.ChkSqlNum((dt.Rows[i]["CAUSES_NO"] + "").ToString(), "N");
                        sqlstr += " ,CAUSES = " + cls.ChkSqlStr((dt.Rows[i]["CAUSES"] + "").ToString(), 4000);
                        sqlstr += " ,CONSEQUENCES_NO = " + cls.ChkSqlNum((dt.Rows[i]["CONSEQUENCES_NO"] + "").ToString(), "N");
                        sqlstr += " ,CONSEQUENCES = " + cls.ChkSqlStr((dt.Rows[i]["CONSEQUENCES"] + "").ToString(), 4000);
                        sqlstr += " ,CATEGORY_NO = " + cls.ChkSqlNum((dt.Rows[i]["CATEGORY_NO"] + "").ToString(), "N");
                        sqlstr += " ,CATEGORY_TYPE = " + cls.ChkSqlStr((dt.Rows[i]["CATEGORY_TYPE"] + "").ToString(), 4000);
                        sqlstr += " ,RAM_BEFOR_SECURITY = " + cls.ChkSqlStr((dt.Rows[i]["RAM_BEFOR_SECURITY"] + "").ToString(), 10);
                        sqlstr += " ,RAM_BEFOR_LIKELIHOOD = " + cls.ChkSqlStr((dt.Rows[i]["RAM_BEFOR_LIKELIHOOD"] + "").ToString(), 10);
                        sqlstr += " ,RAM_BEFOR_RISK = " + cls.ChkSqlStr((dt.Rows[i]["RAM_BEFOR_RISK"] + "").ToString(), 10);
                        sqlstr += " ,MAJOR_ACCIDENT_EVENT = " + cls.ChkSqlStr((dt.Rows[i]["MAJOR_ACCIDENT_EVENT"] + "").ToString(), 10);
                        sqlstr += " ,EXISTING_SAFEGUARDS = " + cls.ChkSqlStr((dt.Rows[i]["EXISTING_SAFEGUARDS"] + "").ToString(), 4000);
                        sqlstr += " ,RAM_AFTER_SECURITY = " + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_SECURITY"] + "").ToString(), 10);
                        sqlstr += " ,RAM_AFTER_LIKELIHOOD = " + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_LIKELIHOOD"] + "").ToString(), 10);
                        sqlstr += " ,RAM_AFTER_RISK = " + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_RISK"] + "").ToString(), 10);
                        sqlstr += " ,RECOMMENDATIONS = " + cls.ChkSqlStr((dt.Rows[i]["RECOMMENDATIONS"] + "").ToString(), 4000);
                        sqlstr += " ,RESPONDER_USER_NAME = " + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_NAME"] + "").ToString(), 50);
                        sqlstr += " ,RESPONDER_USER_DISPLAYNAME = " + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_DISPLAYNAME"] + "").ToString(), 4000);

                        sqlstr += " ,UPDATE_DATE = getdate()";
                        sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and ID_NODE = " + cls.ChkSqlNum((dt.Rows[i]["ID_NODE"] + "").ToString(), "N");
                        sqlstr += " and ID_GUIDE_WORD = " + cls.ChkSqlNum((dt.Rows[i]["ID_GUIDE_WORD"] + "").ToString(), "N");

                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "delete from EPHA_T_NODE_WORKSHEET ";

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and ID_NODE = " + cls.ChkSqlNum((dt.Rows[i]["ID_NODE"] + "").ToString(), "N");
                        sqlstr += " and ID_GUIDE_WORD = " + cls.ChkSqlNum((dt.Rows[i]["ID_GUIDE_WORD"] + "").ToString(), "N");
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ret = cls_conn.ExecuteNonQuery(sqlstr);
                        if (ret != "true") { break; }
                    }
                }
                if (ret == "") { ret = "true"; }
                if (ret != "true") { return ret; }
            }
            #endregion update data nodeworksheet
            return ret;

        }
        public string set_hazop_partiv(ref DataSet dsData, ref ClassConnectionDb cls_conn, string seq_header_now)
        {
            string ret = "";
            #region update data managerecom
            if (dsData.Tables["managerecom"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["managerecom"].Copy(); dt.AcceptChanges();

                #region delete
                sqlstr = "delete from EPHA_T_MANAGE_RECOM ";
                sqlstr += " where ID_PHA = " + cls.ChkSqlNum((dt.Rows[0]["ID_PHA"] + "").ToString(), "N");
                ret = cls_conn.ExecuteNonQuery(sqlstr);
                if (ret != "true") { return ret; }
                #endregion delete

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    if (action_type == "insert")
                    {
                        #region insert
                        //SEQ Auto running
                        sqlstr = "insert into EPHA_T_MANAGE_RECOM (" +
                            "SEQ,ID,ID_PHA,RESPONDER_USER_NAME" +
                            ",ESTIMATED_START_DATE,ESTIMATED_END_DATE" +
                            ",DOCUMENT_FILE_NAME,DOCUMENT_FILE_PATH,ACTION_STATUS,RESPONDER_ACTION_TYPE" +
                            ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";
                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_NAME"] + "").ToString(), 50);

                        sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_START_DATE"] + "").ToString());
                        sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_END_DATE"] + "").ToString());
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_NAME"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_PATH"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["ACTION_STATUS"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["RESPONDER_ACTION_TYPE"] + "").ToString(), "N");//0,1

                        sqlstr += " ,getdate()";
                        sqlstr += " ,null";
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CREATE_BY"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += ")";
                        #endregion insert

                    }
                    else if (action_type == "update")
                    {
                        #region update

                        sqlstr = "update EPHA_T_MANAGE_RECOM set ";

                        sqlstr += " ESTIMATED_START_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_START_DATE"] + "").ToString());
                        sqlstr += " ,ESTIMATED_END_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_END_DATE"] + "").ToString());
                        sqlstr += " ,DOCUMENT_FILE_NAME = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_NAME"] + "").ToString(), 4000);
                        sqlstr += " ,DOCUMENT_FILE_PATH = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_PATH"] + "").ToString(), 4000);
                        sqlstr += " ,ACTION_STATUS = " + cls.ChkSqlStr((dt.Rows[i]["ACTION_STATUS"] + "").ToString(), 50);
                        sqlstr += " ,RESPONDER_ACTION_TYPE = " + cls.ChkSqlNum((dt.Rows[i]["RESPONDER_ACTION_TYPE"] + "").ToString(), "N");//0,1

                        sqlstr += " ,UPDATE_DATE = getdate()";
                        sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and RESPONDER_USER_NAME = " + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_NAME"] + "").ToString(), 50);

                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        //sqlstr = "delete from EPHA_T_MANAGE_RECOM ";

                        //sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        //sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        //sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        //sqlstr += " and RESPONDER_USER_NAME = " + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_NAME"] + "").ToString(), 50);
                        #endregion delete
                    }

                    if (action_type != "")
                    {
                        ret = cls_conn.ExecuteNonQuery(sqlstr);
                        if (ret != "true") { break; }
                    }
                }
                if (ret == "") { ret = "true"; }
                if (ret != "true") { return ret; }
            }
            #endregion update data managerecom
            return ret;

        }


        public string set_follow_up(SetDocHazopModel param)
        {
            string msg = "";
            string ret = "";
            cls_json = new ClassJSON();

            DataSet dsData = new DataSet();
            string user_name = (param.user_name + "");
            string flow_action = param.flow_action;
            string sqlstr_check = "";

            //$scope.flow_role_type = "admin";//admin,request,responder,approver
            string role_type = ("admin");
            Boolean bOwnerAction = true;//เป็น Owner Action ด้วยหรือป่าว


            jsper = param.json_managerecom + "";
            try
            {
                if (jsper.Trim() != "")
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "managerecom";
                        dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                        ret = "";
                    }
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }
            if (ret.ToLower() == "error") { goto Next_Line_Convert; }

            if (true)
            {
                #region connection transaction
                cls = new ClassFunctions();

                cls_conn = new ClassConnectionDb();
                cls_conn.OpenConnection();
                cls_conn.BeginTransaction();
                #endregion connection transaction
                try
                {
                    dt = new DataTable();
                    dt = dsData.Tables["managerecom"].Copy(); dt.AcceptChanges();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        #region update 
                        sqlstr = "update EPHA_T_MANAGE_RECOM set ";
                        sqlstr += " DOCUMENT_FILE_NAME = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_NAME"] + "").ToString(), 4000);
                        sqlstr += " ,DOCUMENT_FILE_PATH = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_PATH"] + "").ToString(), 4000);
                        sqlstr += " ,ACTION_STATUS = " + cls.ChkSqlStr((dt.Rows[i]["ACTION_STATUS"] + "").ToString(), 50);
                        sqlstr += " ,RESPONDER_ACTION_TYPE = 1";//0,1,2-> 2 = ห้ามแก้ไข

                        sqlstr += " ,UPDATE_DATE = getdate()";
                        sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr(user_name, 50);

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and RESPONDER_USER_NAME = " + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_NAME"] + "").ToString(), 50);
                        #endregion update

                        ret = cls_conn.ExecuteNonQuery(sqlstr);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }


                    }

                }
                catch (Exception ex) { ret = ex.Message.ToString(); goto Next_Line; }

            Next_Line:;

                #region connection transaction
                if (ret == "") { ret = "true"; }
                if (ret == "true")
                {
                    cls_conn.CommitTransaction();
                }
                else
                {
                    cls_conn.RollbackTransaction();
                }
                cls_conn.CloseConnection();
                #endregion connection transaction


                #region  flow action  submit  
                if (ret == "true" && dt.Rows.Count > 0)
                {
                    sqlstr_check = "";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        //ไว้สำหรับส่ง mail แจ้งเตือน admin ข้อมูลต้องเรียงตาม id_pha, responder_user_name
                        string seq = (dt.Rows[i]["ID_PHA"] + "").ToString();
                        string responder_user_name = (dt.Rows[i]["RESPONDER_USER_NAME"] + "").ToString();
                        ClassEmail clsmail = new ClassEmail();
                        clsmail.MailNotificationToAdminOwnerUpdateAction(seq, responder_user_name, "hazop");


                        //ไว้สำหรับส่ง mail แจ้งเตือน admin ข้อมูลต้องเรียงตาม id_pha, responder_user_name
                        //ตอนนี้จะมีแค่ รายการเดียว เท่านั้นก่อน 
                        if (sqlstr_check != "") { sqlstr_check += " or "; }
                        sqlstr_check += " t.id_pha = " + cls.ChkSqlNum(seq, "N");
                    }

                    #region check pha no - Action Owner update action items closed all  -> คนสุดท้ายของใบงาน
                    if (true)
                    {
                        ClassHazop classHazop = new ClassHazop();
                        sqlstr = classHazop.QueryFollowUpDetail("", "", "", "hazop", false);

                        sqlstr = "select t.id_pha,sum(t.status_open) as status_open from (" + sqlstr + ")t where (" + sqlstr_check + ") group by t.id_pha";
                        sqlstr = "select distinct t2.id_pha from (" + sqlstr + ")t2 where t2.status_open = 0 ";

                        cls_conn = new ClassConnectionDb();
                        DataTable dtaction = new DataTable();
                        dtaction = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
                        if (dtaction.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtaction.Rows.Count; i++)
                            {
                                string seq = (dtaction.Rows[i]["ID_PHA"] + "").ToString();
                                //mail not admin กรณีที่ Action Owner Update status Closed All 
                                ClassEmail clsmail = new ClassEmail();
                                clsmail.MailToAdminReviewAll(seq, "hazop");



                                #region update pha status 
                                string pha_status_new = "14";

                                cls = new ClassFunctions();
                                cls_conn = new ClassConnectionDb();
                                cls_conn.OpenConnection();
                                cls_conn.BeginTransaction();

                                #region update responder_action_type ให้เป็น responder_action_type = 2 ห้ามแก้ไข
                                sqlstr = "update EPHA_T_MANAGE_RECOM set responder_action_type = 2 where id_pha = " + cls.ChkSqlNum(seq, "N");
                                ret = cls_conn.ExecuteNonQuery(sqlstr);
                                #endregion update responder_action_type ให้เป็น responder_action_type = 2 ห้ามแก้ไข

                                #region update
                                sqlstr = "update EPHA_F_HEADER set ";
                                sqlstr += " PHA_STATUS = " + cls.ChkSqlNum((pha_status_new).ToString(), "N");
                                sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                                #endregion update

                                ret = cls_conn.ExecuteNonQuery(sqlstr);
                                if (ret == "") { ret = "true"; }
                                if (ret == "true")
                                {
                                    cls_conn.CommitTransaction();
                                }
                                else
                                {
                                    cls_conn.RollbackTransaction();
                                }
                                cls_conn.CloseConnection();

                                #endregion update pha status 

                            }
                        }
                    }
                    #endregion check pha no - Action Owner update action items closed all   -> คนสุดท้ายของใบงาน

                }
                #endregion  flow action  submit 

            }

        Next_Line_Convert:;
            return cls_json.SetJSONresult(refMsg(ret, msg));
        }

    }
}
