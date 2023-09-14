using dotnet6_epha_api.Class;
using iTextSharp.text;
using Model;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.DirectoryServices;

namespace Class
{

    public class ClassLogin
    {
        string sqlstr = "";
        ClassFunctions cls = new ClassFunctions();
        ClassJSON cls_json = new ClassJSON();
        ClassConnectionDb cls_conn = new ClassConnectionDb();

        public Boolean LogonAD(ref String msg, string user_name, string user_password)
        {
            String domain = "ThaioilNT";
            String str = String.Format("LDAP://{0}", domain);
            String str2 = String.Format(("{0}\\" + user_name.ToString()), domain);//เท่ากับ domainUser
            String pass = user_password.ToString();
            DirectoryEntry Entry = new DirectoryEntry(str, str2, pass);
            DirectorySearcher Searcher = new DirectorySearcher(Entry);
            SearchResultCollection results;
            try
            {
                results = Searcher.FindAll();

                return true;
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                return false;
            }

        }

        public string QueryAdminUser_Role(string user_name)
        {
            sqlstr = @" select a.user_name, a.user_id, a.user_email, a.user_displayname
                        ,lower(coalesce(c.name,'employee')) as role_type
                        ,'images/user-avatar.png' as user_img
                        from EPHA_PERSON_DETAILS a  
                        inner join EPHA_M_ROLE_SETTING b on lower(a.user_name) = lower(b.user_name) and b.active_type = 1 
                        inner join EPHA_M_ROLE_TYPE c on lower(c.id) = lower(b.id_role_group) and c.active_type = 1   
                        where a.active_type = 1 ";
            if (user_name != "")
            {
                sqlstr += " and lower(a.user_name)  = lower(" + cls.ChkSqlStr(user_name, 50) + ")  ";
            }
            sqlstr += " order by a.user_name,c.name";

            return sqlstr;
        }
        public string login(LoginUserModel param)
        {
            string user_name = (param.user_name + "").Trim();
            try
            {
                if (user_name.IndexOf("@") > -1)
                {
                    string[] x = user_name.Split('@');
                    if (x.Length > 1)
                    {
                        user_name = x[0];
                    }
                }
            }
            catch { }


            DataTable dt = new DataTable();
            cls = new ClassFunctions();

            sqlstr = QueryAdminUser_Role(user_name);
            cls_conn = new ClassConnectionDb();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count == 0)
            {
                //กรณีที่เป็น Employee ทั่วไปเข้าใช้งานระบบ  
                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                dt = cls_conn.ExecuteAdapterSQL(sqlstr.Replace("inner join", "left join")).Tables[0];
            }
            else if (user_name.ToLower() == "admin" || user_name.IndexOf("zNitinaip") > -1)
            {
                dt.Rows[0]["role_type"] = "admin";
                dt.Rows[0]["user_name"] = "admin";
                dt.Rows[0]["user_id"] = "00000000";
                dt.Rows[0]["user_email"] = "admin-epha@thaioilgroup.com";
                dt.Rows[0]["user_display"] = user_name + "(Admin)";
                dt.Rows[0]["user_img"] = "images/user-avatar.png";
                dt.AcceptChanges();
            }

            return cls_json.SetJSONresult(dt);
        }


        private int get_max_seq(string table_name)
        {
            DataTable _dt = new DataTable();
            cls = new ClassFunctions();

            sqlstr = @" select coalesce(max(seq),0)+1 as seq from " + table_name;

            cls_conn = new ClassConnectionDb();
            _dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            return Convert.ToInt32(_dt.Rows[0]["seq"].ToString() + "");
        }
        public string register_account(RegisterAccountModel param)
        {
            string user_displayname = (param.user_displayname + "").Trim();
            string user_email = (param.user_email + "").Trim();
            string user_password = (param.user_password + "").Trim();
            string user_password_confirm = (param.user_password_confirm + "").Trim();
            string user_name = "";

            string ret = ""; string msg = "";

            DataTable dt = new DataTable();
            cls = new ClassFunctions();

            sqlstr = @" select a.user_name, a.user_id, a.user_email, a.user_displayname
                        ,lower(coalesce(c.name,'employee')) as role_type
                        from EPHA_PERSON_DETAILS a  
                        inner join EPHA_M_ROLE_SETTING b on lower(a.user_name) = lower(b.user_name) and b.active_type = 1 
                        inner join EPHA_M_ROLE_TYPE c on lower(c.id) = lower(b.id_role_group) and c.active_type = 1   
                        where a.active_type = 1 ";
            if (user_email != "")
            {
                sqlstr += " and lower(a.user_email)  = lower(" + cls.ChkSqlStr(user_email, 100) + ")  ";
            }
            sqlstr += " order by a.user_name,c.name";
            cls_conn = new ClassConnectionDb();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count > 0)
            {
                ret = "true";
                msg = "User already has data in the system.";
            }
            else
            {
                #region insert/update  
                int seq = get_max_seq("PHA_REGISTER_ACCOUNT");

                if (true)
                {
                    #region insert  
                    sqlstr = "insert into PHA_REGISTER_ACCOUNT(SEQ,USER_DISPLAYNAME,USER_EMAIL,USER_PASSWORD,USER_PASSWORD_CONFIRM,ACCEPT_STATUS" +
                        ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                        ") values ";
                    sqlstr += " ( ";
                    sqlstr += " " + cls.ChkSqlNum(seq.ToString(), "N");
                    sqlstr += " ," + cls.ChkSqlStr(user_displayname, 4000);
                    sqlstr += " ," + cls.ChkSqlStr(user_email, 100);
                    sqlstr += " ," + cls.ChkSqlStr(user_password, 50);
                    sqlstr += " ," + cls.ChkSqlStr(user_password_confirm, 50);
                    sqlstr += " ,null";
                    sqlstr += " ,getdate()";
                    sqlstr += " ,null";
                    sqlstr += " ," + cls.ChkSqlStr("system", 50);
                    sqlstr += " ,null";
                    sqlstr += ")";
                    #endregion insert    
                    cls_conn = new ClassConnectionDb();
                    cls_conn.OpenConnection();
                    ret = cls_conn.ExecuteNonQuery(sqlstr);
                    cls_conn.CloseConnection();
                }

                #endregion insert/update  
                if (ret.ToLower() == "true")
                {
                    ret = "true";
                    msg = "User registration is complete. Please wait for the login credentials from the system administrator.";
                }
                else
                {
                    ret = "error";
                    msg = ret;
                }
            }

            if (ret.ToLower() == "true")
            {
                // email แจ้ง admin ให้ accept การ register 
                ClassEmail clsmail = new ClassEmail();
                clsmail.MailToAdminRegisterAccount(user_displayname, user_email, user_password, user_password_confirm);
            }



            dt = new DataTable();
            dt.Columns.Add("status");
            dt.Columns.Add("msg");
            dt.AcceptChanges();

            dt.Rows.Add(dt.NewRow()); dt.AcceptChanges();
            dt.Rows[0]["status"] = ret;
            dt.Rows[0]["msg"] = msg;

            return cls_json.SetJSONresult(dt);
        }

        public string update_register_account(RegisterAccountModel param)
        {
            string user_active = (param.user_active + "").Trim();
            string user_email = (param.user_email + "").Trim();
            string accept_status = (param.accept_status + "").Trim();

            string ret = ""; string msg = "";

            DataTable dt = new DataTable();
            cls = new ClassFunctions();

            sqlstr = @" select a.user_name, a.user_id, a.user_email, a.user_displayname
                        ,lower(coalesce(c.name,'employee')) as role_type
                        from EPHA_PERSON_DETAILS a  
                        inner join EPHA_M_ROLE_SETTING b on lower(a.user_name) = lower(b.user_name) and b.active_type = 1 
                        inner join EPHA_M_ROLE_TYPE c on lower(c.id) = lower(b.id_role_group) and c.active_type = 1   
                        where a.active_type = 1 ";
            if (user_email != "")
            {
                sqlstr += " and lower(a.user_email)  = lower(" + cls.ChkSqlStr(user_email, 100) + ")  ";
            }
            sqlstr += " order by a.user_name,c.name";
            cls_conn = new ClassConnectionDb();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];


            #region insert/update  
            int seq = get_max_seq("PHA_REGISTER_ACCOUNT");

            if (true)
            {
                #region update  
                sqlstr = "update PHA_REGISTER_ACCOUNT set ";
                sqlstr += " ACCEPT_STATUS = " + cls.ChkSqlNum(accept_status, "N");
                sqlstr += " ,UPDATE_DATE = getdate()";
                sqlstr += " ,UPDATE_BY =" + cls.ChkSqlStr(user_active, 400);
                sqlstr += " where USER_EMAIL = " + cls.ChkSqlStr(user_email, 100);
                #endregion update    

                cls_conn = new ClassConnectionDb();
                cls_conn.OpenConnection();
                ret = cls_conn.ExecuteNonQuery(sqlstr);
                cls_conn.CloseConnection();
            }

            #endregion insert/update    
            if (ret.ToLower() == "true")
            {
                ret = "true";
                msg = "User registration is complete. Please wait for the login credentials from the system administrator.";
            }
            else
            {
                ret = "error";
                msg = ret;
            }

            if (ret.ToLower() == "true")
            {
                if (dt.Rows.Count > 0)
                {
                    string user_displayname = (dt.Rows[0]["user_displayname"] + "").Trim();
                    string user_password = (dt.Rows[0]["user_password"] + "").Trim();
                    string user_password_confirm = (dt.Rows[0]["user_password_confirm"] + "").Trim();

                    // email แจ้ง admin ให้ accept การ register 
                    ClassEmail clsmail = new ClassEmail();
                    clsmail.MailToUserRegisterAccount(user_displayname, user_email, user_password, user_password_confirm, accept_status);
                }
            }



            dt = new DataTable();
            dt.Columns.Add("status");
            dt.Columns.Add("msg");
            dt.AcceptChanges();

            dt.Rows.Add(dt.NewRow()); dt.AcceptChanges();
            dt.Rows[0]["status"] = ret;
            dt.Rows[0]["msg"] = msg;

            return cls_json.SetJSONresult(dt);
        }


    }
}
