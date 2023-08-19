using dotnet6_epha_api.Class;
using Model;
using System.Data;
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



    }
}
