using dotnet6_epha_api.Class;
using Model;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Xml.Linq;

namespace Class
{

    public class ClassHazop
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
        private string get_pha_no(string sub_software, string year)
        {
            //hazop format : HAZOP-2013-1000002
            DataTable _dt = new DataTable();
            cls = new ClassFunctions();

            sqlstr = @" select '" + sub_software.ToUpper() + "-" + year.ToUpper() + "-' + right('0000000' + trim(str(coalesce(max(replace(upper(pha_no),'" + sub_software.ToUpper() + "-" + year.ToUpper() + "-','')+1),0))),7) as pha_no ";
            sqlstr += @" from EPHA_F_HEADER where lower(pha_sub_software) = lower('" + sub_software + "') and year = " + year;

            cls_conn = new ClassConnectionDb();
            _dt = new DataTable();
            _dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            return (_dt.Rows[0]["pha_no"].ToString() + "");
        }
        private void get_history_search_follow(ref DataSet _dsData, string id_pha, string user_name_active)
        {
            sqlstr = @" select distinct isnull(h.approver_user_name,'') as id,  isnull(h.approver_user_displayname,'') as name
                         from EPHA_F_HEADER h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha 
                         inner join EPHA_T_MANAGE_RECOM w on h.id = w.id_pha and  lower(nw.responder_user_name) =  lower(w.responder_user_name) 
                         where h.approver_user_displayname is not null ";
            if (id_pha != "") { sqlstr += @" and lower(h.seq) = lower(" + cls.ChkSqlStr(id_pha, 50) + ")  "; }
            if (user_name_active != "") { sqlstr += @" and lower(nw.responder_user_name) = lower(" + cls.ChkSqlStr(user_name_active, 50) + ")  "; }
            sqlstr += " order by isnull(h.approver_user_displayname,'')";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            dt.TableName = "his_approver";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

        }
        private void get_history_hazop(ref DataSet _dsData)
        {
            sqlstr = @" select * from(select distinct b.pha_request_name  as name
            from EPHA_F_HEADER a inner join EPHA_T_GENERAL b on a.id = b.id_pha 
            where a.pha_status not in (81) )t order by t.name  ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            dt.TableName = "his_request_name";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();


            sqlstr = @" select * from(select distinct b.descriptions  as name
            from EPHA_F_HEADER a inner join EPHA_T_GENERAL b on a.id = b.id_pha 
            where a.pha_status not in (81) )t order by t.name  ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            dt.TableName = "his_descriptions";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

        }
        private void get_master_hazop(ref DataSet _dsData)
        {
            sqlstr = @" select seq,seq as id,user_id as employee_id, user_name as employee_name, user_displayname as employee_displayname, user_email as employee_email
                        , 'assets/img/team/avatar.webp' as employee_img, user_type as employee_type
                        , 0 as selected_type
                         from VW_EPHA_PERSON_DETAILS t 
                         order by user_name";
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            dt.TableName = "employee";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();


            #region master storagelocation  
            string sqlstr_def = @"  select a.id as id_company, a.name as name_company 
                         , b.id as id_area, b.name as name_area
                         , c.id as id_apu, c.name as name_apu
                         , d.id as id_bussiness_unit, d.name as name_bussiness_unit
                         , d.id as id_unit_no, d.name as name_unit_no
                         from EPHA_M_COMPANY a
                         left join EPHA_M_AREA b on a.id = b.id_company 
                         left join EPHA_M_APU c on a.id = c.id_company and b.id = c.id_area
                         left join EPHA_M_BUSSINESS_UNIT d on a.id = d.id_company and b.id = d.id_area and c.id = d.id_apu
                         left join EPHA_M_UNIT_NO e on a.id = e.id_company and b.id = e.id_area and c.id = e.id_apu and d.id = e.id_bussiness_unit
                         ";

            sqlstr = @"  select * from (" + sqlstr_def + ")t order by id_company,name_area,name_apu,name_bussiness_unit,name_unit_no ";
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            dt.Rows.Add(dt.NewRow()); dt.AcceptChanges();

            dt.TableName = "storagelocation";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            sqlstr = @"  select id_company as id, name_company as name from (" + sqlstr_def + ")t order by id_company ";
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
            dt.TableName = "company";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            sqlstr = @"  select id_company, id_area as id, name_area as name from (" + sqlstr_def + ")t order by id_company, id_area ";
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
            dt.TableName = "area";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            sqlstr = @"  select id_company, id_area, id_apu as id, name_apu as name from (" + sqlstr_def + ")t order by id_company, id_area, id_apu  ";
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
            dt.TableName = "apu_map";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            sqlstr = @"  select id_company, id_area, id_apu, id_bussiness_unit as id, name_bussiness_unit as name from (" + sqlstr_def + ")t order by id_company, id_area, id_apu, id_bussiness_unit ";
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
            dt.TableName = "bussiness_unit_map";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            sqlstr = @"  select * from (" + sqlstr_def + ")t order by id_company,name_area,name_apu,name_bussiness_unit,name_unit_no ";
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
            dt.TableName = "unit_no_map";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion storagelocation  

            #region master guidwords  
            sqlstr = @" select seq, parameter, deviations, guide_words, guide_words as guidewords, process_deviation, area_application, 0 as selected_type, 0 as main_parameter, def_selected
                        from EPHA_M_GUIDE_WORDS where active_type = 1 order by  parameter, deviations, guide_words, process_deviation, area_application ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            //sort data 
            string befor_parameter = "";
            string after_parameter = "";
            int irow = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                befor_parameter = (dt.Rows[i]["parameter"] + "").ToString();
                if (befor_parameter != after_parameter)
                {
                    after_parameter = befor_parameter;
                    dt.Rows[i]["main_parameter"] = 1;
                    dt.AcceptChanges();
                }
            }
            if (befor_parameter != after_parameter)
            {
                after_parameter = befor_parameter;
                dt.Rows[dt.Rows.Count - 1]["main_parameter"] = 1;
                dt.AcceptChanges();
            }
            dt.TableName = "guidwords";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion guidwords  

            #region master functional location
            sqlstr = @" select *, a.functional_location as id, a.functional_location as name, 0 as selected_type
                         from EPHA_M_FUNCTIONAL_LOCATION a
                         where active_type = 1 order by seq ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            dt.TableName = "functional";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion functional location  

            #region master ram
            sqlstr = @" select seq, id, name, 0 as selected_type, category_type
                        from EPHA_M_RAM where active_type = 1 order by seq ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            dt.TableName = "ram";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            sqlstr = @"   select s.*, s.security_level, s.name as security_text, s.name as security_show
                          , s.people as people_text, s.assets as assets_text, s.enhancement as enhancement_text, s.reputation as reputation_text, s.product_quality as product_quality_text
                          , 0 as selected_type ,a.category_type
                          from  EPHA_M_RAM a
                          inner join EPHA_M_RAM_SECURITY s on a.id = s.id_ram  
                          order by s.id_ram, s.sort_by";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            dt.TableName = "security_level";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            sqlstr = @"  select  b.*, 0 as selected_type ,a.category_type
                         , '' as security_text
                         , '' as people_text, '' as assets_text, '' as enhancement_text, '' as reputation_text, '' as product_quality_text
                         from  EPHA_M_RAM a
                         inner join EPHA_M_RAM_LEVEL b on a.id = b.id_ram 
                         order by a.id , b.sort_by ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string security_level = (dt.Rows[i]["security_level"] + "");
                DataRow[] dr = dsData.Tables["security_level"].Select("security_level = " + security_level);
                if (dr.Length > 0)
                {
                    dt.Rows[i]["security_text"] = (dr[0]["name"] + "");
                    dt.Rows[i]["people_text"] = (dr[0]["people"] + "");
                    dt.Rows[i]["assets_text"] = (dr[0]["assets"] + "");
                    dt.Rows[i]["enhancement_text"] = (dr[0]["enhancement"] + "");
                    dt.Rows[i]["reputation_text"] = (dr[0]["reputation"] + "");
                    dt.Rows[i]["product_quality_text"] = (dr[0]["product_quality"] + "");

                    //string col_show = "likelihood_show";
                    //try { dt.Columns.Add(col_show); dt.AcceptChanges(); } catch { }
                    //try { dt.Rows[i][col_show] = (dt.Rows[i]["likelihood" + security_level + "_text"] + ""); } catch { }

                    dt.AcceptChanges();
                }
            }
            dt.TableName = "ram_level";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            if (dt.Rows.Count > 0)
            {
                DataTable dtNew = new DataTable();
                dtNew.Columns.Add("id_ram", typeof(int));
                dtNew.Columns.Add("selected_type", typeof(int));
                dtNew.Columns.Add("category_type", typeof(int));
                dtNew.Columns.Add("likelihood_level");
                dtNew.Columns.Add("likelihood_show");
                dtNew.Columns.Add("likelihood_text");
                dtNew.Columns.Add("likelihood_desc");
                dtNew.Columns.Add("likelihood_criterion");
                dtNew.AcceptChanges();

                dt = new DataTable();
                dt = dsData.Tables["ram"].Copy(); dt.AcceptChanges();

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        int id_ram = Convert.ToInt32(dt.Rows[i]["id"]);
                        int category_type = Convert.ToInt32(dt.Rows[i]["category_type"]);
                        int iNo = (i + 1);

                        DataRow[] dr = (dsData.Tables["ram_level"]).Select("id_ram=" + id_ram);
                        if (dr.Length > 0)
                        {
                            for (int rl = 0; rl < dr.Length; rl++)
                            {
                                for (int j = 1; j < 8; j++)
                                {
                                    if ((dr[rl]["likelihood" + j + "_level"] + "") == "") { break; }
                                    int iNewRow = dtNew.Rows.Count;
                                    dtNew.Rows.Add(dtNew.NewRow()); dtNew.AcceptChanges();
                                    dtNew.Rows[iNewRow]["id_ram"] = id_ram;
                                    dtNew.Rows[iNewRow]["selected_type"] = 0;
                                    dtNew.Rows[iNewRow]["category_type"] = category_type;
                                    try
                                    {
                                        dtNew.Rows[iNewRow]["likelihood_level"] = (dr[rl]["likelihood" + j + "_level"] + ""); 
                                        dtNew.Rows[iNewRow]["likelihood_show"] = (dr[rl]["likelihood" + j + "_text"] + "");
                                        if (category_type == 1)
                                        {
                                            dtNew.Rows[iNewRow]["likelihood_text"] = (dr[rl]["likelihood" + j + "_text"] + "");
                                            dtNew.Rows[iNewRow]["likelihood_desc"] = (dr[rl]["likelihood" + j + "_desc"] + "");
                                            dtNew.Rows[iNewRow]["likelihood_criterion"] = (dr[rl]["likelihood" + j + "_criterion"] + "");
                                        }
                                    }
                                    catch { }

                                    dtNew.AcceptChanges();
                                    if (category_type == 0 && j == 3) { break; }
                                }
                                break;
                            }
                        }
                    }
                }
                dtNew.TableName = "likelihood_level";
                _dsData.Tables.Add(dtNew.Copy()); dsData.AcceptChanges();


            }

            #endregion ram

            #region master apu
            sqlstr = @" select id_company, id_area, id, name from EPHA_M_APU t order by id_company, id_area, id  ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            dt.TableName = "apu";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion apu

            #region master business unit
            sqlstr = @" select id_company, id_area, id_apu, id, name from EPHA_M_BUSSINESS_UNIT t order by id_company, id_area, id_apu, id  ";
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            dt.TableName = "business_unit";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion business unit

            #region master unit no
            sqlstr = @" select id_company, id_area, id_apu, id_bussiness_unit, id, name from EPHA_M_UNIT_NO t order by id_company, id_area, id_apu, id_bussiness_unit, id    ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            dt.TableName = "unit_no";
            _dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion unit no
        }
        private void set_max_id(ref DataTable dtmax, string name, string values)
        {
            if (dtmax.Rows.Count == 0)
            {
                dtmax = new DataTable();
                dtmax.Columns.Add("name");
                dtmax.Columns.Add("values");
                dtmax.AcceptChanges();
            }

            int irow = dtmax.Rows.Count;
            dtmax.Rows.Add(dtmax.NewRow());
            dtmax.Rows[irow]["name"] = name;
            dtmax.Rows[irow]["values"] = values;
            dtmax.AcceptChanges();
        }
        public string get_hazop_details(LoadDocModel param)
        {
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();
            string token_doc = (param.token_doc + "").Trim();
            string sub_software = (param.sub_software + "").Trim();
            string seq = token_doc;

            get_master_hazop(ref dsData);
            get_history_hazop(ref dsData);

            DataHazop(ref dsData, user_name, seq, sub_software);

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;


        }

        public static string SetJSONresult(DataTable _dtJson)
        {
            string JSONresult;
            JSONresult = JsonConvert.SerializeObject(_dtJson);
            return JSONresult;
        }

        public string get_hazop_search(LoadDocModel param)
        {
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();
            string token_doc = (param.token_doc + "").Trim();
            string sub_software = (param.sub_software + "").Trim();
            string type_doc = (param.type_doc + "").Trim();
            string seq = token_doc;

            get_master_hazop(ref dsData);
            DataHazopSearch(ref dsData, user_name, seq, sub_software);


            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;
        }
        public string get_hazop_followup(LoadDocModel param)
        {
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();
            string token_doc = (param.token_doc + "").Trim();
            string sub_software = (param.sub_software + "").Trim();
            string seq = token_doc;

            get_master_hazop(ref dsData);
            get_history_search_follow(ref dsData, seq, user_name);

            DataHazopSearchFollowUp(ref dsData, user_name, seq, sub_software);

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public string get_hazop_followup_detail(LoadDocFollowModel param)
        {
            dsData = new DataSet();
            string user_name = (param.user_name + "").Trim();
            string token_doc = (param.token_doc + "").Trim();
            string sub_software = (param.sub_software + "").Trim();
            string pha_no = (param.pha_no + "").Trim();
            string responder_user_name = (param.responder_user_name + "").Trim();
            string seq = token_doc;

            DataHazopSearchFollowUpDetail(ref dsData, user_name, seq, pha_no, responder_user_name, sub_software);

            string json = JsonConvert.SerializeObject(dsData, Formatting.Indented);

            return json;

        }
        public void DataHazop(ref DataSet dsData, string user_name, string seq, string sub_software)
        {
            DataTable dtma = new DataTable();
            string pha_no = "";
            int id_pha = 0;
            int id_node = 0;
            int id_nodeworksheet = 0;
            int id_managerecom = 0;

            string year_now = System.DateTime.Now.Year.ToString();
            if (Convert.ToInt64(year_now) > 2500) { year_now = (Convert.ToInt64(year_now) - 543).ToString(); }


            dt = new DataTable();
            cls = new ClassFunctions();

            //--user_name, user_displayname, user_email
            sqlstr = @" select *  from EPHA_PERSON_DETAILS a where 1=1 ";
            sqlstr += " and lower(a.user_name) = lower(coalesce(" + cls.ChkSqlStr(user_name, 50) + ",'x'))  ";
            cls_conn = new ClassConnectionDb();
            DataTable dtemp = new DataTable();
            dtemp = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            #region header
            sqlstr = @" select a.*,b.name as pha_status_name, b.descriptions as pha_status_displayname
                        ,case when a.year = year(getdate()) then vw.user_name else a.request_user_name end request_user_name
                        ,case when a.year = year(getdate()) then vw.user_name else a.request_user_displayname end request_user_displayname
                        ,null as approver_user_img
                        , 'update' as action_type, 0 as action_change
                        from EPHA_F_HEADER a
                        left join EPHA_M_STATUS b on a.pha_status = b.id
                        left join VW_EPHA_PERSON_DETAILS vw on lower(a.pha_request_by) = lower(vw.user_name)
                        where 1=1";

            sqlstr += " and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count == 0)
            {
                pha_no = get_pha_no(sub_software, year_now);
                id_pha = (get_max("EPHA_F_HEADER"));
                set_max_id(ref dtma, "header", id_pha.ToString());


                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_pha;
                dt.Rows[0]["id"] = id_pha;
                dt.Rows[0]["year"] = year_now;
                dt.Rows[0]["pha_no"] = pha_no;
                dt.Rows[0]["pha_version"] = 0;
                dt.Rows[0]["pha_status"] = 11;
                dt.Rows[0]["pha_sub_software"] = sub_software;
                dt.Rows[0]["request_approver"] = 0;

                dt.Rows[0]["pha_status_name"] = "DF";
                dt.Rows[0]["pha_status_displayname"] = "Draft";
                if (dtemp.Rows.Count > 0)
                {
                    dt.Rows[0]["pha_request_by"] = (dtemp.Rows[0]["user_name"] + "");
                    dt.Rows[0]["request_user_name"] = (dtemp.Rows[0]["user_name"] + "");
                    dt.Rows[0]["request_user_displayname"] = (dtemp.Rows[0]["user_displayname"] + "");
                }
                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }

            pha_no = (dt.Rows[0]["pha_no"] + "");
            id_pha = Convert.ToInt32(dt.Rows[0]["id"] + "");


            dt.TableName = "header";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion header

            #region general 
            sqlstr = @" select b.* , '' as functional_location_audition, '' as business_unit_name, '' as unit_no_name, 'update' as action_type, 0 as action_change
                        from EPHA_F_HEADER a inner join EPHA_T_GENERAL b on a.id  = b.id_pha
                        where 1=1 ";
            sqlstr += " and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";
            sqlstr += " order by a.seq,b.seq";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_pha;
                dt.Rows[0]["id"] = id_pha;// (get_max("EPHA_T_GENERAL")); ข้อมูล 1 ต่อ 1 ให้ใช้กับ header ได้เลย
                dt.Rows[0]["id_pha"] = id_pha;

                dt.Rows[0]["functional_location_audition"] = "TPX-76-LICSA-001-TX,TPX-76-LICSA-002-TX,TPX-76-LICSA-003-TX";

                //default values 
                DataTable dtram = dsData.Tables["ram"].Copy(); dtram.AcceptChanges();
                dt.Rows[0]["id_ram"] = dtram.Rows[0]["id"];

                dt.Rows[0]["expense_type"] = "OPEX";
                dt.Rows[0]["sub_expense_type"] = "Normal";

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }
            dt.AcceptChanges();

            dt.TableName = "general";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion general


            #region functional_audition 
            sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                        from EPHA_F_HEADER a inner join EPHA_T_FUNCTIONAL_AUDITION b on a.id  = b.id_pha
                        where 1=1 ";
            sqlstr += " and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";
            sqlstr += " order by a.seq,b.seq";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            int id_functional_audition = (get_max("EPHA_T_FUNCTIONAL_AUDITION"));
            set_max_id(ref dtma, "functional_audition", id_functional_audition.ToString());

            if (dt.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_functional_audition;
                dt.Rows[0]["id"] = id_functional_audition;
                dt.Rows[0]["id_pha"] = id_pha;
                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }

            dt.TableName = "functional_audition";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion functional_audition

            #region session 
            sqlstr = @" select b.* , 0 as no, 'update' as action_type, 0 as action_change
                        from EPHA_F_HEADER a inner join EPHA_T_SESSION b on a.id  = b.id_pha
                        where 1=1 ";
            sqlstr += " and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";
            sqlstr += " order by a.seq,b.seq";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            int id_session = (get_max("EPHA_T_SESSION"));
            set_max_id(ref dtma, "session", id_session.ToString());

            if (dt.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_session;
                dt.Rows[0]["id"] = id_session;
                dt.Rows[0]["id_pha"] = id_pha;

                dt.Rows[0]["no"] = 1;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }
            id_session = Convert.ToInt32(dt.Rows[0]["id"]);

            dt.TableName = "session";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion session


            #region memberteam 
            sqlstr = @" select c.* , 'assets/img/team/avatar.webp' as user_img, 'update' as action_type, 0 as action_change
                        from EPHA_F_HEADER a 
                        inner join EPHA_T_SESSION b  on a.id  = b.id_pha 
                        inner join EPHA_T_MEMBER_TEAM c on a.id  = c.id_pha and b.id  = c.id_session";
            sqlstr += " and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";
            sqlstr += " order by a.seq,b.seq,c.seq";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            int id_memberteam = (get_max("EPHA_T_MEMBER_TEAM"));
            set_max_id(ref dtma, "memberteam", id_memberteam.ToString());

            if (dt.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_memberteam;
                dt.Rows[0]["id"] = id_memberteam;
                dt.Rows[0]["id_pha"] = id_pha;
                dt.Rows[0]["id_session"] = id_session;
                dt.Rows[0]["no"] = 1;

                dt.Rows[0]["no"] = 1;
                dt.Rows[0]["user_img"] = "assets/img/team/avatar.webp";

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }

            dt.TableName = "memberteam";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion memberteam

            #region drawing 
            sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                        from EPHA_F_HEADER a inner join EPHA_T_DRAWING b on a.id  = b.id_pha
                        where 1=1 ";
            sqlstr += " and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";
            sqlstr += " order by a.seq,b.seq";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            int id_drawing = (get_max("EPHA_T_DRAWING"));
            set_max_id(ref dtma, "drawing", id_drawing.ToString());

            if (dt.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_drawing;
                dt.Rows[0]["id"] = id_drawing;
                dt.Rows[0]["id_pha"] = id_pha;

                dt.Rows[0]["no"] = 1;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }
            dt.TableName = "drawing";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion drawing


            #region node 
            sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                        from EPHA_F_HEADER a inner join EPHA_T_NODE b on a.id  = b.id_pha
                        where 1=1 ";
            sqlstr += " and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";
            sqlstr += " order by a.seq,b.seq";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count == 0)
            {
                id_node = (get_max("EPHA_T_NODE"));

                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_node;
                dt.Rows[0]["id"] = id_node;
                dt.Rows[0]["id_pha"] = id_pha;

                dt.Rows[0]["no"] = 1;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }
            id_node = Convert.ToInt32(dt.Rows[0]["id"] + "");
            set_max_id(ref dtma, "node", id_node.ToString());

            dt.TableName = "node";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion node


            #region nodedrawing 
            sqlstr = @" select b.* , 'update' as action_type, 0 as action_change
                        , b.id_node as seq_node
                        from EPHA_F_HEADER a inner join EPHA_T_NODE_DRAWING b on a.id  = b.id_pha
                        where 1=1 ";
            sqlstr += " and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";
            sqlstr += " order by a.seq,b.seq";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            int id_nodedrawing = 0;
            if (dt.Rows.Count == 0)
            {
                id_nodedrawing = (get_max("EPHA_T_NODE_DRAWING"));

                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_nodedrawing;
                dt.Rows[0]["id"] = id_nodedrawing;
                dt.Rows[0]["id_node"] = id_node;
                dt.Rows[0]["id_pha"] = id_pha;

                dt.Rows[0]["seq_node"] = id_node;

                dt.Rows[0]["no"] = 1;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }
            id_nodedrawing = Convert.ToInt32(dt.Rows[0]["id"] + "");
            set_max_id(ref dtma, "nodedrawing", id_nodedrawing.ToString());

            dt.TableName = "nodedrawing";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion nodedrawing

            #region nodeguidwords 
            sqlstr = @"  select b.* ,coalesce(def_selected,0) as selected_type , 'update' as action_type, 0 as action_change
                        , b.id_node as seq_node, g.guide_words as guidewords, g.deviations, 0 as no_guide_word
                        from EPHA_F_HEADER a inner join EPHA_T_NODE_GUIDE_WORDS b on a.id  = b.id_pha
                        left join EPHA_M_GUIDE_WORDS g on b.id_guide_word = g.id
                        where 1=1 ";
            sqlstr += " and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";
            sqlstr += " order by a.seq,b.seq";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            int id_nodeguidwords = (get_max("EPHA_T_NODE_GUIDE_WORDS"));
            set_max_id(ref dtma, "nodeguidwords", id_nodeguidwords.ToString());

            if (dt.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_nodeguidwords;
                dt.Rows[0]["id"] = id_nodeguidwords;
                dt.Rows[0]["id_node"] = id_node;
                dt.Rows[0]["id_pha"] = id_pha;

                dt.Rows[0]["seq_node"] = id_node;
                dt.Rows[0]["no"] = 1;

                ////หาหน้าบ้าน
                //dt.Rows[0]["id_guide_words"] = id_node;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }

            DataTable dtnodeguidwords = new DataTable();
            dtnodeguidwords = dt.Copy(); dtnodeguidwords.AcceptChanges();

            dt.TableName = "nodeguidwords";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion nodeguidwords

            #region nodeworksheet 
            sqlstr = @" select b.* , 0 as no  
                        , 'update' as action_type, 0 as action_change
                        , b.id_node as seq_node, g.guide_words as guidewords, g.deviations
                        , vw.user_id as responder_user_id, vw.user_email as responder_user_email
                        , 'assets/img/team/avatar.webp' as responder_user_img
                        from EPHA_F_HEADER a 
                        inner join EPHA_T_NODE_WORKSHEET b on a.id  = b.id_pha
                        inner join EPHA_M_GUIDE_WORDS g on b.id_guide_word = g.id    
                        left join VW_EPHA_PERSON_DETAILS vw on lower(b.responder_user_name) = lower(vw.user_name) 
                        where 1=1";
            sqlstr += " and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";
            sqlstr += " order by b.no, b.causes_no, b.consequences_no, b.category_no";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            id_nodeworksheet = (get_max("EPHA_T_NODE_WORKSHEET"));
            set_max_id(ref dtma, "nodeworksheet", id_nodeworksheet.ToString());

            if (dt.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่ เดียวให้หน้าบ้านเช็คแล้ว loop เอา -> logic เดียวต้องรวมกับ functions add อยู่แล้ว
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_nodeworksheet;
                dt.Rows[0]["id"] = id_nodeworksheet;
                dt.Rows[0]["id_node"] = id_node;
                dt.Rows[0]["id_pha"] = id_pha;

                dt.Rows[0]["seq_node"] = id_node;

                dt.Rows[0]["no"] = 1;

                dt.Rows[0]["row_type"] = "causes";//guideword,causes,consequences,cat

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "new";
                dt.Rows[0]["action_change"] = 0;
                dt.Rows[0]["action_status"] = "Open";
                dt.AcceptChanges();
            }

            dt.TableName = "nodeworksheet";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion nodeworksheet


            #region managerecom 
            sqlstr = @" select b.* , 0 as no, 'update' as action_type, 0 as action_change
                        , vw.user_id as responder_user_id, 'assets/img/team/avatar.webp' as responder_user_img
                        from EPHA_F_HEADER a inner join EPHA_T_MANAGE_RECOM b on a.id  = b.id_pha
                        left join VW_EPHA_PERSON_DETAILS vw on lower(b.responder_user_name) = lower(vw.user_name)
                        where 1=1 ";
            sqlstr += " and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";
            sqlstr += " order by a.seq,b.seq";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            id_managerecom = (get_max("EPHA_T_MANAGE_RECOM"));
            set_max_id(ref dtma, "managerecom", id_nodeworksheet.ToString());

            if (dt.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่ เดียวให้หน้าบ้านเช็คแล้ว loop เอา -> logic เดียวต้องรวมกับ functions add อยู่แล้ว
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_managerecom;
                dt.Rows[0]["id"] = id_managerecom;
                dt.Rows[0]["id_pha"] = id_pha;

                dt.Rows[0]["no"] = 1;

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }

            dt.TableName = "managerecom";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion managerecom

            dtma.TableName = "max";
            dsData.Tables.Add(dtma.Copy()); dsData.AcceptChanges();

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

        }

        public void DataHazopSearch(ref DataSet dsData, string user_name, string seq, string sub_software)
        {
            DataTable dtma = new DataTable();
            string pha_no = "";
            int id_pha = 0;
            user_name = (user_name == "" ? "zkuluwat" : user_name);

            string year_now = System.DateTime.Now.Year.ToString();
            if (Convert.ToInt64(year_now) > 2500) { year_now = (Convert.ToInt64(year_now) - 543).ToString(); }

            dt = new DataTable();
            cls = new ClassFunctions();

            //--user_name, user_displayname, user_email
            sqlstr = @" select *  from EPHA_PERSON_DETAILS a where 1=1 ";
            sqlstr += " and lower(a.user_name) = lower(coalesce(" + cls.ChkSqlStr(user_name, 50) + ",'x'))  ";
            cls_conn = new ClassConnectionDb();
            DataTable dtemp = new DataTable();
            dtemp = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            string sqlstr_def = "";
            #region header 
            sqlstr = @" select a.*,b.*,ms.name as pha_status_name, ms.descriptions as pha_status_displayname
                        ,case when a.year = year(getdate()) then vw.user_name else a.request_user_name end request_user_name
                        ,case when a.year = year(getdate()) then vw.user_name else a.request_user_displayname end request_user_displayname
                        ,null as approver_user_img
                        , 'update' as action_type, 0 as action_change
                        from EPHA_F_HEADER a
						inner join EPHA_T_GENERAL b on a.id  = b.id_pha
                        left join EPHA_M_STATUS ms on a.pha_status = ms.id
                        left join VW_EPHA_PERSON_DETAILS vw on lower(a.pha_request_by) = lower(vw.user_name)
                        where 1=1";
            if (seq != "") { sqlstr += " and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  "; }
            sqlstr += " order by a.pha_no";

            sqlstr_def = sqlstr;

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count == 0)
            {
                pha_no = get_pha_no(sub_software, year_now);
                id_pha = (get_max("EPHA_F_HEADER"));
                set_max_id(ref dtma, "header", id_pha.ToString());

                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_pha;
                dt.Rows[0]["id"] = id_pha;
                dt.Rows[0]["year"] = year_now;
                dt.Rows[0]["pha_no"] = pha_no;
                dt.Rows[0]["pha_version"] = 0;
                dt.Rows[0]["pha_status"] = 11;
                dt.Rows[0]["pha_sub_software"] = sub_software;
                dt.Rows[0]["request_approver"] = 0;

                dt.Rows[0]["pha_status_name"] = "DF";
                dt.Rows[0]["pha_status_displayname"] = "Draft";

                dt.Rows[0]["pha_request_by"] = (dtemp.Rows[0]["user_name"] + "");
                dt.Rows[0]["request_user_name"] = (dtemp.Rows[0]["user_name"] + "");
                dt.Rows[0]["request_user_displayname"] = (dtemp.Rows[0]["user_displayname"] + "");

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }

            pha_no = (dt.Rows[0]["pha_no"] + "");
            id_pha = Convert.ToInt32(dt.Rows[0]["id"] + "");


            dt.TableName = "header";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion header

            #region general 
            sqlstr = @" select b.* , '' as functional_location_audition, '' as business_unit_name, '' as unit_no_name, 'update' as action_type, 0 as action_change
                        from EPHA_F_HEADER a inner join EPHA_T_GENERAL b on a.id  = b.id_pha
                        where 1=2 ";
            sqlstr += " order by a.pha_no";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_pha;
                dt.Rows[0]["id"] = id_pha;// (get_max("EPHA_T_GENERAL")); ข้อมูล 1 ต่อ 1 ให้ใช้กับ header ได้เลย
                dt.Rows[0]["id_pha"] = id_pha;

                dt.Rows[0]["functional_location_audition"] = "";

                //default values 
                DataTable dtram = dsData.Tables["ram"].Copy(); dtram.AcceptChanges();
                dt.Rows[0]["id_ram"] = dtram.Rows[0]["id"];

                dt.Rows[0]["expense_type"] = "OPEX";
                dt.Rows[0]["sub_expense_type"] = "Normal";

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }
            dt.AcceptChanges();

            dt.TableName = "general";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion general


            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

        }

        private void check_role_user_active(string user_name, ref string role_type)
        {

            ClassLogin classLogin = new ClassLogin();
            sqlstr = classLogin.QueryAdminUser_Role(user_name);
            cls_conn = new ClassConnectionDb();
            DataTable dtrole = new DataTable();
            dtrole = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
            if (dtrole.Rows.Count > 0)
            {
                for (int i = 0; i < dtrole.Rows.Count; i++)
                {
                    role_type = (dtrole.Rows[0]["role_type"] + "").ToString();
                    if (role_type == "admin") { break; }
                }
            }
        }
        public void DataHazopSearchFollowUp(ref DataSet dsData, string user_name, string seq, string sub_software)
        {
            DataTable dtma = new DataTable();
            string pha_no = "";
            int id_pha = 0;
            user_name = (user_name == "" ? "zkuluwat" : user_name);

            string role_type = "";
            check_role_user_active(user_name, ref role_type);

            string year_now = System.DateTime.Now.Year.ToString();
            if (Convert.ToInt64(year_now) > 2500) { year_now = (Convert.ToInt64(year_now) - 543).ToString(); }

            dt = new DataTable();
            cls = new ClassFunctions();

            //--user_name, user_displayname, user_email
            sqlstr = @" select *  from EPHA_PERSON_DETAILS a where 1=1 ";
            sqlstr += " and lower(a.user_name) = lower(coalesce(" + cls.ChkSqlStr(user_name, 50) + ",'x'))  ";
            cls_conn = new ClassConnectionDb();
            DataTable dtemp = new DataTable();
            dtemp = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            #region header
            string sqlstr_w = "";
            string sqlstr_r = "";
            sqlstr_w = @" select 0 as no, a.pha_sub_software, a.seq as pha_seq,a.pha_no, g.pha_request_name, '' as responder_user_displayname 
                         ,count(1) as status_total
                         , count(case when lower(w.action_status) = 'closed' then null else 1 end) status_open
                         , count(case when lower(w.action_status) = 'closed' then 1 else null end) status_closed
                         , 'worksheet' as data_by, '' as responder_user_name
                         , case when a.pha_status  = 13 then 'Waiting Follow Up' else 'Waiting Review Follow Up' end as pha_status_name
                         , 'update' as action_type, 0 as action_change
                         from EPHA_F_HEADER a 
                         inner join EPHA_T_GENERAL g on a.id = g.id_pha 
						 inner join EPHA_T_NODE_WORKSHEET nw on a.id = nw.id_pha 
                         inner join EPHA_T_MANAGE_RECOM w on a.id = w.id_pha and  lower(nw.responder_user_name) =  lower(w.responder_user_name) 
                         where a.pha_status in (13,14) and nw.responder_user_name is not null";
            if (seq != "") { sqlstr_w += @" and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  "; }

            if (role_type != "admin") { sqlstr_w += @" and ( a.pha_status in (13,14) and isnull(nw.responder_action_type,0) <> 2 )"; }

            sqlstr_w += @" group by a.pha_status, a.pha_sub_software, a.seq, a.pha_no, g.pha_request_name ";

            sqlstr_r = @" select  0 as no, a.pha_sub_software, '' as pha_seq, '' as pha_no, '' as pha_request_name, vw.user_displayname as responder_user_displayname
                         ,count(1) as status_total
                         , count(case when lower(w.action_status) = 'closed' then null else 1 end) status_open
                         , count(case when lower(w.action_status) = 'closed' then 1 else null end) status_closed
                         , 'responder' as data_by, w.responder_user_name
                         , case when a.pha_status  = 13 then 'Waiting Follow Up' else 'Waiting Review Follow Up' end as pha_status_name
                         , 'update' as action_type, 0 as action_change
                         from EPHA_F_HEADER a 
                         inner join EPHA_T_GENERAL g on a.id = g.id_pha 
						 inner join EPHA_T_NODE_WORKSHEET nw on a.id = nw.id_pha 
                         inner join EPHA_T_MANAGE_RECOM w on a.id = w.id_pha and  lower(nw.responder_user_name) =  lower(w.responder_user_name) 
                         inner join VW_EPHA_PERSON_DETAILS vw on lower(w.responder_user_name) = lower(vw.user_name) 
                         where a.pha_status in (13,14) and nw.responder_user_name is not null";
            if (seq != "") { sqlstr_r += @" and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  "; }
            if (role_type != "admin") { sqlstr_w += @" and ( a.pha_status in (13,14) and isnull(nw.responder_action_type,0) <> 2 )"; }

            sqlstr_r += @" group by a.pha_status, a.pha_sub_software, vw.user_displayname, w.responder_user_name ";

            sqlstr = " select t.* from (" + sqlstr_w + " union " + sqlstr_r + ")t order by data_by, pha_sub_software, pha_no, pha_request_name, responder_user_displayname  ";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["pha_sub_software"] = pha_no;

                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;

                dt.AcceptChanges();
            }

            dt.TableName = "header";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion header


            #region general 
            sqlstr = @" select b.* , '' as functional_location_audition, '' as business_unit_name, '' as unit_no_name, 'update' as action_type, 0 as action_change
                        from EPHA_F_HEADER a inner join EPHA_T_GENERAL b on a.id  = b.id_pha
                        where 1=2 ";
            sqlstr += " order by a.pha_no";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["seq"] = id_pha;
                dt.Rows[0]["id"] = id_pha;// (get_max("EPHA_T_GENERAL")); ข้อมูล 1 ต่อ 1 ให้ใช้กับ header ได้เลย
                dt.Rows[0]["id_pha"] = id_pha;

                dt.Rows[0]["functional_location_audition"] = "";

                //default values 
                DataTable dtram = dsData.Tables["ram"].Copy(); dtram.AcceptChanges();
                dt.Rows[0]["id_ram"] = dtram.Rows[0]["id"];

                dt.Rows[0]["expense_type"] = "OPEX";
                dt.Rows[0]["sub_expense_type"] = "Normal";

                dt.Rows[0]["create_by"] = user_name;
                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;
                dt.AcceptChanges();
            }
            dt.AcceptChanges();

            dt.TableName = "general";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion general


            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

        }

        public string QueryFollowUpDetail(string seq, string pha_no, string responder_user_name, string sub_software, Boolean bOrderBy)
        {

            sqlstr = @"  select  'update' as action_type, 0 as action_change
                         , 0 as no,a.id as id_pha, upper(a.pha_sub_software) as pha_sub_software, a.pha_no, g.pha_request_name, vw.user_displayname as responder_user_displayname, nw.responder_user_name
                         , (w.action_status) as action_status
						 , count(1) as status_total
                         , count(case when lower(w.action_status) = 'closed' then null else 1 end) status_open
                         , count(case when lower(w.action_status) = 'closed' then 1 else null end) status_closed
						 , w.document_file_name, w.document_file_path, 0 as document_file_size
						 , format(w.estimated_start_date,'dd MMM yyyy') as estimated_start_date_text 
						 , format(w.estimated_end_date,'dd MMM yyyy') as estimated_end_date_text 
						 , isnull(datediff(day,case when w.estimated_end_date > getdate() then getdate() else w.estimated_end_date end,getdate()),0) as over_due
                         , '' as recommendations, '' as causes
                         , w.seq, w.id, isnull(w.responder_action_type,'') as responder_action_type
                         from EPHA_F_HEADER a 
                         inner join EPHA_T_GENERAL g on a.id = g.id_pha 
						 inner join EPHA_T_NODE_WORKSHEET nw on a.id = nw.id_pha 
                         inner join EPHA_T_MANAGE_RECOM w on a.id = w.id_pha and  lower(nw.responder_user_name) =  lower(w.responder_user_name) 
                         inner join VW_EPHA_PERSON_DETAILS vw on lower(w.responder_user_name) = lower(vw.user_name) 
                         where nw.responder_user_name is not null ";

            //ใช้จริงค่อยเปืด
            //sqlstr += " and a.pha_status in (14,15)"; 

            if (seq != "") { sqlstr += @" and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  "; }
            if (pha_no != "") { sqlstr += @" and lower(a.pha_no) = lower(" + cls.ChkSqlStr(pha_no, 50) + ")  "; }
            if (responder_user_name != "") { sqlstr += @" and lower(w.responder_user_name) = lower(" + cls.ChkSqlStr(responder_user_name, 50) + ")  "; }

            sqlstr += @" group by a.id, w.seq, w.id, a.pha_sub_software, a.pha_no, g.pha_request_name, vw.user_displayname, nw.responder_user_name
                         , document_file_name, document_file_path, w.estimated_start_date, w.estimated_end_date, w.action_status, isnull(w.responder_action_type,'') ";

            if (bOrderBy) { sqlstr += @" order by convert(int, a.id), a.pha_sub_software, a.pha_no, g.pha_request_name, vw.user_displayname, nw.responder_user_name"; }


            return sqlstr;
        }
        public void DataHazopSearchFollowUpDetail(ref DataSet dsData, string user_name, string seq, string pha_no, string responder_user_name, string sub_software)
        {
            user_name = (user_name == "" ? "zkuluwat" : user_name);

            Boolean bNewRow = false;
            dt = new DataTable();
            cls = new ClassFunctions();

            #region header 
            sqlstr = QueryFollowUpDetail(seq, pha_no, responder_user_name, sub_software, true);
            // a.pha_status in (13,14) and nw.responder_user_name is not null

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count == 0)
            {
                //กรณีที่เป็นใบงานใหม่
                dt.Rows.Add(dt.NewRow());
                dt.Rows[0]["pha_sub_software"] = pha_no;

                dt.Rows[0]["action_type"] = "insert";
                dt.Rows[0]["action_change"] = 0;

                dt.AcceptChanges();
                bNewRow = true;
            }

            if (bNewRow == false)
            {
                #region worksheet 
                sqlstr = @" select a.pha_no, lower(w.responder_user_name) as responder_user_name, w.causes_no, isnull(w.recommendations,'') as recommendations
                        from EPHA_F_HEADER a 
                        inner join EPHA_T_NODE_WORKSHEET w on a.id  = w.id_pha  
                        inner join VW_EPHA_PERSON_DETAILS vw on lower(w.responder_user_name) = lower(vw.user_name) 
                        where w.responder_user_name is not null  ";
                if (seq != "") { sqlstr += @" and lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  "; }
                if (pha_no != "") { sqlstr += @" and lower(a.pha_no) = lower(" + cls.ChkSqlStr(pha_no, 50) + ")  "; }
                if (responder_user_name != "") { sqlstr += @" and lower(w.responder_user_name) = lower(" + cls.ChkSqlStr(responder_user_name, 50) + ")  "; }
                sqlstr += " order by a.pha_no, w.causes_no, w.recommendations ";

                cls_conn = new ClassConnectionDb();
                DataTable dtworksheet = new DataTable();
                dtworksheet = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

                #endregion group recommendations, causes
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    int ino = 1;
                    string responder_user_name_def = (dt.Rows[i]["responder_user_name"] + "");

                    DataRow[] arr_recomm = dtworksheet.Select("responder_user_name='" + responder_user_name_def + "'");

                    string recomm_text = "";
                    string cause_text = "";

                    for (int r = 0; r < arr_recomm.Length; r++)
                    {
                        if ((arr_recomm[r]["recommendations"] + "") != "")
                        {
                            if (recomm_text != "") { recomm_text += '\n'; }
                            recomm_text += ino + "." + arr_recomm[r]["recommendations"];
                            ino += 1;
                        }
                        if ((arr_recomm[r]["causes_no"] + "") != "")
                        {
                            if (cause_text != "") { cause_text += '\n'; }
                            cause_text += arr_recomm[r]["causes_no"];
                        }
                    }
                    dt.Rows[i]["recommendations"] = recomm_text;
                    dt.Rows[i]["causes"] = cause_text;


                    ///get document_file_size จาก document_file_path
                    try
                    {
                        dt.Rows[i]["document_file_size"] = file_size((dt.Rows[i]["document_file_path"] + ""));
                    }
                    catch { }

                    dt.AcceptChanges();
                }
            }
            dt.TableName = "details";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            #endregion header

            dsData.DataSetName = "dsData"; dsData.AcceptChanges();

        }


        private string file_size(string filePath)
        {
            //string filePath = @"C:\path\to\your\file.txt"; // Replace with your file's path
            FileInfo fileInfo = new FileInfo(filePath);

            if (fileInfo.Exists)
            {
                long fileSizeInBytes = fileInfo.Length;
                Console.WriteLine($"File size: {fileSizeInBytes} bytes");

                // You can convert bytes to other units for better readability
                double fileSizeInKB = fileSizeInBytes / 1024.0;
                //Console.WriteLine($"File size: {fileSizeInKB:F2} KB");
                return ($"({fileSizeInKB:F2} KB)");
                //double fileSizeInMB = fileSizeInKB / 1024.0;
                //Console.WriteLine($"File size: {fileSizeInMB:F2} MB");
            }
            else
            {
                return "";
            }
        }


    }
}
