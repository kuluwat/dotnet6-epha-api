using System.Data;

namespace dotnet6_epha_api.Class
{
    public class ClassFunctions
    {
        public DataTable refMsg(string status, string remark)
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


        #region function
        public string ChkSqlNum(Object Str, String nType)
        {
            string xNum = "NULL";
            if (Str == null || Convert.IsDBNull(Str))
            {
                return "NULL";
            }
            else if (Str == "NULL")
            {
                return "NULL";
            }
            else
            {
                try
                {
                    if (nType == "N")
                    {
                        xNum = Convert.ToString(Convert.ToInt64(Str.ToString()));
                    }
                    else if (nType == "D")
                    {
                        xNum = Convert.ToString(Convert.ToDouble(Str.ToString()));
                    }
                }
                catch
                {
                    return "NULL";
                }
            }
            return xNum;
        }

        public string ChkSqlNum(Object Str, String nType, int iLength)
        {
            string xNum = "NULL";
            if (Str == null || Convert.IsDBNull(Str))
            {
                return "NULL";
            }
            else if (Str == "NULL")
            {
                return "NULL";
            }
            else
            {
                try
                {
                    if (nType == "N")
                    {
                        xNum = Convert.ToString(Convert.ToInt64(Str.ToString()));
                    }
                    else if (nType == "D")
                    {
                        if (iLength == 0)
                        {
                            xNum = Convert.ToString(Convert.ToDouble(Str.ToString()).ToString("##0"));
                        }
                        else if (iLength == 1)
                        {
                            xNum = Convert.ToString(Convert.ToDouble(Str.ToString()).ToString("##0.0"));
                        }
                        else if (iLength == 2)
                        {
                            xNum = Convert.ToString(Convert.ToDouble(Str.ToString()).ToString("##0.00"));
                        }
                        else if (iLength == 3)
                        {
                            xNum = Convert.ToString(Convert.ToDouble(Str.ToString()).ToString("##0.000"));
                        }
                        else if (iLength == 4)
                        {
                            xNum = Convert.ToString(Convert.ToDouble(Str.ToString()).ToString("##0.0000"));
                        }
                        else
                        {
                            xNum = Convert.ToString(Convert.ToDouble(Str.ToString()));
                        }
                    }
                }
                catch
                {
                    return "NULL";
                }
            }
            return xNum;
        }
        public string ChkSqlStr(object Str, int Length)
        {
            //วิธีที่ 1 --> แทนที่ ' ด้วย ช่องว่าง 1 ช่อง --> " " ทำให้ ' ใน base หายไป
            //วิธีที่ 2 --> แทนที่ ' ด้วย ''         --> Chr(39) & Chr(39) ทำให้ ' ใน base ยังอยู่ 

            //Str = "เลี้ยงตอบแทน' บ.Cyberouis, XX'XX'xxx'xxx"

            string Str1;

            if (Str == null || Convert.IsDBNull(Str))
            {
                return "null";
            }

            if (Str.ToString().ToLower() == "null")
            {
                return "null";
            }

            if (Str.ToString().Trim() == "")
            {
                return "null";
            }

            Str1 = Str.ToString();

            //วิธีที่ 1
            //Str1 = Replace(Str1, Chr(39), " ")

            //วิธีที่ 2
            //Str1 = Replace(Str1, Chr(39), Chr(39) & Chr(39))
            Str1 = Str1.Replace("'", "''");

            if (Str1.ToString().Length >= Length)
            {
                return "'" + Str1.ToString().Substring(0, Length) + "'";
            }
            else
            {
                return "'" + Str1.ToString().Trim() + "'";
            }
        } 

        public string ChkSqlDateYYYYMMDD(Object sDate)
        {
            //20191123 or 2019-11-23 
            try
            {
                int dd;
                int mm;
                int yyyy;

                if (Convert.IsDBNull(sDate))
                {
                    return "NULL";
                }
                else if (sDate.ToString().Replace(" ", "") == "")
                {
                    return "NULL";
                }
                else
                {
                    //sDate = sDate.ToString().Replace("-", "");
                    if (sDate.ToString().IndexOf("-") > -1)
                    {
                        string[] xsDate = sDate.ToString().Split('-');
                        if (xsDate.Length > 2)
                        {
                            sDate = xsDate[0].ToString();
                            if (xsDate[1].ToString().Length == 1) { sDate += "0"; }
                            sDate += xsDate[1].ToString();
                            if (xsDate[2].ToString().Length == 1) { sDate += "0"; }
                            sDate += xsDate[2].ToString();
                        }
                    }

                    String xDate = "";
                    DateTime tsDate = new DateTime();
                    System.Globalization.CultureInfo TmpConvert = new System.Globalization.CultureInfo("en-US");

                    // กรณีที่เป็น Date จาหหน้าจอ ให้เป็น MM/dd/yyyy 
                    try
                    {
                        tsDate = new DateTime(Convert.ToInt16(sDate.ToString().Substring(0, 4)), Convert.ToInt16(sDate.ToString().Substring(4, 2)), Convert.ToInt16(sDate.ToString().Substring(6, 2)));
                    }
                    catch
                    {

                    }
                    if (tsDate.Year > 2500) { tsDate = tsDate.AddYears(-543); }
                    if (tsDate.Year < 2000) { tsDate = tsDate.AddYears(543); }
                    xDate = "CONVERT(date,'" + tsDate.ToString("yyyyMMdd", TmpConvert) + "')";  //CONVERT(date, '20230727')

                    return xDate;
                }
            }
            catch// ทดสอบถ้า debug ได้ให้เอาออก
            {

                return "NULL";
            }
        }



        #endregion function


    }
}
