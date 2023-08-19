using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;

namespace Class
{
    public class ClassConnectionDb
    {
        string[] sMonth = ("JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC").Split(',');
        string[] sMonths = ("Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec").Split(',');

        String ConnStrSQL = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ConnectionConfig")["ConnString"];

        public DataSet ExecuteAdapterSQL(string SqlStatement)
        {
            DataSet dssql = new DataSet();
            SqlConnection connsql = new SqlConnection(ConnStrSQL);
            connsql.Open();
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(SqlStatement, connsql);
                dssql = new DataSet(); 
                da.Fill(dssql);
            }
            catch { }
            connsql.Close();
            return dssql;

        }
        public DataSet ExecuteAdapterSQL(string ConnStr, string SqlStatement)
        {
            DataSet dssql = new DataSet();
            SqlConnection connsql = new SqlConnection(ConnStr);
            connsql.Open();
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(SqlStatement, connsql);
                dssql = new DataSet();
                da.Fill(dssql);
            }
            catch { }
            connsql.Close();
            return dssql;
        }
        SqlTransaction trans;
        SqlConnection conn;
        SqlCommand comm;
        public string ExecuteNonQuery(string SqlStatement)
        {
            string ret = "";
            try
            {
                comm = new SqlCommand(SqlStatement, conn);
                if (trans != null)
                {
                    comm.Connection = conn;
                    comm.Transaction = trans;
                }

                comm.ExecuteNonQuery();

                ret = "true";
            }
            catch (Exception ex)
            {
                ret = ex.ToString() + " sqlstr :" + SqlStatement;
            }
            return ret;
        }

        public void OpenConnection()
        {
            if (conn == null)
            {
                conn = new SqlConnection(ConnStrSQL);
            }

            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
        }
        public void CloseConnection()
        {
            if (conn.State != ConnectionState.Closed)
            {
                conn.Close();
                conn.Dispose();
            }
        }
        public void BeginTransaction()
        {
            if (trans == null)
            {
                trans = conn.BeginTransaction();
            }
        }
        public void CommitTransaction()
        {
            if (trans != null)
            {
                trans.Commit();
            }
        }
        public void RollbackTransaction()
        {
            if (trans != null)
            {
                trans.Rollback();
            }
        }



    }

}
