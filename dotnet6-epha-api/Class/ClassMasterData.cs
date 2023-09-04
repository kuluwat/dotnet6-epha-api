using dotnet6_epha_api.Class;
using Model;
using System.Data;

namespace Class
{
    
    public class ClassMasterData
    {
        DataTable dt, dtcopy, dtcheck;

        string sqlstr, sql_all;

        string[] sMonth = ("JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC").Split(',');
   
    }
}
