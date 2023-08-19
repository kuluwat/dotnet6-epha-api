using Newtonsoft.Json;
using System.Data;

namespace Class
{
    public class ClassJSON
    {
        public string SetJSONresult(DataTable _dtJson)
        {
            string JSONresult;
            JSONresult = JsonConvert.SerializeObject(_dtJson);
            return JSONresult;
        } 
        public DataTable ConvertJSONresult(String jsper)
        {
            DataTable _dtJson = (DataTable)JsonConvert.DeserializeObject(jsper, typeof(DataTable));
            try
            {
                if (_dtJson != null)
                {
                    if (_dtJson.Rows.Count > 0) { if (_dtJson.Rows[0]["json_check_null"].ToString() == "true") { _dtJson.Rows[0].Delete(); _dtJson.AcceptChanges(); } }
                    _dtJson.Columns.Remove("json_check_null"); _dtJson.AcceptChanges();
                }
            }
            catch { }

            return _dtJson;
        }

    }
}
