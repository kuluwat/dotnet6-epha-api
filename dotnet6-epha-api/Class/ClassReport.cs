using dotnet6_epha_api.Class; 
using System.Data;
namespace Class
{

    public class ClassReport
    {
        string[] sMonth = ("JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC").Split(',');

        public string word_hazop_worksheet(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _export_file_name, string _export_type, DataSet dsData)
        {
            FileInfo template = new FileInfo(_FolderTemplate + "HAZOP Template.docx");

            //Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            //Document doc = wordApp.Documents.Open(template);

            //// Replace placeholders in the template
            //FindAndReplace(doc, "{{Title}}", "Edited Title");
            //FindAndReplace(doc, "{{Content}}", "This is the edited content of the document.");

            //// Save and close the edited document
            //doc.SaveAs2(_Path + _export_file_name);
            //doc.Close();
            //wordApp.Quit();

            return _DownloadPath + _export_file_name;
        }

        //static void FindAndReplace(Document doc, string findText, string replaceWith)
        //{
        //    foreach (Microsoft.Office.Interop.Word.Range range in doc.StoryRanges)
        //    {
        //        range.Find.Execute(findText, Replace: WdReplace.wdReplaceAll, ReplaceWith: replaceWith);
        //    }
        //}
    }
}
