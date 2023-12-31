﻿using Aspose.Cells;
using dotnet6_epha_api.Class;
using iTextSharp.text.pdf;
using Model;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SkiaSharp;
using System.ComponentModel;
using System.Data;
using System.Reflection.Metadata;
using System.Xml.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;
using static System.Net.Mime.MediaTypeNames;

using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Extensions.Options;
using static Org.BouncyCastle.Crypto.Digests.SkeinEngine;
using System.IO.Packaging;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.AspNetCore.Http;
using iTextSharp.text.pdf.qrcode;
using Newtonsoft.Json;
using System;
using Org.BouncyCastle.Ocsp;
using System.Threading.Tasks;

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

        #region function
        private static DataTable refMsg(string status, string remark)
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.Columns.Add("seq_ref");
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            return dtMsg;
        }
        private static DataTable refMsg(string status, string remark, string seq_new)
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.Columns.Add("seq_new");
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            dtMsg.Rows[0]["seq_new"] = seq_new;
            return dtMsg;
        }
        private static DataTable refMsgSave(string status, string remark, string seq_new, string pha_seq, string pha_no)
        {
            DataTable dtMsg = new DataTable();
            dtMsg.Columns.Add("status");
            dtMsg.Columns.Add("remark");
            dtMsg.Columns.Add("seq_new");
            dtMsg.Columns.Add("pha_seq");
            dtMsg.Columns.Add("pha_no");
            dtMsg.AcceptChanges();

            dtMsg.Rows.Add(dtMsg.NewRow());
            dtMsg.Rows[0]["status"] = status;
            dtMsg.Rows[0]["remark"] = remark;
            dtMsg.Rows[0]["seq_new"] = seq_new;
            dtMsg.Rows[0]["pha_seq"] = pha_seq;
            dtMsg.Rows[0]["pha_no"] = pha_no;
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
        public string uploadfile_data(uploadFile uploadFile, string folder)
        {
            DataTable dtdef = new DataTable();
            IFormFileCollection files = uploadFile.file_obj;
            var file_seq = uploadFile.file_seq;
            var file_name = uploadFile.file_name;

            var file_FullName = "";
            var file_FullPath = "";

            string _Folder = "/wwwroot/AttachedFileTemp/" + folder + "/";
            string _DownloadPath = "/AttachedFileTemp/" + folder + "/";
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

        public string config_email_test(EmailConfigModel param)
        {
            string ret = "";
            sqlstr = "insert into  EPHA_M_CONFIGMAIL (seq, email, active_type) values (1," + cls.ChkSqlStr((param.user_email + "").ToString(), 50) + ", 1)";

            cls_conn = new ClassConnectionDb();
            cls_conn.OpenConnection();
            ret = cls_conn.ExecuteNonQuery(sqlstr);
            cls_conn.CloseConnection();
            return ret;
        }

        #endregion function

        #region export excel hazop

        public string export_hazop_report_word(ReportModel param)
        {
            string _FolderTemplate = MapPathFiles("/wwwroot/AttachedFileTemp/");
            string _Folder = MapPathFiles("/wwwroot/AttachedFileTemp/Hazop/");
            string templateFilePath = _FolderTemplate + "HAZOP Template.docx";
            string outputFilePath = _Folder + "HAZOP Template xx.docx";

            using (DocX templateDoc = DocX.Load(templateFilePath))
            {
                // Replace placeholders in the template with actual data
                templateDoc.ReplaceText("{Title}", "Sample Document Title");
                templateDoc.ReplaceText("{Content}", "This is a sample paragraph in the document.");

                // Save the generated document
                templateDoc.SaveAs(outputFilePath);
            }


            return ("Document created successfully.");
        }
        public string word_hazop_report(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _export_file_name, string _export_type)
        {
            DataSet dsData = new DataSet();
            sqlstr = @" select distinct
                        h.seq, nl.id as id_node, g.pha_request_name, convert(varchar,g.create_date,106) as create_date, nl.node, nl.design_intent, nl.descriptions, nl.design_conditions, nl.node_boundary, nl.operating_conditions
                        , d.document_no
                        , mgw.guide_words as guideword, mgw.deviations as deviation, nw.causes, nw.consequences, nw.category_type, nw.ram_befor_security, nw.ram_befor_likelihood, nw.ram_befor_risk
                        , nw.major_accident_event, nw.safety_critical_equipment, nw.safety_critical_equipment_tag, nw.existing_safeguards, nw.ram_after_security, nw.ram_after_likelihood, nw.ram_after_risk
                        , nw.recommendations, nw.recommendations_no, nw.responder_user_displayname
                        , g.descriptions
                        , nl.no as node_no, nw.no, nw.causes_no, nw.consequences_no, nw.category_no
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word    
                        where h.seq = '" + seq + "' ";
            sqlstr += @" order by cast(nl.no as int),cast(nw.no as int), cast(nw.causes_no as int), cast(nw.consequences_no as int), cast(nw.category_no as int)";

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
            dt.TableName = "header";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();

            ClassReport classReport = new ClassReport();
            classReport.word_hazop_worksheet(seq, _Path, _FolderTemplate, _DownloadPath, _export_file_name, _export_type, dsData);
            return "";
        }

        public string export_hazop_report(ReportModel param)
        {
            string seq = param.seq;
            string export_type = param.export_type;

            DataTable dtdef = new DataTable();

            #region Determine whether the directory exists.
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.AcceptChanges();
            dtdef = dt.Clone(); dtdef.AcceptChanges();

            #endregion Determine whether the directory exists.

            string msg_error = "";
            string _DownloadPath = "/AttachedFileTemp/Hazop/";
            string _Folder = "/wwwroot/AttachedFileTemp/Hazop/";
            string _FolderTemplate = MapPathFiles("/wwwroot/AttachedFileTemp/");
            string _Path = MapPathFiles(_Folder);

            var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
            string export_file_name = "HAZOP Report " + datetime_run;
            string export_file_name_full = "";
            if (export_type == "excel" || export_type == "pdf")
            {
                export_file_name_full = excel_hazop_report(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name + ".xlsx", export_type);
                if (export_type == "excel") { export_file_name += ".xlsx"; } else { export_file_name += ".pdf"; }
            }

            try
            {
                dtdef.Rows.Add(dtdef.NewRow()); dtdef.AcceptChanges();
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_NAME"] = export_file_name;
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_PATH"] = export_file_name_full;
                dtdef.Rows[dtdef.Rows.Count - 1]["IMPORT_DATA_MSG"] = msg_error;
                dtdef.AcceptChanges();
            }
            catch (Exception ex) { ex.Message.ToString(); }

            return cls_json.SetJSONresult(dtdef);
        }
        public string excel_hazop_report(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type)
        {
            string export_file_name = _Path + _excel_name;
            FileInfo template = new FileInfo(_FolderTemplate + "HAZOP Report Template.xlsx");
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                excelPackage.SaveAs(new FileInfo(export_file_name));
            }
            //Study Objective and Work Scope, Drawing & Reference, Node List
            excel_hazop_general(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name, export_type, true);

            //HAZOP Attendee Sheet 
            excel_hazop_atendeesheet(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name, export_type, true);

            //HAZOP Recommendation
            excel_hazop_recommendation(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name, export_type, true);

            // Major Accident Event (MAE),
            //MAJOR ACCIDENT EVENT  (Y/N) ให้ดึงที่เป็น Y -> running no , node , cause, R ของ  UNMITIGATED RISK ASSESSMENT MATRIX

            excel_hazop_worksheet(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name, export_type, true);

            excel_hazop_ram(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name, export_type, true, "Hazop");

            excel_hazop_guidewords(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name, export_type, true);

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(export_file_name))
            {
                string SheetName_befor = excelPackage.Workbook.Worksheets[excelPackage.Workbook.Worksheets.Count - 1].Name;
                string SheetName = "Drawing PIDs & PFDs";

                excelPackage.Workbook.Worksheets.MoveAfter(SheetName, SheetName_befor);

                // Save changes
                excelPackage.Save();

                // Save the workbook as PDF
                if (export_type == "pdf")
                {
                    Workbook workbookPDF = new Workbook(export_file_name);
                    PdfSaveOptions options = new PdfSaveOptions
                    {
                        AllColumnsInOnePagePerSheet = true
                    };
                    export_file_name = export_file_name.Replace(".xlsx", ".pdf");

                    workbookPDF.Save(export_file_name, options);

                    add_drawing_to_appendix(seq, _Path, export_file_name, true);

                    if (true)
                    {
                        #region move file to _temp  
                        File.Copy(export_file_name, export_file_name.Replace(@"/Hazop/", @"/_temp/"));
                        try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                        try { File.Delete(export_file_name); } catch { }
                        #endregion move file to _temp
                    }
                    return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/Hazop/", @"/_temp/");

                }
            }

            if (true)
            {
                #region move file to _temp  
                File.Copy(export_file_name, (export_file_name).Replace(@"/Hazop/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
            }
            return (_DownloadPath + _excel_name).Replace(@"/Hazop/", @"/_temp/");
        }
        public string excel_hazop_general(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, Boolean report_all)
        {
            #region get data
            sqlstr = @" select g.work_scope
                         from EPHA_F_HEADER h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                         where h.seq = '" + seq + "' ";

            cls_conn = new ClassConnectionDb();
            DataTable dtWorkScope = new DataTable();
            dtWorkScope = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];


            sqlstr = @" select distinct d.no, d.document_name, d.document_no, d.document_file_name, d.descriptions 
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                        inner join EPHA_T_DRAWING d on h.id = d.id_pha    
                        where h.seq = '" + seq + "' and d.document_name is not null order by convert(int,d.no) ";

            cls_conn = new ClassConnectionDb();
            DataTable dtDrawing = new DataTable();
            dtDrawing = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];


            sqlstr = @" select nl.no, nl.node, nl.design_intent, nl.design_conditions, nl.operating_conditions, nl.node_boundary
                         , d.document_no
                         , isnull(replace(replace( convert(char,nd.page_start_first) + (case when isnull(nd.page_start_first,'') ='' then '' else
                         (case when isnull(nd.page_end_first,'') ='' then '' else 'to'end)  end) 
                         + convert(char,nd.page_end_first)  ,' ',''),'to',' to '),'All') as  document_page
                         from EPHA_F_HEADER h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                         left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                         left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                         where h.seq = '" + seq + "'  and nl.node is not null order by convert(int,nl.no), convert(int,nd.no) ";

            cls_conn = new ClassConnectionDb();
            DataTable dtNode = new DataTable();
            dtNode = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            //MAJOR ACCIDENT EVENT  (Y/N) ให้ดึงที่เป็น Y -> running no , node , cause, R ของ  UNMITIGATED RISK ASSESSMENT MATRIX
            sqlstr = @" select 0 as no, nl.node, nw.causes, nw.causes_no, nw.ram_befor_risk
                         from EPHA_F_HEADER h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_NODE nl on h.id = nl.id_pha  
                         left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node 
                         and lower(isnull(nw.major_accident_event,'')) = lower('Y') 
                         where h.seq = '" + seq + "' order by convert(int,nw.no) ";

            cls_conn = new ClassConnectionDb();
            DataTable dtMajor = new DataTable();
            dtMajor = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            FileInfo template = new FileInfo(_FolderTemplate + "HAZOP Report Template.xlsx");
            if (report_all == true) { template = new FileInfo(_excel_name); }

            #endregion get data

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];  // Replace "SourceSheet" with the actual source sheet name
                ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("HAZOP Attendee Sheet", sourceWorksheet);

                //Study Objective and Work Scope
                worksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];
                worksheet.Cells["A2"].Value = (dtWorkScope.Rows[0]["work_scope"] + "");

                //Drawing & Reference
                #region Drawing & Reference
                if (true)
                {
                    worksheet = excelPackage.Workbook.Worksheets["Drawing & Reference"];

                    int startRows = 3;
                    int icol_end = 6;
                    int ino = 1;
                    for (int i = 0; i < dtDrawing.Rows.Count; i++)
                    {
                        //No.	Document Name	Drawing No	Document File	Comment
                        worksheet.InsertRow(startRows, 1);
                        worksheet.Cells["A" + (i + startRows)].Value = (i + 1); ;
                        worksheet.Cells["B" + (i + startRows)].Value = (dtDrawing.Rows[i]["document_name"] + "");
                        worksheet.Cells["C" + (i + startRows)].Value = (dtDrawing.Rows[i]["document_no"] + "");
                        worksheet.Cells["D" + (i + startRows)].Value = (dtDrawing.Rows[i]["document_file_name"] + "");
                        worksheet.Cells["E" + (i + startRows)].Value = (dtDrawing.Rows[i]["descriptions"] + "");
                        startRows++;
                    }
                    // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                    DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);

                    //var eRange = worksheet.Cells[worksheet.Cells["A3"].Address + ":" + worksheet.Cells["D" + startRows].Address];
                    //eRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //eRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
                #endregion Drawing & Reference

                //Node List
                #region Node List
                if (true)
                {
                    worksheet = excelPackage.Workbook.Worksheets["Drawing & Reference"];

                    int startRows = 3;
                    int icol_end = 6;
                    for (int i = 0; i < dtNode.Rows.Count; i++)
                    {
                        //No.	Node	Design Intent	Design Conditions	Operating Conditions	Node Boundary	Drawing	Drawing Page (From-To)
                        worksheet.InsertRow(startRows, 1);
                        worksheet.Cells["A" + (i + startRows)].Value = (i + 1);
                        worksheet.Cells["B" + (i + startRows)].Value = (dtNode.Rows[i]["node"] + "");
                        worksheet.Cells["C" + (i + startRows)].Value = (dtNode.Rows[i]["design_intent"] + "");
                        worksheet.Cells["D" + (i + startRows)].Value = (dtNode.Rows[i]["design_conditions"] + "");
                        worksheet.Cells["E" + (i + startRows)].Value = (dtNode.Rows[i]["operating_conditions"] + "");
                        worksheet.Cells["F" + (i + startRows)].Value = (dtNode.Rows[i]["node_boundary"] + "");
                        worksheet.Cells["G" + (i + startRows)].Value = (dtNode.Rows[i]["document_no"] + "");
                        worksheet.Cells["H" + (i + startRows)].Value = (dtNode.Rows[i]["document_page"] + "");

                        startRows++;
                    }
                    // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                    DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);
                }
                #endregion Node List

                // Major Accident Event (MAE),
                #region Major Accident Event (MAE)
                if (true)
                {
                    worksheet = excelPackage.Workbook.Worksheets["Major Accident Event (MAE)"];

                    int startRows = 3;
                    int icol_end = 6;
                    for (int i = 0; i < dtMajor.Rows.Count; i++)
                    {
                        //No.	 nl.node, nw.causes, nw.causes_no, nw.ram_befor_risk
                        worksheet.InsertRow(startRows, 1);
                        worksheet.Cells["A" + (i + startRows)].Value = (i + 1);
                        worksheet.Cells["B" + (i + startRows)].Value = (dtMajor.Rows[i]["node"] + "");
                        worksheet.Cells["C" + (i + startRows)].Value = (dtMajor.Rows[i]["causes"] + "");
                        worksheet.Cells["D" + (i + startRows)].Value = (dtMajor.Rows[i]["ram_befor_risk"] + "");

                        startRows++;
                    }
                    // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                    DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);
                }
                #endregion Node List



                //Study Objective and Work Scope
                #region Study Objective and Work Scope
                if (true)
                {
                    worksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];
                    worksheet.Cells["A2"].Value = (dtWorkScope.Rows[0]["work_scope"] + "");
                }
                #endregion Study Objective and Work Scope

                if (report_all == true)
                {
                    //excelPackage.Workbook.Worksheets.MoveBefore("HAZOP Attendee Sheet", "HAZOP Cover Page"); 
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;

                    if (!Directory.Exists(_Path))
                    {
                        Directory.CreateDirectory(_Path);
                    }
                    excelPackage.Save();
                }
                else
                {
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;

                    excelPackage.SaveAs(new FileInfo(_Path + _excel_name));

                    // Save the workbook as PDF
                    if (export_type == "pdf")
                    {
                        Workbook workbookPDF = new Workbook(_Path + _excel_name);
                        PdfSaveOptions options = new PdfSaveOptions
                        {
                            AllColumnsInOnePagePerSheet = true
                        };
                        workbookPDF.Save(_Path + _excel_name.Replace(".xlsx", ".pdf"), options);
                        //return _DownloadPath + _excel_name.Replace(".xlsx", ".pdf");
                        if (true)
                        {
                            #region move file to _temp  
                            string export_file_name = _Path + _excel_name.Replace(".xlsx", ".pdf");
                            File.Copy(export_file_name, export_file_name.Replace(@"/Hazop/", @"/_temp/"));
                            try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                            try { File.Delete(export_file_name); } catch { }
                            #endregion move file to _temp
                        }
                        return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/Hazop/", @"/_temp/");
                    }
                }
            }

            if (!report_all)
            {
                #region move file to _temp  
                string export_file_name = _Path + _excel_name;
                File.Copy(export_file_name, (export_file_name).Replace(@"/Hazop/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
                return (_DownloadPath + _excel_name).Replace(@"/Hazop/", @"/_temp/");
            }
            else { return (_DownloadPath + _excel_name); }
        }

        public string add_drawing_to_appendix(string seq, string _Path, string _excel_name, Boolean report_all)
        {
            //Drawing PIDs & PFDs 
            sqlstr = @" select distinct d.no, d.document_name, d.document_no, d.document_file_name, d.descriptions, h.pha_sub_software as sub_software
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                        inner join EPHA_T_DRAWING d on h.id = d.id_pha    
                        where h.seq = '" + seq + "' and isnull(d.document_file_name,'') <>'' order by convert(int,d.no) ";

            cls_conn = new ClassConnectionDb();
            DataTable dtDrawing = new DataTable();
            dtDrawing = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            string existingPdfFilePath = (_Path + _excel_name);
            if (report_all == true) { existingPdfFilePath = _excel_name.Replace(".xlsx", ".pdf"); }

            for (int i = 0; i < dtDrawing.Rows.Count; i++)
            {
                string pdfToBeAddedFilePath = _Path + (dtDrawing.Rows[i]["document_file_name"] + "");
                addfile_pdf_to_pdf(_Path, existingPdfFilePath, pdfToBeAddedFilePath);
            }

            return existingPdfFilePath;
        }

        public string addfile_pdf_to_pdf(string _Path, string existingPdfFilePath, string pdfToBeAddedFilePath)
        {

            string sourceFilePath = existingPdfFilePath;
            string destinationFilePath = "outfile.pdf";

            try { File.Delete(_Path + destinationFilePath); } catch { }

            File.Copy(sourceFilePath, _Path + destinationFilePath);

            try
            {
                // Create a FileStream for the output PDF 
                using (FileStream outputStream = new FileStream(destinationFilePath, FileMode.Create))
                {
                    // Create a Document object
                    using (iTextSharp.text.Document document = new iTextSharp.text.Document())
                    {
                        // Create a PdfCopy object that will merge the PDFs
                        using (PdfCopy copy = new PdfCopy(document, outputStream))
                        {
                            document.Open();

                            // Open the existing PDF
                            using (PdfReader reader = new PdfReader(existingPdfFilePath))
                            {
                                // Add the existing PDF to the new PDF
                                for (int pageNum = 1; pageNum <= reader.NumberOfPages; pageNum++)
                                {
                                    copy.AddPage(copy.GetImportedPage(reader, pageNum));
                                }
                            }

                            // Open the PDF to be added
                            using (PdfReader pdfToBeAddedReader = new PdfReader(pdfToBeAddedFilePath))
                            {
                                // Add the pages from the PDF to be added to the new PDF
                                for (int pageNum = 1; pageNum <= pdfToBeAddedReader.NumberOfPages; pageNum++)
                                {
                                    copy.AddPage(copy.GetImportedPage(pdfToBeAddedReader, pageNum));
                                }
                            }
                        }
                    }

                }


                //delete source file
                try { File.Delete(sourceFilePath); } catch { }

                //copy destination to source file
                File.Copy(destinationFilePath, sourceFilePath);

                //delete destination file
                try { File.Delete(destinationFilePath); } catch { }



            }
            catch { }

            return "";
        }

        public string export_hazop_atendeesheet(ReportModel param)
        {
            string seq = param.seq;
            string export_type = param.export_type;

            DataTable dtdef = new DataTable();

            #region Determine whether the directory exists.
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.AcceptChanges();
            dtdef = dt.Clone(); dtdef.AcceptChanges();

            #endregion Determine whether the directory exists.

            string msg_error = "";
            string _DownloadPath = "/AttachedFileTemp/Hazop/";
            string _Folder = "/wwwroot/AttachedFileTemp/Hazop/";
            string _FolderTemplate = MapPathFiles("/wwwroot/AttachedFileTemp/");
            string _Path = MapPathFiles(_Folder);

            var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
            string export_file_name = "HAZOP AttendeeSheet " + datetime_run;
            string export_file_name_full = "";
            if (export_type == "excel" || export_type == "pdf")
            {
                export_file_name_full = excel_hazop_atendeesheet(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name + ".xlsx", export_type, false);
                if (export_type == "excel") { export_file_name += ".xlsx"; } else { export_file_name += ".pdf"; }

            }

            try
            {
                dtdef.Rows.Add(dtdef.NewRow()); dtdef.AcceptChanges();
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_NAME"] = export_file_name;
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_PATH"] = export_file_name_full;
                dtdef.Rows[dtdef.Rows.Count - 1]["IMPORT_DATA_MSG"] = msg_error;
                dtdef.AcceptChanges();
            }
            catch (Exception ex) { ex.Message.ToString(); }

            return cls_json.SetJSONresult(dtdef);
        }

        public string excel_hazop_atendeesheet(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, Boolean report_all)
        {

            sqlstr = @" select s.id_pha, s.seq as seq_session, s.no as session_no
                         , convert(varchar,s.meeting_date,106) as meeting_date
                         , mt.no as member_no, isnull(mt.user_name,'') as user_name, emp.user_displayname
                         from EPHA_F_HEADER h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_SESSION s on h.id = s.id_pha 
                         left join EPHA_T_MEMBER_TEAM mt on h.id = mt. id_pha and mt.id_session = s.id
                         left join VW_EPHA_PERSON_DETAILS emp on lower(emp.user_name) = lower(mt.user_name)
                         where h.seq = '" + seq + "' and lower(mt.user_name) is not null ";

            cls_conn = new ClassConnectionDb();
            DataTable dtAll = new DataTable();
            dtAll = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            cls_conn = new ClassConnectionDb();
            DataTable dtMember = new DataTable();
            dtMember = cls_conn.ExecuteAdapterSQL(" select distinct 0 as no, t.user_name, t.user_displayname, '' as company_text from (" + sqlstr + " )t where t.user_name <> '' order by t.user_name").Tables[0];

            cls_conn = new ClassConnectionDb();
            DataTable dtSession = new DataTable();
            dtSession = cls_conn.ExecuteAdapterSQL(" select distinct t.seq_session, t.session_no, t.meeting_date from (" + sqlstr + ")t order by t.session_no ").Tables[0];

            Boolean bCheckNewFile = false;
            FileInfo template = new FileInfo(_FolderTemplate + "HAZOP AttendeeSheet Template.xlsx");
            if (report_all == true) { template = new FileInfo(_excel_name); }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];  // Replace "SourceSheet" with the actual source sheet name
                sourceWorksheet.Name = "HAZOP Attendee Sheet";
                ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("HAZOP Attendee Sheet", sourceWorksheet);

                int i = 0;
                int startRows = 4;
                int icol_start = 4;
                int icol_end = icol_start + (dtSession.Rows.Count > 6 ? dtSession.Rows.Count : 6);

                for (int imember = 0; imember < dtMember.Rows.Count; imember++)
                {
                    worksheet.InsertRow(startRows, 1);
                    string user_name = (dtMember.Rows[imember]["user_name"] + "");
                    //No.
                    worksheet.Cells["A" + (i + startRows)].Value = (imember + 1);
                    //Name
                    worksheet.Cells["B" + (i + startRows)].Value = (dtMember.Rows[imember]["user_displayname"] + "");
                    //Company
                    worksheet.Cells["C" + (i + startRows)].Value = (dtMember.Rows[imember]["company_text"] + "");

                    int irow_session = 0;
                    if (imember == 0)
                    {
                        if (dtSession.Rows.Count < 6)
                        {
                            //worksheet.Cells[2, icol_start, 2, icol_end].Merge = true; 
                            for (int c = icol_end; c < 30; c++)
                            {
                                worksheet.DeleteColumn(icol_end);

                            }
                        }

                        irow_session = 0;
                        for (int c = icol_start; c < icol_end; c++)
                        {
                            try
                            {
                                //header 
                                if ((dtSession.Rows[irow_session]["meeting_date"] + "") == "")
                                {
                                    worksheet.Cells[3, c].Value = "";
                                }
                                else
                                {
                                    worksheet.Cells[3, c].Value = (dtSession.Rows[irow_session]["meeting_date"] + "");
                                }
                            }
                            catch { worksheet.Cells[3, c].Value = ""; }
                            irow_session += 1;
                        }
                    }

                    irow_session = 0;
                    for (int c = icol_start; c < icol_end; c++)
                    {
                        try
                        {
                            string session_no = "";
                            try { session_no = (dtSession.Rows[irow_session]["session_no"] + ""); } catch { }

                            DataRow[] dr = dtAll.Select("user_name = '" + user_name + "' and session_no = '" + session_no + "'");
                            if (dr.Length > 0)
                            {
                                worksheet.Cells[startRows, c].Value = "X";
                            }
                            else { worksheet.Cells[startRows, c].Value = ""; }
                        }
                        catch { }
                        irow_session++;

                    }

                    startRows++;
                }

                // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);

                if (report_all == true)
                {
                    //excelPackage.Workbook.Worksheets.MoveBefore("HAZOP Attendee Sheet", "Study Objective and Work Scope"); 
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;

                    if (!Directory.Exists(_Path))
                    {
                        Directory.CreateDirectory(_Path);
                    }
                    excelPackage.Save();
                }
                else
                {
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;

                    excelPackage.SaveAs(new FileInfo(_Path + _excel_name));

                    // Save the workbook as PDF
                    if (export_type == "pdf")
                    {
                        Workbook workbookPDF = new Workbook(_Path + _excel_name);
                        PdfSaveOptions options = new PdfSaveOptions
                        {
                            AllColumnsInOnePagePerSheet = true
                        };
                        workbookPDF.Save(_Path + _excel_name.Replace(".xlsx", ".pdf"), options);
                        //return _DownloadPath + _excel_name.Replace(".xlsx", ".pdf");
                        if (true)
                        {
                            #region move file to _temp  
                            string export_file_name = _Path + _excel_name.Replace(".xlsx", ".pdf");
                            File.Copy(export_file_name, export_file_name.Replace(@"/Hazop/", @"/_temp/"));
                            try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                            try { File.Delete(export_file_name); } catch { }
                            #endregion move file to _temp
                        }
                        return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/Hazop/", @"/_temp/");
                    }
                }
            }

            //return _DownloadPath + _excel_name;  
            if (!report_all)
            {
                #region move file to _temp  
                string export_file_name = _Path + _excel_name;
                File.Copy(export_file_name, (export_file_name).Replace(@"/Hazop/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
                return (_DownloadPath + _excel_name).Replace(@"/Hazop/", @"/_temp/");
            }
            else { return (_DownloadPath + _excel_name); }
        }

        public string export_hazop_worksheet(ReportModel param)
        {
            string seq = param.seq;
            string export_type = param.export_type;

            DataTable dtdef = new DataTable();

            #region Determine whether the directory exists.
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.AcceptChanges();
            dtdef = dt.Clone(); dtdef.AcceptChanges();

            #endregion Determine whether the directory exists.

            string msg_error = "";
            string _DownloadPath = "/AttachedFileTemp/Hazop/";
            string _Folder = "/wwwroot/AttachedFileTemp/Hazop/";
            string _FolderTemplate = MapPathFiles("/wwwroot/AttachedFileTemp/");
            string _Path = MapPathFiles(_Folder);

            var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
            string export_file_name = "HAZOP WORKSHEET & RECOMMENDATION RESPONSE SHEET " + datetime_run;
            string export_file_name_full = "";
            if (export_type == "excel" || export_type == "pdf")
            {
                export_file_name_full = excel_hazop_worksheet(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name + ".xlsx", export_type, false);
                if (export_type == "excel") { export_file_name += ".xlsx"; } else { export_file_name += ".pdf"; }

            }

            try
            {
                dtdef.Rows.Add(dtdef.NewRow()); dtdef.AcceptChanges();
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_NAME"] = export_file_name;
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_PATH"] = export_file_name_full;
                dtdef.Rows[dtdef.Rows.Count - 1]["IMPORT_DATA_MSG"] = msg_error;
                dtdef.AcceptChanges();
            }
            catch (Exception ex) { ex.Message.ToString(); }

            return cls_json.SetJSONresult(dtdef);
        }

        public string excel_hazop_worksheet(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, Boolean report_all)
        {
            sqlstr = @" select distinct nl.no, nl.id as id_node
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        where h.seq = '" + seq + "' ";
            sqlstr += @" order by cast(nl.no as int)";
            cls_conn = new ClassConnectionDb();
            DataTable dtNode = new DataTable();
            dtNode = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            sqlstr = @" select distinct
                        h.seq, nl.id as id_node, g.pha_request_name, convert(varchar,g.create_date,106) as create_date, nl.node, nl.design_intent, nl.descriptions as descriptions_worksheet, nl.design_conditions, nl.node_boundary, nl.operating_conditions
                        , d.document_no
                        , mgw.guide_words as guideword, mgw.deviations as deviation, nw.causes, nw.consequences, nw.category_type, nw.ram_befor_security, nw.ram_befor_likelihood, nw.ram_befor_risk
                        , nw.major_accident_event, nw.safety_critical_equipment, nw.safety_critical_equipment_tag, nw.existing_safeguards, nw.ram_after_security, nw.ram_after_likelihood, nw.ram_after_risk
                        , nw.recommendations, nw.recommendations_no, nw.responder_user_displayname
                        , g.descriptions
                        , nl.no as node_no, nw.no, nw.causes_no, nw.consequences_no, nw.category_no
                        , case when g.id_ram = 5 then 1 else 0 end show_cat
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word    
                        where h.seq = '" + seq + "' ";
            sqlstr += @" order by cast(nl.no as int),cast(nw.no as int), cast(nw.causes_no as int), cast(nw.consequences_no as int), cast(nw.category_no as int)";

            cls_conn = new ClassConnectionDb();
            DataTable dtWorksheet = new DataTable();
            dtWorksheet = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            sqlstr = @"  select distinct
                         h.seq,nl.no as node_no,nl.node, 0 as no, nw.safety_critical_equipment_tag
                         , str(nw.consequences_no) + '.' + nw.consequences as consequences, isnull(nw.ram_befor_risk,'') as  ram_befor_risk
                         from EPHA_F_HEADER h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_NODE nl on h.id = nl.id_pha  
                         left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                         left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word    
                         where nw.safety_critical_equipment = 'Y'  
                         and h.seq = '" + seq + "' ";
            sqlstr += @" order by cast(nl.no as int),nl.node, nw.safety_critical_equipment_tag  ";

            cls_conn = new ClassConnectionDb();
            DataTable dtSCE = new DataTable();
            dtSCE = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            Boolean bCheckNewFile = false;
            FileInfo template = new FileInfo(_FolderTemplate + "HAZOP Study Worksheet Template.xlsx");
            if (report_all == true)
            {
                template = new FileInfo(_excel_name);
                if (!template.Exists) { template = new FileInfo(_FolderTemplate + "HAZOP Report Template.xlsx"); bCheckNewFile = true; }
            }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                string worksheet_name = "";
                string worksheet_name_target = "";
                for (int inode = 0; inode < dtNode.Rows.Count; inode++)
                {
                    if (worksheet_name_target == "") { worksheet_name_target = "WorksheetTemplate"; }
                    else { worksheet_name_target = "HAZOP Worksheet Node (" + (inode) + ")"; }
                    worksheet_name = "HAZOP Worksheet Node (" + (inode + 1) + ")";

                    ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["WorksheetTemplate"];  // Replace "SourceSheet" with the actual source sheet name
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(worksheet_name, sourceWorksheet);

                    string id_node = (dtNode.Rows[inode]["id_node"] + "");

                    int i = 0;
                    int startRows = 3;

                    DataRow[] dr = dtWorksheet.Select("id_node=" + id_node);

                    string show_cat = (dr[0]["show_cat"] + "");
                    if (dr.Length > 0)
                    {
                        #region head text
                        i = 0;
                        //Project
                        worksheet.Cells["B" + (i + startRows)].Value = (dr[0]["pha_request_name"] + "");
                        //NODE
                        worksheet.Cells["N" + (i + startRows)].Value = (dr[0]["node"] + "");
                        startRows++;

                        //Design Intent :
                        worksheet.Cells["B" + (i + startRows)].Value = (dr[0]["design_intent"] + "");
                        //System
                        worksheet.Cells["N" + (i + startRows)].Value = (dr[0]["descriptions"] + "");
                        startRows++;

                        //"Design Conditions: -->design_conditions
                        worksheet.Cells["B" + (i + startRows)].Value = (dr[0]["design_conditions"] + "");
                        //HAZOP Boundary
                        worksheet.Cells["N" + (i + startRows)].Value = (dr[0]["node_boundary"] + "");
                        startRows++;

                        //"Operating Conditions: -->operating_conditions
                        worksheet.Cells["B" + (i + startRows)].Value = (dr[0]["operating_conditions"] + "");
                        startRows++;

                        //PFD, PID No. : --> document_no
                        worksheet.Cells["B" + (i + startRows)].Value = (dr[0]["document_no"] + "");
                        //Date
                        worksheet.Cells["N" + (i + startRows)].Value = (dr[0]["create_date"] + "");
                        startRows++;

                        #endregion head text
                        startRows = 14;
                        for (i = 0; i < dr.Length; i++)
                        {
                            worksheet.InsertRow(startRows, 1);

                            worksheet.Cells["A" + (startRows)].Value = dr[i]["guideword"].ToString();
                            worksheet.Cells["B" + (startRows)].Value = dr[i]["deviation"].ToString();
                            worksheet.Cells["C" + (startRows)].Value = dr[i]["causes"].ToString();
                            worksheet.Cells["D" + (startRows)].Value = dr[i]["consequences"].ToString();
                            worksheet.Cells["E" + (startRows)].Value = dr[i]["category_type"].ToString();

                            worksheet.Cells["F" + (startRows)].Value = dr[i]["ram_befor_security"].ToString();
                            worksheet.Cells["G" + (startRows)].Value = dr[i]["ram_befor_likelihood"].ToString();
                            worksheet.Cells["H" + (startRows)].Value = dr[i]["ram_befor_risk"];
                            worksheet.Cells["I" + (startRows)].Value = dr[i]["major_accident_event"].ToString();
                            worksheet.Cells["J" + (startRows)].Value = dr[i]["existing_safeguards"].ToString();

                            worksheet.Cells["K" + (startRows)].Value = dr[i]["ram_after_security"].ToString();
                            worksheet.Cells["L" + (startRows)].Value = dr[i]["ram_after_likelihood"].ToString();
                            worksheet.Cells["M" + (startRows)].Value = dr[i]["ram_after_risk"].ToString();
                            worksheet.Cells["N" + (startRows)].Value = dr[i]["recommendations_no"].ToString();
                            worksheet.Cells["O" + (startRows)].Value = dr[i]["recommendations"].ToString();
                            worksheet.Cells["P" + (startRows)].Value = dr[i]["responder_user_displayname"].ToString();


                            startRows++;
                        }
                        // วาดเส้นตาราง โดยใช้เซลล์ A3 ถึง P3 
                        DrawTableBorders(worksheet, 14, 1, startRows - 1, 16);

                        worksheet.Cells["A" + (startRows)].Value = (dr[0]["descriptions_worksheet"] + "");

                        if (show_cat == "0")
                        {
                            worksheet.DeleteColumn(5);
                        }

                    }

                    //new worksheet move after WorksheetTemplate 
                    excelPackage.Workbook.Worksheets.MoveBefore(worksheet_name, worksheet_name_target);
                }
                if (report_all == true)
                {
                    if (dtSCE.Rows.Count > 0)
                    {
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Safety Critical Equipment"];

                        int startRows = 3;
                        for (int s = 0; s < dtSCE.Rows.Count; s++)
                        {
                            worksheet.InsertRow(startRows, 1);

                            worksheet.Cells["A" + (startRows)].Value = (s + 1);
                            if (s > 0)
                            {
                                if (dtSCE.Rows[s - 1]["node"].ToString() != dtSCE.Rows[s]["node"].ToString())
                                {
                                    worksheet.Cells["B" + (startRows)].Value = dtSCE.Rows[s]["node"].ToString();
                                }
                            }
                            else
                            {
                                worksheet.Cells["B" + (startRows)].Value = dtSCE.Rows[s]["node"].ToString();
                            }
                            worksheet.Cells["C" + (startRows)].Value = dtSCE.Rows[s]["safety_critical_equipment_tag"].ToString();
                            worksheet.Cells["D" + (startRows)].Value = dtSCE.Rows[s]["consequences"].ToString();
                            worksheet.Cells["E" + (startRows)].Value = dtSCE.Rows[s]["ram_befor_risk"].ToString();
                            startRows++;
                        }
                        // วาดเส้นตาราง โดยใช้เซลล์ A1 ถึง E3
                        DrawTableBorders(worksheet, 3, 1, startRows - 1, 5);

                    }
                }

                if (report_all == true && bCheckNewFile == false)
                {
                    //"HAZOP Worksheet Node(" + (inode + 1) +")"
                    ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["WorksheetTemplate"];
                    SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                    excelPackage.Save();
                }
                else
                {
                    ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["WorksheetTemplate"];
                    SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                    excelPackage.SaveAs(new FileInfo(_Path + _excel_name));

                    // Save the workbook as PDF
                    if (export_type == "pdf")
                    {
                        Workbook workbookPDF = new Workbook(_Path + _excel_name);
                        PdfSaveOptions options = new PdfSaveOptions
                        {
                            AllColumnsInOnePagePerSheet = true
                        };
                        workbookPDF.Save(_Path + _excel_name.Replace(".xlsx", ".pdf"), options);
                        //return _DownloadPath + _excel_name.Replace(".xlsx", ".pdf");
                        if (true)
                        {
                            #region move file to _temp  
                            string export_file_name = _Path + _excel_name.Replace(".xlsx", ".pdf");
                            File.Copy(export_file_name, export_file_name.Replace(@"/Hazop/", @"/_temp/"));
                            try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                            try { File.Delete(export_file_name); } catch { }
                            #endregion move file to _temp
                        }
                        return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/Hazop/", @"/_temp/");

                    }
                }
            }


            //return _DownloadPath + _excel_name;  
            if (!report_all)
            {
                #region move file to _temp  
                string export_file_name = _Path + _excel_name;
                File.Copy(export_file_name, (export_file_name).Replace(@"/Hazop/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
                return (_DownloadPath + _excel_name).Replace(@"/Hazop/", @"/_temp/");
            }
            else { return (_DownloadPath + _excel_name); }
        }
        public string excel_hazop_safety_critical_equipment(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, Boolean report_all)
        {

            sqlstr = @"  select distinct
                         h.seq,nl.no as node_no,nl.node, 0 as no, nw.safety_critical_equipment_tag
                         , str(nw.consequences_no) + '.' + nw.consequences
                         from EPHA_F_HEADER h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_NODE nl on h.id = nl.id_pha  
                         left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                         left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word    
                         where nw.safety_critical_equipment = 'Y'  
                         and h.seq = '" + seq + "' ";
            sqlstr += @" order by cast(nl.no as int), nw.safety_critical_equipment_tag  ";


            cls_conn = new ClassConnectionDb();
            DataTable dtAll = new DataTable();
            dtAll = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            cls_conn = new ClassConnectionDb();
            DataTable dtMember = new DataTable();
            dtMember = cls_conn.ExecuteAdapterSQL(" select distinct 0 as no, t.user_name, t.user_displayname, '' as company_text from (" + sqlstr + " )t where t.user_name <> '' order by t.user_name").Tables[0];

            cls_conn = new ClassConnectionDb();
            DataTable dtSession = new DataTable();
            dtSession = cls_conn.ExecuteAdapterSQL(" select distinct t.seq_session, t.session_no, t.meeting_date from (" + sqlstr + ")t order by t.session_no ").Tables[0];

            Boolean bCheckNewFile = false;
            FileInfo template = new FileInfo(_FolderTemplate + "HAZOP AttendeeSheet Template.xlsx");
            if (report_all == true) { template = new FileInfo(_excel_name); }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];  // Replace "SourceSheet" with the actual source sheet name
                sourceWorksheet.Name = "HAZOP Attendee Sheet";
                ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("HAZOP Attendee Sheet", sourceWorksheet);

                int i = 0;
                int startRows = 4;
                int icol_start = 4;
                int icol_end = icol_start + (dtSession.Rows.Count > 6 ? dtSession.Rows.Count : 6);

                for (int imember = 0; imember < dtMember.Rows.Count; imember++)
                {
                    worksheet.InsertRow(startRows, 1);
                    string user_name = (dtMember.Rows[imember]["user_name"] + "");
                    //No.
                    worksheet.Cells["A" + (i + startRows)].Value = (imember + 1);
                    //Name
                    worksheet.Cells["B" + (i + startRows)].Value = (dtMember.Rows[imember]["user_displayname"] + "");
                    //Company
                    worksheet.Cells["C" + (i + startRows)].Value = (dtMember.Rows[imember]["company_text"] + "");

                    int irow_session = 0;
                    if (imember == 0)
                    {
                        if (dtSession.Rows.Count < 6)
                        {
                            //worksheet.Cells[2, icol_start, 2, icol_end].Merge = true; 
                            for (int c = icol_end; c < 30; c++)
                            {
                                worksheet.DeleteColumn(icol_end);

                            }
                        }

                        irow_session = 0;
                        for (int c = icol_start; c < icol_end; c++)
                        {
                            try
                            {
                                //header 
                                if ((dtSession.Rows[irow_session]["meeting_date"] + "") == "")
                                {
                                    worksheet.Cells[3, c].Value = "";
                                }
                                else
                                {
                                    worksheet.Cells[3, c].Value = (dtSession.Rows[irow_session]["meeting_date"] + "");
                                }
                            }
                            catch { worksheet.Cells[3, c].Value = ""; }
                            irow_session += 1;
                        }
                    }

                    irow_session = 0;
                    for (int c = icol_start; c < icol_end; c++)
                    {
                        try
                        {
                            string session_no = "";
                            try { session_no = (dtSession.Rows[irow_session]["session_no"] + ""); } catch { }

                            DataRow[] dr = dtAll.Select("user_name = '" + user_name + "' and session_no = '" + session_no + "'");
                            if (dr.Length > 0)
                            {
                                worksheet.Cells[startRows, c].Value = "X";
                            }
                            else { worksheet.Cells[startRows, c].Value = ""; }
                        }
                        catch { }
                        irow_session++;

                    }

                    startRows++;
                }

                // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);

                if (report_all == true)
                {
                    //excelPackage.Workbook.Worksheets.MoveBefore("HAZOP Attendee Sheet", "Study Objective and Work Scope"); 
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;

                    if (!Directory.Exists(_Path))
                    {
                        Directory.CreateDirectory(_Path);
                    }
                    excelPackage.Save();
                }
                else
                {
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;

                    excelPackage.SaveAs(new FileInfo(_Path + _excel_name));

                    // Save the workbook as PDF
                    if (export_type == "pdf")
                    {
                        Workbook workbookPDF = new Workbook(_Path + _excel_name);
                        PdfSaveOptions options = new PdfSaveOptions
                        {
                            AllColumnsInOnePagePerSheet = true
                        };
                        workbookPDF.Save(_Path + _excel_name.Replace(".xlsx", ".pdf"), options);
                        //return _DownloadPath + _excel_name.Replace(".xlsx", ".pdf");
                        if (true)
                        {
                            #region move file to _temp  
                            string export_file_name = _Path + _excel_name.Replace(".xlsx", ".pdf");
                            File.Copy(export_file_name, export_file_name.Replace(@"/Hazop/", @"/_temp/"));
                            try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                            try { File.Delete(export_file_name); } catch { }
                            #endregion move file to _temp
                        }
                        return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/Hazop/", @"/_temp/");
                    }
                }
            }


            //return _DownloadPath + _excel_name;  
            if (!report_all)
            {
                #region move file to _temp  
                string export_file_name = _Path + _excel_name;
                File.Copy(export_file_name, (export_file_name).Replace(@"/Hazop/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
                return (_DownloadPath + _excel_name).Replace(@"/Hazop/", @"/_temp/");
            }
            else { return (_DownloadPath + _excel_name); }
        }

        public string export_hazop_recommendation(ReportModel param)
        {
            string seq = param.seq;
            string export_type = param.export_type;

            DataTable dtdef = new DataTable();

            #region Determine whether the directory exists.
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.AcceptChanges();
            dtdef = dt.Clone(); dtdef.AcceptChanges();

            #endregion Determine whether the directory exists.

            string msg_error = "";
            string _DownloadPath = "/AttachedFileTemp/Hazop/";
            string _Folder = "/wwwroot/AttachedFileTemp/Hazop/";
            string _FolderTemplate = MapPathFiles("/wwwroot/AttachedFileTemp/");
            string _Path = MapPathFiles(_Folder);

            var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
            string export_file_name = "HAZOP RECOMMENDATION RESPONSE SHEET & RECCOMENDATION STATUS TRACKING TABLE " + datetime_run;
            string export_file_name_full = "";
            if (export_type == "excel" || export_type == "pdf")
            {
                export_file_name_full = excel_hazop_recommendation(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name + ".xlsx", export_type, false);
                if (export_type == "excel") { export_file_name += ".xlsx"; } else { export_file_name += ".pdf"; }
            }

            try
            {
                dtdef.Rows.Add(dtdef.NewRow()); dtdef.AcceptChanges();
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_NAME"] = export_file_name;
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_PATH"] = export_file_name_full;
                dtdef.Rows[dtdef.Rows.Count - 1]["IMPORT_DATA_MSG"] = msg_error;
                dtdef.AcceptChanges();
            }
            catch (Exception ex) { ex.Message.ToString(); }

            return cls_json.SetJSONresult(dtdef);
        }

        public string excel_hazop_recommendation(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, Boolean report_all)
        {
            sqlstr = @" select distinct
                        h.seq, h.pha_no, nl.id as id_node, g.pha_request_name
                        , nl.node, nl.node as node_check, nl.design_intent, nl.descriptions, nl.design_conditions, nl.node_boundary, nl.operating_conditions
                        , d.document_no, d.document_file_name
                        , mgw.guide_words as guideword, mgw.deviations as deviation, nw.causes, nw.consequences
                        , nw.category_type, nw.ram_befor_security, nw.ram_befor_likelihood, nw.ram_befor_risk
                        , nw.existing_safeguards, nw.recommendations, nw.recommendations_no, nw.responder_user_name, nw.responder_user_displayname
                        , nw.action_status
                        , nl.no as node_no, nw.no, nw.causes_no, nw.consequences_no
                        , nw.seq as seq_worksheet
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word
                        where h.seq = '" + seq + "' and nw.responder_user_name is not null ";
            sqlstr += @" order by cast(nl.no as int),cast(nw.no as int), cast(nw.causes_no as int), cast(nw.consequences_no as int)";

            cls_conn = new ClassConnectionDb();
            DataTable dtWorksheet = new DataTable();
            dtWorksheet = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            sqlstr = @" select distinct nl.no, nw.no, nw.seq, 0 as ref, nl.node, nl.node as node_check
                        , nw.ram_after_risk, nw.ram_after_risk_action, nw.recommendations, nw.recommendations_no, nw.action_status, nw.responder_user_name, nw.responder_user_displayname 
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word
                        where h.seq = '" + seq + "' and nw.responder_user_name is not null ";
            sqlstr += @" order by nl.no, nw.no ";
            cls_conn = new ClassConnectionDb();
            DataTable dtTrack = new DataTable();
            dtTrack = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
            if (true)
            {
                for (int t = 0; t < dtTrack.Rows.Count; t++)
                {
                    dtTrack.Rows[t]["ref"] = (t + 1);
                    dtTrack.AcceptChanges();
                }
            }


            FileInfo template = new FileInfo(_FolderTemplate + "HAZOP Recommendation Template.xlsx");
            if (report_all == true) { template = new FileInfo(_excel_name); }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {

                dt = new DataTable(); dt = dtWorksheet.Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    #region Sheet
                    if (true)
                    {
                        ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["RecommTemplate"];
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("RecommTemplate" + i, sourceWorksheet);

                        string ref_no = (i + 1).ToString();
                        worksheet.Name = "Response Sheet(Ref." + ref_no + ")";

                        string responder_user_name = (dt.Rows[i]["responder_user_name"] + "");
                        string responder_user_displayname = (dt.Rows[i]["responder_user_displayname"] + "");
                        string pha_request_name = (dt.Rows[i]["pha_request_name"] + "");
                        string pha_no = (dt.Rows[i]["pha_no"] + "");
                        string seq_worksheet = (dt.Rows[i]["seq_worksheet"] + "");


                        int startRows = 2;
                        if (true)
                        {
                            string node = "";
                            string drawing_doc = "";
                            string deviation = "";
                            string causes = "";
                            string consequences = "";
                            string existing_safeguards = "";
                            string recommendations = "";
                            string recommendations_no = "";
                            int action_no = 0;

                            #region loop drawing_doc 
                            drawing_doc = (dt.Rows[i]["document_no"] + "");
                            if ((dt.Rows[i]["document_file_name"] + "") != "")
                            {
                                drawing_doc += " (" + dt.Rows[i]["document_file_name"] + ")";
                            }
                            #endregion loop drawing_doc 

                            #region loop workksheet
                            DataRow[] drWorksheet = dt.Select("seq_worksheet = '" + seq_worksheet + "'");
                            for (int n = 0; n < drWorksheet.Length; n++)
                            {
                                if ((drWorksheet[n]["deviation"] + "") != "")
                                {
                                    if (deviation != "") { deviation += ","; }
                                    deviation += (drWorksheet[n]["guideword"] + "") + "/" + (drWorksheet[n]["deviation"] + "");
                                }
                                if ((drWorksheet[n]["causes"] + "") != "")
                                {
                                    if (causes != "") { causes += ","; }
                                    causes += (drWorksheet[n]["causes"] + "");
                                }
                                if ((drWorksheet[n]["consequences"] + "") != "")
                                {
                                    if (consequences != "") { consequences += ","; }
                                    consequences += (drWorksheet[n]["consequences"] + "");
                                }

                                if ((drWorksheet[n]["existing_safeguards"] + "") != "")
                                {
                                    if (existing_safeguards.IndexOf((drWorksheet[n]["existing_safeguards"] + "")) > -1) { }
                                    else
                                    {
                                        if (existing_safeguards != "") { existing_safeguards += ","; }
                                        existing_safeguards += (drWorksheet[n]["existing_safeguards"] + "");
                                    }
                                }

                                if ((drWorksheet[n]["recommendations"] + "") != "")
                                {
                                    if (recommendations != "") { recommendations += ","; }
                                    recommendations += (drWorksheet[n]["recommendations"] + "");
                                    action_no += 1;

                                    if (recommendations_no != "") { recommendations_no += ","; }
                                    recommendations_no += (drWorksheet[n]["recommendations_no"] + "");
                                }

                            }

                            #endregion loop workksheet

                            worksheet.Cells["A" + (startRows)].Value = "Project Title:" + pha_request_name;
                            startRows += 1;
                            worksheet.Cells["A" + (startRows)].Value = "Project No:" + pha_no;
                            startRows += 1;
                            worksheet.Cells["A" + (startRows)].Value = "Node:" + node;
                            startRows += 1;

                            worksheet.Cells["B" + (startRows)].Value = responder_user_displayname;
                            worksheet.Cells["E" + (startRows)].Value = responder_user_displayname;
                            startRows += 1;

                            worksheet.Cells["B" + (startRows)].Value = action_no;
                            startRows += 1;

                            worksheet.Cells["B" + (startRows)].Value = drawing_doc;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = deviation;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = causes;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = consequences;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = existing_safeguards;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = recommendations;
                            startRows += 1;
                        }

                    }
                    #endregion Sheet

                }

                #region TrackTemplate
                if (dtTrack.Rows.Count > 0)
                {
                    //ข้อมูลทั้งหมด
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TrackTemplate"];
                    worksheet.Name = "Status Tracking Table";

                    int i = 0;
                    int startRows = 3;

                    dt = new DataTable(); dt = dtTrack.Copy(); dt.AcceptChanges();
                    if (dt.Rows.Count > 0)
                    {
                        for (i = 0; i < dt.Rows.Count; i++)
                        {
                            worksheet.InsertRow(startRows, 1);
                            worksheet.Cells["A" + (startRows)].Value = dt.Rows[i]["ref"].ToString();
                            worksheet.Cells["B" + (startRows)].Value = dt.Rows[i]["node"].ToString();
                            worksheet.Cells["C" + (startRows)].Value = dt.Rows[i]["ram_after_risk"].ToString();
                            worksheet.Cells["D" + (startRows)].Value = dt.Rows[i]["recommendations"].ToString();
                            worksheet.Cells["E" + (startRows)].Value = dt.Rows[i]["action_status"].ToString();
                            worksheet.Cells["F" + (startRows)].Value = dt.Rows[i]["responder_user_displayname"].ToString();
                            startRows++;
                        }

                        // วาดเส้นตาราง โดยใช้เซลล์ A1 ถึง C3
                        DrawTableBorders(worksheet, 3, 1, startRows - 1, 6);
                    }
                }
                #endregion Response Sheet

                if (report_all == true)
                {
                    ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["RecommTemplate"];
                    SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                    excelPackage.Save();
                }
                else
                {
                    ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["RecommTemplate"];
                    SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                    excelPackage.SaveAs(new FileInfo(_Path + _excel_name));

                    // Save the workbook as PDF
                    if (export_type == "pdf")
                    {
                        Workbook workbookPDF = new Workbook(_Path + _excel_name);
                        PdfSaveOptions options = new PdfSaveOptions
                        {
                            AllColumnsInOnePagePerSheet = true
                        };
                        workbookPDF.Save(_Path + _excel_name.Replace(".xlsx", ".pdf"), options);
                        //return _DownloadPath + _excel_name.Replace(".xlsx", ".pdf");
                        if (true)
                        {
                            #region move file to _temp  
                            string export_file_name = _Path + _excel_name.Replace(".xlsx", ".pdf");
                            File.Copy(export_file_name, export_file_name.Replace(@"/Hazop/", @"/_temp/"));
                            try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                            try { File.Delete(export_file_name); } catch { }
                            #endregion move file to _temp
                        }
                        return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/Hazop/", @"/_temp/");

                    }
                }
            }


            //return _DownloadPath + _excel_name;  
            if (!report_all)
            {
                #region move file to _temp  
                string export_file_name = _Path + _excel_name;
                File.Copy(export_file_name, (export_file_name).Replace(@"/Hazop/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
                return (_DownloadPath + _excel_name).Replace(@"/Hazop/", @"/_temp/");
            }
            else { return (_DownloadPath + _excel_name); }
        }

        public string excel_hazop_recommendation_by_responder(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, Boolean report_all)
        {
            sqlstr = @" select distinct h.pha_no, g.pha_request_name, nw.responder_user_name, nw.responder_user_displayname
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        where h.seq = '" + seq + "' and nw.responder_user_name is not null ";
            sqlstr += @" order by nw.responder_user_name";
            cls_conn = new ClassConnectionDb();
            DataTable dtRepons = new DataTable();
            dtRepons = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            sqlstr = @" select distinct nl.id as id_node, nl.node, nl.node as node_check, nw.responder_user_name
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        where h.seq = '" + seq + "' and nw.responder_user_name is not null ";
            sqlstr += @" order by nl.id";
            cls_conn = new ClassConnectionDb();
            DataTable dtNode = new DataTable();
            dtNode = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            sqlstr = @" select distinct d.document_no, d.document_file_name, nw.responder_user_name
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        where h.seq = '" + seq + "' and nw.responder_user_name is not null ";
            sqlstr += @" order by d.document_no";
            cls_conn = new ClassConnectionDb();
            DataTable dtDrawing = new DataTable();
            dtDrawing = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            sqlstr = @"select distinct
                        h.seq, h.pha_no, nl.id as id_node, g.pha_request_name
                        , nl.node, nl.node as node_check, nl.design_intent, nl.descriptions, nl.design_conditions, nl.node_boundary, nl.operating_conditions
                        , d.document_no, d.document_file_name
                        , mgw.guide_words as guideword, mgw.deviations as deviation, nw.causes, nw.consequences
                        , nw.category_type, nw.ram_befor_security, nw.ram_befor_likelihood, nw.ram_befor_risk
                        , nw.existing_safeguards, nw.recommendations, nw.recommendations_no, nw.responder_user_name, nw.responder_user_displayname
                        , nw.action_status
                        , nl.no as node_no, nw.no, nw.causes_no, nw.consequences_no
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word
                        where h.seq = '" + seq + "' and nw.responder_user_name is not null ";
            sqlstr += @" order by cast(nl.no as int),cast(nw.no as int), cast(nw.causes_no as int), cast(nw.consequences_no as int)";

            cls_conn = new ClassConnectionDb();
            DataTable dtWorksheet = new DataTable();
            dtWorksheet = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            sqlstr = @"select distinct 0 as ref, nl.node, nl.node as node_check, '' as ram_befor_risk, '' as recommendations, 0 as recommendations_no, '' as action_status, nw.responder_user_name, nw.responder_user_displayname 
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word
                        where h.seq = '" + seq + "' and nw.responder_user_name is not null ";
            sqlstr += @" order by nl.node, nw.responder_user_displayname ";
            cls_conn = new ClassConnectionDb();
            DataTable dtTrack = new DataTable();
            dtTrack = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
            if (true)
            {
                for (int t = 0; t < dtTrack.Rows.Count; t++)
                {
                    dtTrack.Rows[t]["ref"] = (t + 1);
                    dtTrack.AcceptChanges();
                }
            }


            FileInfo template = new FileInfo(_FolderTemplate + "HAZOP Recommendation Template.xlsx");
            if (report_all == true) { template = new FileInfo(_excel_name); }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {

                dt = new DataTable(); dt = dtRepons.Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    #region Response Sheet
                    if (true)
                    {
                        ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["RecommTemplate"];
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("RecommTemplate" + i, sourceWorksheet);

                        string ref_no = (i + 1).ToString();
                        worksheet.Name = "Response Sheet(Ref." + ref_no + ")";

                        string responder_user_name = (dt.Rows[i]["responder_user_name"] + "");
                        string responder_user_displayname = (dt.Rows[i]["responder_user_displayname"] + "");
                        string pha_request_name = (dt.Rows[i]["pha_request_name"] + "");
                        string pha_no = (dt.Rows[i]["pha_no"] + "");

                        int startRows = 2;
                        if (true)
                        {
                            string node = "";
                            string drawing_doc = "";
                            string deviation = "";
                            string causes = "";
                            string consequences = "";
                            string existing_safeguards = "";
                            string recommendations = "";
                            string recommendations_no = "";
                            int action_no = 0;

                            #region loop node 
                            DataRow[] drNode = dtNode.Select("responder_user_name = '" + responder_user_name + "' ");
                            for (int n = 0; n < drNode.Length; n++)
                            {
                                if (node != "") { node += ","; }
                                node += (drNode[n]["node"] + "");


                                for (int t = 0; t < dtTrack.Rows.Count; t++)
                                {
                                    if ((dtTrack.Rows[t]["responder_user_name"] + "") == responder_user_name
                                        && (dtTrack.Rows[t]["node_check"] + "") == (drNode[n]["node_check"] + ""))
                                    {
                                        dtTrack.Rows[t]["node"] = "Node: (" + ref_no + ")";
                                        dtTrack.AcceptChanges();
                                        break;
                                    }
                                }

                            }
                            #endregion loop node 

                            #region loop drawing_doc 
                            DataRow[] drDrawing = dtDrawing.Select("responder_user_name = '" + responder_user_name + "' ");
                            for (int n = 0; n < drDrawing.Length; n++)
                            {
                                if (drawing_doc != "") { drawing_doc += ","; }
                                drawing_doc += (drDrawing[n]["document_no"] + "");
                                if ((drDrawing[n]["document_file_name"] + "") != "")
                                {
                                    drawing_doc += " (" + drDrawing[n]["document_file_name"] + ")";
                                }
                            }
                            #endregion loop drawing_doc 

                            #region loop workksheet
                            DataRow[] drWorksheet = dtWorksheet.Select("responder_user_name = '" + responder_user_name + "' ");
                            for (int n = 0; n < drWorksheet.Length; n++)
                            {
                                if ((drWorksheet[n]["deviation"] + "") != "")
                                {
                                    if (deviation != "") { deviation += ","; }
                                    deviation += (drWorksheet[n]["guideword"] + "") + "/" + (drWorksheet[n]["deviation"] + "");
                                }
                                if ((drWorksheet[n]["causes"] + "") != "")
                                {
                                    if (causes != "") { causes += ","; }
                                    causes += (drWorksheet[n]["causes"] + "");
                                }
                                if ((drWorksheet[n]["consequences"] + "") != "")
                                {
                                    if (consequences != "") { consequences += ","; }
                                    consequences += (drWorksheet[n]["consequences"] + "");
                                }

                                if ((drWorksheet[n]["existing_safeguards"] + "") != "")
                                {
                                    if (existing_safeguards.IndexOf((drWorksheet[n]["existing_safeguards"] + "")) > -1) { }
                                    else
                                    {
                                        if (existing_safeguards != "") { existing_safeguards += ","; }
                                        existing_safeguards += (drWorksheet[n]["existing_safeguards"] + "");
                                    }
                                }

                                if ((drWorksheet[n]["recommendations"] + "") != "")
                                {
                                    if (recommendations != "") { recommendations += ","; }
                                    recommendations += (drWorksheet[n]["recommendations"] + "");
                                    action_no += 1;

                                    if (recommendations_no != "") { recommendations_no += ","; }
                                    recommendations_no += (drWorksheet[n]["recommendations_no"] + "");
                                }

                            }

                            #endregion loop workksheet

                            worksheet.Cells["A" + (startRows)].Value = "Project Title:" + pha_request_name;
                            startRows += 1;
                            worksheet.Cells["A" + (startRows)].Value = "Project No:" + pha_no;
                            startRows += 1;
                            worksheet.Cells["A" + (startRows)].Value = "Node:" + node;
                            startRows += 1;

                            worksheet.Cells["B" + (startRows)].Value = responder_user_displayname;
                            worksheet.Cells["E" + (startRows)].Value = responder_user_displayname;
                            startRows += 1;

                            worksheet.Cells["B" + (startRows)].Value = action_no;
                            startRows += 1;

                            worksheet.Cells["B" + (startRows)].Value = drawing_doc;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = deviation;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = causes;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = consequences;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = existing_safeguards;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = recommendations;
                            startRows += 1;
                        }

                    }
                    #endregion Response Sheet

                }

                #region TrackTemplate
                if (dtTrack.Rows.Count > 0)
                {
                    string recommendations = "";
                    string action_status = "";
                    string ram_befor_risk = "";
                    for (int t = 0; t < dtTrack.Rows.Count; t++)
                    {
                        DataRow[] drWorksheet = dtWorksheet.Select("responder_user_name = '" + dtTrack.Rows[t]["responder_user_name"] + "' "
                            + " and node = '" + dtTrack.Rows[t]["node_check"] + "'");
                        for (int n = 0; n < drWorksheet.Length; n++)
                        {
                            if (recommendations != "") { recommendations += ","; }
                            recommendations += (drWorksheet[n]["recommendations"] + "");

                            action_status = (drWorksheet[n]["action_status"] + "");
                            ram_befor_risk = (drWorksheet[n]["ram_befor_risk"] + "");
                        }
                        dtTrack.Rows[t]["recommendations"] = recommendations;
                        dtTrack.Rows[t]["action_status"] = action_status;
                        dtTrack.Rows[t]["ram_befor_risk"] = ram_befor_risk;
                        dtTrack.AcceptChanges();
                    }


                    if (true)
                    {
                        //ข้อมูลทั้งหมด
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TrackTemplate"];
                        worksheet.Name = "Status Tracking Table";

                        int i = 0;
                        int startRows = 3;

                        dt = new DataTable(); dt = dtTrack.Copy(); dt.AcceptChanges();
                        if (dt.Rows.Count > 0)
                        {
                            for (i = 0; i < dt.Rows.Count; i++)
                            {
                                worksheet.InsertRow(startRows, 1);
                                worksheet.Cells["A" + (startRows)].Value = dt.Rows[i]["ref"].ToString();
                                worksheet.Cells["B" + (startRows)].Value = dt.Rows[i]["node"].ToString();
                                worksheet.Cells["C" + (startRows)].Value = dt.Rows[i]["ram_befor_risk"].ToString();
                                worksheet.Cells["D" + (startRows)].Value = dt.Rows[i]["recommendations"].ToString();
                                worksheet.Cells["E" + (startRows)].Value = dt.Rows[i]["action_status"].ToString();
                                worksheet.Cells["F" + (startRows)].Value = dt.Rows[i]["responder_user_displayname"].ToString();
                                startRows++;
                            }

                            // วาดเส้นตาราง โดยใช้เซลล์ A1 ถึง C3
                            DrawTableBorders(worksheet, 3, 1, startRows - 1, 6);
                        }
                    }
                }
                #endregion Response Sheet



                if (true)
                {
                    dt = new DataTable(); dt = dtRepons.Copy(); dt.AcceptChanges();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string SheetName_befor = "Response Sheet(Ref." + (i).ToString() + ")";
                        string SheetName = "Response Sheet(Ref." + (i + 1).ToString() + ")";
                        if (i == 0) { SheetName_befor = "Status Tracking Table"; }

                        excelPackage.Workbook.Worksheets.MoveBefore(SheetName_befor, SheetName);
                    }
                }

                if (report_all == true)
                {
                    ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["RecommTemplate"];
                    SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                    excelPackage.Save();
                }
                else
                {
                    ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["RecommTemplate"];
                    SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                    excelPackage.SaveAs(new FileInfo(_Path + _excel_name));

                    // Save the workbook as PDF
                    if (export_type == "pdf")
                    {
                        Workbook workbookPDF = new Workbook(_Path + _excel_name);
                        PdfSaveOptions options = new PdfSaveOptions
                        {
                            AllColumnsInOnePagePerSheet = true
                        };
                        workbookPDF.Save(_Path + _excel_name.Replace(".xlsx", ".pdf"), options);
                        //return _DownloadPath + _excel_name.Replace(".xlsx", ".pdf");
                        if (true)
                        {
                            #region move file to _temp  
                            string export_file_name = _Path + _excel_name.Replace(".xlsx", ".pdf");
                            File.Copy(export_file_name, export_file_name.Replace(@"/Hazop/", @"/_temp/"));
                            try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                            try { File.Delete(export_file_name); } catch { }
                            #endregion move file to _temp
                        }
                        return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/Hazop/", @"/_temp/");

                    }
                }
            }


            //return _DownloadPath + _excel_name;  
            if (!report_all)
            {
                #region move file to _temp  
                string export_file_name = _Path + _excel_name;
                File.Copy(export_file_name, (export_file_name).Replace(@"/Hazop/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
                return (_DownloadPath + _excel_name).Replace(@"/Hazop/", @"/_temp/");
            }
            else { return (_DownloadPath + _excel_name); }
        }
        public string export_hazop_ram(ReportModel param)
        {
            string seq = param.seq;
            string export_type = param.export_type;

            DataTable dtdef = new DataTable();

            #region Determine whether the directory exists.
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.AcceptChanges();
            dtdef = dt.Clone(); dtdef.AcceptChanges();

            #endregion Determine whether the directory exists.

            string msg_error = "";
            string _DownloadPath = "/AttachedFileTemp/Hazop/";
            string _Folder = "/wwwroot/AttachedFileTemp/Hazop/";
            string _FolderTemplate = MapPathFiles("/wwwroot/AttachedFileTemp/");
            string _Path = MapPathFiles(_Folder);

            var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
            string export_file_name = "HAZOP RAM " + datetime_run;
            string export_file_name_full = "";
            if (export_type == "excel" || export_type == "pdf")
            {
                export_file_name_full = excel_hazop_ram(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name + ".xlsx", export_type, false, "Hazop");
                if (export_type == "excel") { export_file_name += ".xlsx"; } else { export_file_name += ".pdf"; }
            }

            try
            {
                dtdef.Rows.Add(dtdef.NewRow()); dtdef.AcceptChanges();
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_NAME"] = export_file_name;
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_PATH"] = export_file_name_full;
                dtdef.Rows[dtdef.Rows.Count - 1]["IMPORT_DATA_MSG"] = msg_error;
                dtdef.AcceptChanges();
            }
            catch (Exception ex) { ex.Message.ToString(); }

            return cls_json.SetJSONresult(dtdef);
        }
        public string excel_hazop_ram(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, Boolean report_all, string sub_software)
        {
            sqlstr = @"  select a.name as ram_type, a.descriptions, a.document_file_name
                        from epha_m_ram a where a.active_type = 1
                        and a.id in (select b.id_ram from epha_t_general b where b.id_pha = '" + seq + "'  )";
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];


            FileInfo template = new FileInfo(_FolderTemplate + "Risk Assessment Matrix Template.xlsx");
            if (report_all == true) { template = new FileInfo(_excel_name); }



            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Risk Assessment Matrix"];

                // Define picture dimensions and position
                int left = 2;     // column index
                int top = 2;      // row index
                int width = 400;  // width in pixels
                int height = 400; // height in pixels 

                // Define the picture file path
                string _FolderFile = MapPathFiles("/wwwroot/AttachedFileTemp/");
                string pictureFilePath = _FolderFile + (dt.Rows[0]["document_file_name"] + "");
                // Insert the picture
                var picture = worksheet.Drawings.AddPicture("RAM", new FileInfo(pictureFilePath));
                picture.From.Column = left;
                picture.From.Row = top;

                if ((dt.Rows[0]["ram_type"] + "") == "4x4") { width = 300; height = 300; }
                else if ((dt.Rows[0]["ram_type"] + "") == "5x5") { width = 400; height = 400; }
                else if ((dt.Rows[0]["ram_type"] + "") == "6x6") { width = 400; height = 400; }
                else if ((dt.Rows[0]["ram_type"] + "") == "7x7") { width = 400; height = 400; }
                else if ((dt.Rows[0]["ram_type"] + "") == "8x8") { width = 400; height = 400; }

                picture.SetSize(width, height);

                //descriptions  
                int startRows = 27;
                worksheet.Cells["A" + (startRows)].Value = dt.Rows[0]["descriptions"].ToString();

                if (report_all == true)
                {
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["Risk Assessment Matrix"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                    excelPackage.Save();
                }
                else
                {
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["Risk Assessment Matrix"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                    excelPackage.SaveAs(new FileInfo(_Path + _excel_name));

                    // Save the workbook as PDF
                    if (export_type == "pdf")
                    {
                        Workbook workbookPDF = new Workbook(_Path + _excel_name);
                        PdfSaveOptions options = new PdfSaveOptions
                        {
                            AllColumnsInOnePagePerSheet = true
                        };
                        workbookPDF.Save(_Path + _excel_name.Replace(".xlsx", ".pdf"), options);
                        //return _DownloadPath + _excel_name.Replace(".xlsx", ".pdf");
                        if (true)
                        {
                            #region move file to _temp  
                            string export_file_name = _Path + _excel_name.Replace(".xlsx", ".pdf");
                            File.Copy(export_file_name, export_file_name.Replace(@"/" + sub_software + @"/", @"/_temp/"));
                            try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                            try { File.Delete(export_file_name); } catch { }
                            #endregion move file to _temp
                        }
                        return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/" + sub_software + @"/", @"/_temp/");

                    }
                }
            }


            //return _DownloadPath + _excel_name;  
            if (!report_all)
            {
                #region move file to _temp  
                string export_file_name = _Path + _excel_name;
                File.Copy(export_file_name, (export_file_name).Replace(@"/" + sub_software + @"/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
                return (_DownloadPath + _excel_name).Replace(@"/" + sub_software + @"/", @"/_temp/");
            }
            else { return (_DownloadPath + _excel_name); }
        }
        public string export_hazop_guidewords(ReportModel param)
        {
            string seq = param.seq;
            string export_type = param.export_type;

            DataTable dtdef = new DataTable();

            #region Determine whether the directory exists.
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.AcceptChanges();
            dtdef = dt.Clone(); dtdef.AcceptChanges();

            #endregion Determine whether the directory exists.

            string msg_error = "";
            string _DownloadPath = "/AttachedFileTemp/Hazop/";
            string _Folder = "/wwwroot/AttachedFileTemp/Hazop/";
            string _FolderTemplate = MapPathFiles("/wwwroot/AttachedFileTemp/");
            string _Path = MapPathFiles(_Folder);

            var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
            string export_file_name = "HAZOP Guidewords " + datetime_run;
            string export_file_name_full = "";
            if (export_type == "excel" || export_type == "pdf")
            {
                export_file_name_full = excel_hazop_guidewords(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name + ".xlsx", export_type, false);
                if (export_type == "excel") { export_file_name += ".xlsx"; } else { export_file_name += ".pdf"; }
            }

            try
            {
                dtdef.Rows.Add(dtdef.NewRow()); dtdef.AcceptChanges();
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_NAME"] = export_file_name;
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_PATH"] = export_file_name_full;
                dtdef.Rows[dtdef.Rows.Count - 1]["IMPORT_DATA_MSG"] = msg_error;
                dtdef.AcceptChanges();
            }
            catch (Exception ex) { ex.Message.ToString(); }

            return cls_json.SetJSONresult(dtdef);
        }
        public string excel_hazop_guidewords(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, Boolean report_all)
        {
            sqlstr = @" select distinct parameter
                        from epha_m_guide_words where active_type = 1
                        order by parameter ";
            cls_conn = new ClassConnectionDb();
            DataTable dtParam = new DataTable();
            dtParam = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            sqlstr = @" select '' as usef_selected, def_selected, parameter, deviations, guide_words, process_deviation, area_application
                        from epha_m_guide_words where active_type = 1
                        order by parameter, deviations, guide_words, process_deviation, area_application ";
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];


            FileInfo template = new FileInfo(_FolderTemplate + "HAZOP Guidewords Template.xlsx");
            if (report_all == true) { template = new FileInfo(_excel_name); }




            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["GuidewordsTemplate"];
                ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("Guidewords", sourceWorksheet);
                worksheet.Name = "Guidewords";

                int startRows = 3;
                int i = 0;
                for (int m = 0; m < dtParam.Rows.Count; m++)
                {
                    string parameter = (dtParam.Rows[m]["parameter"] + "");
                    worksheet.InsertRow(startRows, 1);
                    var startCell = worksheet.Cells["A" + startRows];
                    var endCell = worksheet.Cells["D" + startRows];
                    var mergeRange = worksheet.Cells[startCell.Address + ":" + endCell.Address];
                    // Merge the cells
                    mergeRange.Merge = true;
                    // Optionally set text alignment in the merged cell
                    mergeRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    mergeRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["A" + (startRows)].Value = parameter;
                    startRows++;

                    DataRow[] dr = dt.Select("parameter = '" + parameter + "'");
                    for (i = 0; i < dr.Length; i++)
                    {
                        worksheet.InsertRow(startRows, 1);
                        worksheet.Cells["A" + (startRows)].Value = dr[i]["deviations"].ToString();
                        worksheet.Cells["B" + (startRows)].Value = dr[i]["guide_words"].ToString();
                        worksheet.Cells["C" + (startRows)].Value = dr[i]["process_deviation"].ToString();
                        worksheet.Cells["D" + (startRows)].Value = dr[i]["area_application"].ToString();
                        startRows++;
                    }
                }

                // วาดเส้นตาราง โดยใช้เซลล์ A1 ถึง D3
                DrawTableBorders(worksheet, 3, 1, startRows - 1, 4);

                if (report_all == true)
                {
                    //excelPackage.Workbook.Worksheets.MoveBefore("HAZOP Attendee Sheet", "HAZOP Cover Page"); 
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["GuidewordsTemplate"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                    excelPackage.Save();
                }
                else
                {
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["GuidewordsTemplate"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                    excelPackage.SaveAs(new FileInfo(_Path + _excel_name));

                    // Save the workbook as PDF
                    if (export_type == "pdf")
                    {
                        Workbook workbookPDF = new Workbook(_Path + _excel_name);
                        PdfSaveOptions options = new PdfSaveOptions
                        {
                            AllColumnsInOnePagePerSheet = true
                        };
                        workbookPDF.Save(_Path + _excel_name.Replace(".xlsx", ".pdf"), options);
                        //return _DownloadPath + _excel_name.Replace(".xlsx", ".pdf");
                        if (true)
                        {
                            #region move file to _temp  
                            string export_file_name = _Path + _excel_name.Replace(".xlsx", ".pdf");
                            File.Copy(export_file_name, export_file_name.Replace(@"/Hazop/", @"/_temp/"));
                            try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                            try { File.Delete(export_file_name); } catch { }
                            #endregion move file to _temp
                        }
                        return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/Hazop/", @"/_temp/");

                    }
                }
            }


            //return _DownloadPath + _excel_name;  
            if (!report_all)
            {
                #region move file to _temp  
                string export_file_name = _Path + _excel_name;
                File.Copy(export_file_name, (export_file_name).Replace(@"/Hazop/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
                return (_DownloadPath + _excel_name).Replace(@"/Hazop/", @"/_temp/");
            }
            else { return (_DownloadPath + _excel_name); }
        }

        public string export_template_jsea(ReportModel param)
        {
            string seq = param.seq;
            string export_type = param.export_type;
            string sub_software = param.sub_software;

            if (export_type == "template") { export_type = "excel"; }

            #region Determine whether the directory exists.
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.AcceptChanges();
            DataTable dtdef = dt.Clone(); dtdef.AcceptChanges();

            #endregion Determine whether the directory exists.

            string msg_error = "";
            string _DownloadPath = "/AttachedFileTemp/Jsea/";
            string _Folder = "/wwwroot/AttachedFileTemp/Jsea/";
            string _FolderTemplate = MapPathFiles("/wwwroot/AttachedFileTemp/");
            string _Path = MapPathFiles(_Folder);

            var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
            string export_file_name = "JSEA Report Template " + datetime_run;
            string export_file_name_full = "";
            if (export_type == "excel" || export_type == "pdf")
            {
                export_file_name_full = excle_template_data(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name + ".xlsx", export_type, sub_software);
                if (export_type == "excel") { export_file_name += ".xlsx"; } else { export_file_name += ".pdf"; }
            }
            else { }

            try
            {
                dtdef.Rows.Add(dtdef.NewRow()); dtdef.AcceptChanges();
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_NAME"] = export_file_name;
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_PATH"] = export_file_name_full;
                dtdef.Rows[dtdef.Rows.Count - 1]["IMPORT_DATA_MSG"] = msg_error;
                dtdef.AcceptChanges();
            }
            catch (Exception ex) { ex.Message.ToString(); }

            return cls_json.SetJSONresult(dtdef);
        }
        public string excle_template_data(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, string sub_software)
        {
            sqlstr = @" select distinct h.pha_no, g.pha_request_name, format(g.target_start_date,'dd MMM yyyy') as target_start_date
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        where h.seq = '" + seq + "'  ";
            sqlstr += @" order by g.pha_request_name";
            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            //JSEA Report Template.xlsx
            FileInfo template = new FileInfo(_FolderTemplate + "JSEA Report Template.xlsx");


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["WorksheetTemplate"];
                ExcelWorksheet worksheet = sourceWorksheet;
                try
                {
                    if (dt.Rows.Count > 0)
                    {
                        // c4
                        worksheet.Cells["C4"].Value = dt.Rows[0]["pha_request_name"].ToString();
                        // i4 = วันที่ทำการประเมิน (Date): 23/5/2566
                        worksheet.Cells["I4"].Value = "วันที่ทำการประเมิน (Date):" + dt.Rows[0]["target_start_date"].ToString();
                    }
                }
                catch { }
                try
                {
                    var startRows = 12;
                    var icol_end = 14;
                    for (int i = 0; i < 10; i++)
                    {
                        worksheet.InsertRow(startRows, 1);
                    }
                    DrawTableBorders(worksheet, startRows - 1, 2, startRows - 1, icol_end - 1);
                }
                catch { }
                excelPackage.SaveAs(new FileInfo(_Path + _excel_name));
            }

            return (_DownloadPath + _excel_name);
        }

        static void DrawTableBorders(ExcelWorksheet worksheet, int startRow, int startCol, int endRow, int endCol)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
            }
        }
        static void ClearTableBorders(ExcelWorksheet worksheet, int startRow, int startCol, int endRow, int endCol)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    cell.Style.Border.Top.Style = ExcelBorderStyle.None;
                    cell.Style.Border.Left.Style = ExcelBorderStyle.None;
                    cell.Style.Border.Right.Style = ExcelBorderStyle.None;
                    cell.Style.Border.Bottom.Style = ExcelBorderStyle.None;
                }
            }
        }
        public string copy_pdf_file(CopyFileModel param)
        {
            //page_start_first,page_start_second,page_end_first,page_end_second

            string file_name = param.file_name;
            string file_path = param.file_path;
            string page_start_first = (param.page_start_first == null ? "" : param.page_start_first).Replace("null", "");
            string page_start_second = (param.page_start_second == null ? "" : param.page_start_second).Replace("null", "");
            string page_end_first = (param.page_end_first == null ? "" : param.page_end_first).Replace("null", "");
            string page_end_second = (param.page_end_second == null ? "" : param.page_end_second).Replace("null", "");

            //D:\dotnet6-epha-api\dotnet6-epha-api\wwwroot\AttachedFileTemp\Hazop\ebook_def.pdf
            //file_name = "ebook_def.pdf";
            //page_start_first = "5";
            //page_start_second = "10";
            //page_end_first = "15";
            //page_end_second = "20";


            DataTable dtdef = new DataTable();
            #region Determine whether the directory exists.
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.AcceptChanges();
            dtdef = dt.Clone(); dtdef.AcceptChanges();

            #endregion Determine whether the directory exists.

            string msg_error = "";
            string _DownloadPath = "/AttachedFileTemp/Hazop/_copy/";
            string _Folder = MapPathFiles("/wwwroot/AttachedFileTemp/Hazop/_copy/");
            string _FolderTemplate = MapPathFiles("/wwwroot/AttachedFileTemp/Hazop/");

            var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
            string sourceFile = _FolderTemplate + file_name;
            string destinationFile = _Folder + file_name.Replace(".pdf", "").Replace(".PDF", "") + datetime_run + ".pdf";
            try
            {
                File.Copy(sourceFile, destinationFile, true);
            }
            catch { }
            string export_file_name = file_name.Replace(".pdf", "").Replace(".PDF", "") + datetime_run + ".pdf";
            string export_file_name_full = _DownloadPath + export_file_name;


            string sourcePdfPath = _FolderTemplate + file_name;// @"D:\dotnet6-epha-api\dotnet6-epha-api\wwwroot\AttachedFileTemp\Hazop\ebook_def.pdf";  // Replace with the path to the source PDF
            string targetPdfPath = _Folder + export_file_name;// @"D:\dotnet6-epha-api\dotnet6-epha-api\wwwroot\AttachedFileTemp\Hazop\ebook_v1.pdf"; // Replace with the path to the target PDF

            int startPagePart1 = (page_start_first == "" ? 1 : Convert.ToInt32(page_start_first));  // Replace with the start page number
            int endPagePart1 = (page_end_first == "" ? 100 : Convert.ToInt32(page_end_first)); ;    // Replace with the end page number
            int startPagePart2 = (page_start_second == "" ? 0 : Convert.ToInt32(page_start_second)); ;  // Replace with the start page number
            int endPagePart2 = (page_end_second == "" ? 0 : Convert.ToInt32(page_end_second)); ;    // Replace with the end page number

            try
            {

                using (var sourcePdfReader = new PdfReader(sourcePdfPath))
                using (var targetPdfStream = new FileStream(targetPdfPath, FileMode.Create))
                using (var targetPdfDoc = new iTextSharp.text.Document())
                using (var targetPdfWriter = new PdfCopy(targetPdfDoc, targetPdfStream))
                {
                    targetPdfDoc.Open();

                    if (startPagePart1 > 0)
                    {
                        for (int pageNumber = startPagePart1; pageNumber <= endPagePart1; pageNumber++)
                        {
                            var page = targetPdfWriter.GetImportedPage(sourcePdfReader, pageNumber);
                            targetPdfWriter.AddPage(page);
                        }
                    }
                    if (startPagePart2 > 0)
                    {
                        for (int pageNumber = startPagePart2; pageNumber <= endPagePart2; pageNumber++)
                        {
                            var page = targetPdfWriter.GetImportedPage(sourcePdfReader, pageNumber);
                            targetPdfWriter.AddPage(page);
                        }
                    }

                    targetPdfDoc.Close();
                    msg_error = "";

                }
            }
            catch (Exception ex) { }

            try
            {
                dtdef.Rows.Add(dtdef.NewRow()); dtdef.AcceptChanges();
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_NAME"] = export_file_name;
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_PATH"] = export_file_name_full;
                dtdef.Rows[dtdef.Rows.Count - 1]["IMPORT_DATA_MSG"] = msg_error;
                dtdef.AcceptChanges();
            }
            catch (Exception ex) { ex.Message.ToString(); }

            return cls_json.SetJSONresult(dtdef);
        }
        #endregion export excel hazop

        #region export excel jsea

        public string export_jsea_report(ReportModel param)
        {
            string seq = param.seq;
            string export_type = param.export_type;

            DataTable dtdef = new DataTable();

            #region Determine whether the directory exists.
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.AcceptChanges();
            dtdef = dt.Clone(); dtdef.AcceptChanges();

            #endregion Determine whether the directory exists.

            string msg_error = "";
            string _DownloadPath = "/AttachedFileTemp/Jsea/";
            string _Folder = "/wwwroot/AttachedFileTemp/Jsea/";
            string _FolderTemplate = MapPathFiles("/wwwroot/AttachedFileTemp/");
            string _Path = MapPathFiles(_Folder);

            var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
            string export_file_name = "JSEA Report " + datetime_run;
            string export_file_name_full = "";
            if (export_type == "excel" || export_type == "pdf")
            {
                export_file_name_full = excel_jsea_report(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name + ".xlsx", export_type);
                if (export_type == "excel") { export_file_name += ".xlsx"; } else { export_file_name += ".pdf"; }
            }

            try
            {
                dtdef.Rows.Add(dtdef.NewRow()); dtdef.AcceptChanges();
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_NAME"] = export_file_name;
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_PATH"] = export_file_name_full;
                dtdef.Rows[dtdef.Rows.Count - 1]["IMPORT_DATA_MSG"] = msg_error;
                dtdef.AcceptChanges();
            }
            catch (Exception ex) { ex.Message.ToString(); }

            return cls_json.SetJSONresult(dtdef);
        }
        public string excel_jsea_report(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type)
        {
            string export_file_name = _Path + _excel_name;
            FileInfo template = new FileInfo(_FolderTemplate + "JSEA Report Template.xlsx");
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                excelPackage.SaveAs(new FileInfo(export_file_name));
            }

            //Study Objective and Work Scope, Drawing & Reference, Node List
            excel_jsea_general(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name, export_type, true);

            //JSEA Attendee Sheet 
            excel_jsea_atendeesheet(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name, export_type, true);

            //JSEA Recommendation
            excel_jsea_recommendation(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name, export_type, true);

            excel_jsea_worksheet(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name, export_type, true);

            excel_hazop_ram(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name, export_type, true, "Jsea");

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(export_file_name))
            {
                string SheetName_befor = excelPackage.Workbook.Worksheets[excelPackage.Workbook.Worksheets.Count - 1].Name;
                string SheetName = "Drawing PIDs & PFDs";

                excelPackage.Workbook.Worksheets.MoveAfter(SheetName, SheetName_befor);

                // Save changes
                excelPackage.Save();
            }
            // Save the workbook as PDF
            if (export_type == "pdf")
            {
                Workbook workbookPDF = new Workbook(export_file_name);
                PdfSaveOptions options = new PdfSaveOptions
                {
                    AllColumnsInOnePagePerSheet = true
                };
                export_file_name = export_file_name.Replace(".xlsx", ".pdf");

                workbookPDF.Save(export_file_name, options);

                add_drawing_to_appendix(seq, _Path, export_file_name, true);


                if (true)
                {
                    #region move file to _temp  
                    if (File.Exists(export_file_name))
                    {
                        _delay_time(export_file_name);

                        File.Copy(export_file_name, (export_file_name.Replace(@"/Jsea/", @"/_temp/")), true); 
                        //File.Copy(export_file_name, _FolderTemplate + @"_temp/" +_excel_name.Replace(".xlsx", ".pdf"), true); 
                        try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                        try { File.Delete(export_file_name); } catch { }
                    }
                    #endregion move file to _temp
                }
                return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/Jsea/", @"/_temp/");

            }


            if (true)
            {
                #region move file to _temp  
                File.Copy(export_file_name, (export_file_name).Replace(@"/Jsea/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
            }
            return (_DownloadPath + _excel_name).Replace(@"/Jsea/", @"/_temp/");
        }
        private void _delay_time(string filePath)
        {

            int maxRetryAttempts = 5;
            int retryDelayMilliseconds = 1000;

            for (int i = 0; i < maxRetryAttempts; i++)
            {
                try
                {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        // Your code to work with the file
                    }
                    break; // If successful, break out of the loop
                }
                catch (IOException ex)
                {
                    // Handle the exception or log it
                    Console.WriteLine($"Attempt {i + 1}: {ex.Message}");
                    System.Threading.Thread.Sleep(retryDelayMilliseconds); // Wait before retrying
                }
            }
        }
        public string excel_jsea_general(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, Boolean report_all)
        {
            #region get data
            sqlstr = @" select g.work_scope
                         from EPHA_F_HEADER h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                         where h.seq = '" + seq + "' ";

            cls_conn = new ClassConnectionDb();
            DataTable dtWorkScope = new DataTable();
            dtWorkScope = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];


            sqlstr = @" select distinct d.no, d.document_name, d.document_no, d.document_file_name, d.descriptions 
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                        inner join EPHA_T_DRAWING d on h.id = d.id_pha    
                        where h.seq = '" + seq + "' and d.document_name is not null order by convert(int,d.no) ";

            cls_conn = new ClassConnectionDb();
            DataTable dtDrawing = new DataTable();
            dtDrawing = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            FileInfo template = new FileInfo(_FolderTemplate + "HAZOP Report Template.xlsx");
            if (report_all == true) { template = new FileInfo(_excel_name); }

            #endregion get data

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];  // Replace "SourceSheet" with the actual source sheet name
                ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("HAZOP Attendee Sheet", sourceWorksheet);

                //Study Objective and Work Scope
                worksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];
                worksheet.Cells["A2"].Value = (dtWorkScope.Rows[0]["work_scope"] + "");

                //Drawing & Reference
                #region Drawing & Reference
                if (true)
                {
                    worksheet = excelPackage.Workbook.Worksheets["Drawing & Reference"];

                    int startRows = 3;
                    int icol_end = 6;
                    int ino = 1;
                    for (int i = 0; i < dtDrawing.Rows.Count; i++)
                    {
                        //No.	Document Name	Drawing No	Document File	Comment
                        worksheet.InsertRow(startRows, 1);
                        worksheet.Cells["A" + (i + startRows)].Value = (i + 1); ;
                        worksheet.Cells["B" + (i + startRows)].Value = (dtDrawing.Rows[i]["document_name"] + "");
                        worksheet.Cells["C" + (i + startRows)].Value = (dtDrawing.Rows[i]["document_no"] + "");
                        worksheet.Cells["D" + (i + startRows)].Value = (dtDrawing.Rows[i]["document_file_name"] + "");
                        worksheet.Cells["E" + (i + startRows)].Value = (dtDrawing.Rows[i]["descriptions"] + "");
                        startRows++;
                    }
                    // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                    DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);

                    //var eRange = worksheet.Cells[worksheet.Cells["A3"].Address + ":" + worksheet.Cells["D" + startRows].Address];
                    //eRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //eRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
                #endregion Drawing & Reference



                //Study Objective and Work Scope
                #region Study Objective and Work Scope
                if (true)
                {
                    worksheet = excelPackage.Workbook.Worksheets["Study Objective and Work Scope"];
                    worksheet.Cells["A2"].Value = (dtWorkScope.Rows[0]["work_scope"] + "");
                }
                #endregion Study Objective and Work Scope

                if (report_all == true)
                {
                    //excelPackage.Workbook.Worksheets.MoveBefore("HAZOP Attendee Sheet", "HAZOP Cover Page"); 
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;

                    if (!Directory.Exists(_Path))
                    {
                        Directory.CreateDirectory(_Path);
                    }
                    excelPackage.Save();
                }
                else
                {
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;

                    excelPackage.SaveAs(new FileInfo(_Path + _excel_name));

                    // Save the workbook as PDF
                    if (export_type == "pdf")
                    {
                        Workbook workbookPDF = new Workbook(_Path + _excel_name);
                        PdfSaveOptions options = new PdfSaveOptions
                        {
                            AllColumnsInOnePagePerSheet = true
                        };
                        workbookPDF.Save(_Path + _excel_name.Replace(".xlsx", ".pdf"), options);
                        //return _DownloadPath + _excel_name.Replace(".xlsx", ".pdf");
                        if (true)
                        {
                            #region move file to _temp  
                            string export_file_name = _Path + _excel_name.Replace(".xlsx", ".pdf");
                            File.Copy(export_file_name, export_file_name.Replace(@"/Jsea/", @"/_temp/"));
                            try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                            try { File.Delete(export_file_name); } catch { }
                            #endregion move file to _temp
                        }
                        return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/Jsea/", @"/_temp/");
                    }
                }
            }

            if (!report_all)
            {
                #region move file to _temp  
                string export_file_name = _Path + _excel_name;
                File.Copy(export_file_name, (export_file_name).Replace(@"/Jsea/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
                return (_DownloadPath + _excel_name).Replace(@"/Jsea/", @"/_temp/");
            }
            else { return (_DownloadPath + _excel_name); }
        }
        public string excel_jsea_atendeesheet(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, Boolean report_all)
        {
            //ตอนนี้รายละเอียดจะเหมือนกับ HAZOP แค่แยกออกมาก่อน รอ comment จาก user อีกที
            sqlstr = @" select s.id_pha, s.seq as seq_session, s.no as session_no
                         , convert(varchar,s.meeting_date,106) as meeting_date
                         , mt.no as member_no, isnull(mt.user_name,'') as user_name, emp.user_displayname
                         from EPHA_F_HEADER h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                         inner join EPHA_T_SESSION s on h.id = s.id_pha 
                         left join EPHA_T_MEMBER_TEAM mt on h.id = mt. id_pha and mt.id_session = s.id
                         left join VW_EPHA_PERSON_DETAILS emp on lower(emp.user_name) = lower(mt.user_name)
                         where h.seq = '" + seq + "' and lower(mt.user_name) is not null ";

            cls_conn = new ClassConnectionDb();
            DataTable dtAll = new DataTable();
            dtAll = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            cls_conn = new ClassConnectionDb();
            DataTable dtMember = new DataTable();
            dtMember = cls_conn.ExecuteAdapterSQL(" select distinct 0 as no, t.user_name, t.user_displayname, '' as company_text from (" + sqlstr + " )t where t.user_name <> '' order by t.user_name").Tables[0];

            cls_conn = new ClassConnectionDb();
            DataTable dtSession = new DataTable();
            dtSession = cls_conn.ExecuteAdapterSQL(" select distinct t.seq_session, t.session_no, t.meeting_date from (" + sqlstr + ")t order by t.session_no ").Tables[0];

            Boolean bCheckNewFile = false;
            FileInfo template = new FileInfo(_FolderTemplate + "JSEA AttendeeSheet Template.xlsx");
            if (report_all == true) { template = new FileInfo(_excel_name); }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];  // Replace "SourceSheet" with the actual source sheet name
                sourceWorksheet.Name = "JSEA Attendee Sheet";
                ExcelWorksheet worksheet = sourceWorksheet;// excelPackage.Workbook.Worksheets.Add("HAZOP Attendee Sheet", sourceWorksheet);

                int i = 0;
                int startRows = 4;
                int icol_start = 4;
                int icol_end = icol_start + (dtSession.Rows.Count > 6 ? dtSession.Rows.Count : 6);

                for (int imember = 0; imember < dtMember.Rows.Count; imember++)
                {
                    worksheet.InsertRow(startRows, 1);
                    string user_name = (dtMember.Rows[imember]["user_name"] + "");
                    //No.
                    worksheet.Cells["A" + (i + startRows)].Value = (imember + 1);
                    //Name
                    worksheet.Cells["B" + (i + startRows)].Value = (dtMember.Rows[imember]["user_displayname"] + "");
                    //Company
                    worksheet.Cells["C" + (i + startRows)].Value = (dtMember.Rows[imember]["company_text"] + "");

                    int irow_session = 0;
                    if (imember == 0)
                    {
                        if (dtSession.Rows.Count < 6)
                        {
                            //worksheet.Cells[2, icol_start, 2, icol_end].Merge = true; 
                            for (int c = icol_end; c < 30; c++)
                            {
                                worksheet.DeleteColumn(icol_end);

                            }
                        }

                        irow_session = 0;
                        for (int c = icol_start; c < icol_end; c++)
                        {
                            try
                            {
                                //header 
                                if ((dtSession.Rows[irow_session]["meeting_date"] + "") == "")
                                {
                                    worksheet.Cells[3, c].Value = "";
                                }
                                else
                                {
                                    worksheet.Cells[3, c].Value = (dtSession.Rows[irow_session]["meeting_date"] + "");
                                }
                            }
                            catch { worksheet.Cells[3, c].Value = ""; }
                            irow_session += 1;
                        }
                    }

                    irow_session = 0;
                    for (int c = icol_start; c < icol_end; c++)
                    {
                        try
                        {
                            string session_no = "";
                            try { session_no = (dtSession.Rows[irow_session]["session_no"] + ""); } catch { }

                            DataRow[] dr = dtAll.Select("user_name = '" + user_name + "' and session_no = '" + session_no + "'");
                            if (dr.Length > 0)
                            {
                                worksheet.Cells[startRows, c].Value = "X";
                            }
                            else { worksheet.Cells[startRows, c].Value = ""; }
                        }
                        catch { }
                        irow_session++;

                    }

                    startRows++;
                }

                // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);

                if (report_all == true)
                {
                    //excelPackage.Workbook.Worksheets.MoveBefore("HAZOP Attendee Sheet", "Study Objective and Work Scope"); 
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;

                    if (!Directory.Exists(_Path))
                    {
                        Directory.CreateDirectory(_Path);
                    }
                    excelPackage.Save();
                }
                else
                {
                    //ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["AttendeeSheetTemplate"];
                    //SheetTemplate.Hidden = eWorkSheetHidden.Hidden;

                    excelPackage.SaveAs(new FileInfo(_Path + _excel_name));

                    // Save the workbook as PDF
                    if (export_type == "pdf")
                    {
                        Workbook workbookPDF = new Workbook(_Path + _excel_name);
                        PdfSaveOptions options = new PdfSaveOptions
                        {
                            AllColumnsInOnePagePerSheet = true
                        };
                        workbookPDF.Save(_Path + _excel_name.Replace(".xlsx", ".pdf"), options);
                        //return _DownloadPath + _excel_name.Replace(".xlsx", ".pdf");
                        if (true)
                        {
                            #region move file to _temp  
                            string export_file_name = _Path + _excel_name.Replace(".xlsx", ".pdf");
                            File.Copy(export_file_name, export_file_name.Replace(@"/Jsea/", @"/_temp/"));
                            try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                            try { File.Delete(export_file_name); } catch { }
                            #endregion move file to _temp
                        }
                        return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/Jsea/", @"/_temp/");
                    }
                }
            }

            //return _DownloadPath + _excel_name;  
            if (!report_all)
            {
                #region move file to _temp  
                string export_file_name = _Path + _excel_name;
                File.Copy(export_file_name, (export_file_name).Replace(@"/Jsea/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
                return (_DownloadPath + _excel_name).Replace(@"/Jsea/", @"/_temp/");
            }
            else { return (_DownloadPath + _excel_name); }
        }
        public string excel_jsea_recommendation(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, Boolean report_all)
        {
            sqlstr = @" select distinct
                        h.seq, h.pha_no, nl.id as id_node, g.pha_request_name
                        , nl.node, nl.node as node_check, nl.design_intent, nl.descriptions, nl.design_conditions, nl.node_boundary, nl.operating_conditions
                        , d.document_no, d.document_file_name
                        , mgw.guide_words as guideword, mgw.deviations as deviation, nw.causes, nw.consequences
                        , nw.category_type, nw.ram_befor_security, nw.ram_befor_likelihood, nw.ram_befor_risk
                        , nw.existing_safeguards, nw.recommendations, nw.recommendations_no, nw.responder_user_name, nw.responder_user_displayname
                        , nw.action_status
                        , nl.no as node_no, nw.no, nw.causes_no, nw.consequences_no
                        , nw.seq as seq_worksheet
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word
                        where h.seq = '" + seq + "' and nw.responder_user_name is not null ";
            sqlstr += @" order by cast(nl.no as int),cast(nw.no as int), cast(nw.causes_no as int), cast(nw.consequences_no as int)";

            cls_conn = new ClassConnectionDb();
            DataTable dtWorksheet = new DataTable();
            dtWorksheet = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            sqlstr = @" select distinct nl.no, nw.no, nw.seq, 0 as ref, nl.node, nl.node as node_check
                        , nw.ram_after_risk, nw.ram_after_risk_action, nw.recommendations, nw.recommendations_no, nw.action_status, nw.responder_user_name, nw.responder_user_displayname 
                        from EPHA_F_HEADER h 
                        inner join EPHA_T_GENERAL g on h.id = g.id_pha 
                        inner join EPHA_T_NODE nl on h.id = nl.id_pha 
                        left join EPHA_T_NODE_DRAWING nd on h.id = nd.id_pha and  nl.id = nd.id_node 
                        left join EPHA_T_DRAWING d on h.id = d.id_pha and  nd.id_drawing = d.id
                        left join EPHA_T_NODE_WORKSHEET nw on h.id = nw.id_pha and  nl.id = nw.id_node   
                        left join EPHA_M_GUIDE_WORDS mgw on mgw.id = nw.id_guide_word
                        where h.seq = '" + seq + "' and nw.responder_user_name is not null ";
            sqlstr += @" order by nl.no, nw.no ";
            cls_conn = new ClassConnectionDb();
            DataTable dtTrack = new DataTable();
            dtTrack = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
            if (true)
            {
                for (int t = 0; t < dtTrack.Rows.Count; t++)
                {
                    dtTrack.Rows[t]["ref"] = (t + 1);
                    dtTrack.AcceptChanges();
                }
            }


            FileInfo template = new FileInfo(_FolderTemplate + "JSEA Recommendation Template.xlsx");
            if (report_all == true) { template = new FileInfo(_excel_name); }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {

                dt = new DataTable(); dt = dtWorksheet.Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    #region Sheet
                    if (true)
                    {
                        ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["RecommTemplate"];
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("RecommTemplate" + i, sourceWorksheet);

                        string ref_no = (i + 1).ToString();
                        worksheet.Name = "Response Sheet(Ref." + ref_no + ")";

                        string responder_user_name = (dt.Rows[i]["responder_user_name"] + "");
                        string responder_user_displayname = (dt.Rows[i]["responder_user_displayname"] + "");
                        string pha_request_name = (dt.Rows[i]["pha_request_name"] + "");
                        string pha_no = (dt.Rows[i]["pha_no"] + "");
                        string seq_worksheet = (dt.Rows[i]["seq_worksheet"] + "");


                        int startRows = 2;
                        if (true)
                        {
                            string tasks = "";
                            string drawing_doc = "";

                            //workstep,taskdesc,potentailhazard,possiblecase 
                            string workstep = "";
                            string taskdesc = "";
                            string potentailhazard = "";
                            string possiblecase = "";

                            string recommendations = "";
                            string recommendations_no = "";

                            int action_no = 0;

                            #region loop drawing_doc 
                            drawing_doc = (dt.Rows[i]["document_no"] + "");
                            if ((dt.Rows[i]["document_file_name"] + "") != "")
                            {
                                drawing_doc += " (" + dt.Rows[i]["document_file_name"] + ")";
                            }
                            #endregion loop drawing_doc 

                            #region loop workksheet
                            DataRow[] drWorksheet = dt.Select("seq_worksheet = '" + seq_worksheet + "'");
                            for (int n = 0; n < drWorksheet.Length; n++)
                            {
                                //workstep,taskdesc,potentailhazard,possiblecase 
                                if ((drWorksheet[n]["workstep"] + "") != "")
                                {
                                    if (workstep != "") { workstep += ","; }
                                    workstep += (drWorksheet[n]["workstep"] + "");
                                }

                                if ((drWorksheet[n]["taskdesc"] + "") != "")
                                {
                                    if (taskdesc != "") { taskdesc += ","; }
                                    taskdesc += (drWorksheet[n]["taskdesc"] + "");
                                }
                                if ((drWorksheet[n]["potentailhazard"] + "") != "")
                                {
                                    if (potentailhazard != "") { potentailhazard += ","; }
                                    potentailhazard += (drWorksheet[n]["potentailhazard"] + "");
                                }
                                if ((drWorksheet[n]["possiblecase"] + "") != "")
                                {
                                    if (possiblecase != "") { possiblecase += ","; }
                                    possiblecase += (drWorksheet[n]["possiblecase"] + "");
                                }

                                if ((drWorksheet[n]["recommendations"] + "") != "")
                                {
                                    if (recommendations != "") { recommendations += ","; }
                                    recommendations += (drWorksheet[n]["recommendations"] + "");
                                    action_no += 1;

                                    if (recommendations_no != "") { recommendations_no += ","; }
                                    recommendations_no += (drWorksheet[n]["recommendations_no"] + "");
                                }

                            }

                            #endregion loop workksheet

                            worksheet.Cells["A" + (startRows)].Value = "Project Title:" + pha_request_name;
                            startRows += 1;
                            worksheet.Cells["A" + (startRows)].Value = "Project No:" + pha_no;
                            startRows += 1;
                            worksheet.Cells["A" + (startRows)].Value = "Tasks:" + tasks;
                            startRows += 1;

                            worksheet.Cells["B" + (startRows)].Value = responder_user_displayname;
                            worksheet.Cells["E" + (startRows)].Value = responder_user_displayname;
                            startRows += 1;

                            worksheet.Cells["B" + (startRows)].Value = action_no;
                            startRows += 1;

                            worksheet.Cells["B" + (startRows)].Value = drawing_doc;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = workstep;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = taskdesc;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = potentailhazard;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = possiblecase;
                            startRows += 1;
                            worksheet.Cells["B" + (startRows)].Value = recommendations;
                            startRows += 1;
                        }

                    }
                    #endregion Sheet

                }

                #region TrackTemplate
                if (dtTrack.Rows.Count > 0)
                {
                    //ข้อมูลทั้งหมด
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["TrackTemplate"];
                    worksheet.Name = "Status Tracking Table";

                    int i = 0;
                    int startRows = 3;

                    dt = new DataTable(); dt = dtTrack.Copy(); dt.AcceptChanges();
                    if (dt.Rows.Count > 0)
                    {
                        for (i = 0; i < dt.Rows.Count; i++)
                        {
                            worksheet.InsertRow(startRows, 1);
                            worksheet.Cells["A" + (startRows)].Value = dt.Rows[i]["ref"].ToString();
                            worksheet.Cells["B" + (startRows)].Value = dt.Rows[i]["workstep"].ToString();
                            worksheet.Cells["C" + (startRows)].Value = dt.Rows[i]["ram_after_risk"].ToString();
                            worksheet.Cells["D" + (startRows)].Value = dt.Rows[i]["recommendations"].ToString();
                            worksheet.Cells["E" + (startRows)].Value = dt.Rows[i]["action_status"].ToString();
                            worksheet.Cells["F" + (startRows)].Value = dt.Rows[i]["responder_user_displayname"].ToString();
                            startRows++;
                        }

                        // วาดเส้นตาราง โดยใช้เซลล์ A1 ถึง C3
                        DrawTableBorders(worksheet, 3, 1, startRows - 1, 6);
                    }
                }
                #endregion Response Sheet

                if (report_all == true)
                {
                    ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["RecommTemplate"];
                    SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                    excelPackage.Save();
                }
                else
                {
                    ExcelWorksheet SheetTemplate = excelPackage.Workbook.Worksheets["RecommTemplate"];
                    SheetTemplate.Hidden = eWorkSheetHidden.Hidden;
                    excelPackage.SaveAs(new FileInfo(_Path + _excel_name));

                    // Save the workbook as PDF
                    if (export_type == "pdf")
                    {
                        Workbook workbookPDF = new Workbook(_Path + _excel_name);
                        PdfSaveOptions options = new PdfSaveOptions
                        {
                            AllColumnsInOnePagePerSheet = true
                        };
                        workbookPDF.Save(_Path + _excel_name.Replace(".xlsx", ".pdf"), options);
                        //return _DownloadPath + _excel_name.Replace(".xlsx", ".pdf");
                        if (true)
                        {
                            #region move file to _temp  
                            string export_file_name = _Path + _excel_name.Replace(".xlsx", ".pdf");
                            File.Copy(export_file_name, export_file_name.Replace(@"/Jsea/", @"/_temp/"));
                            try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                            try { File.Delete(export_file_name); } catch { }
                            #endregion move file to _temp
                        }
                        return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/Jsea/", @"/_temp/");

                    }
                }
            }


            //return _DownloadPath + _excel_name;  
            if (!report_all)
            {
                #region move file to _temp  
                string export_file_name = _Path + _excel_name;
                File.Copy(export_file_name, (export_file_name).Replace(@"/Jsea/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
                return (_DownloadPath + _excel_name).Replace(@"/Jsea/", @"/_temp/");
            }
            else { return (_DownloadPath + _excel_name); }
        }

        public string excel_jsea_worksheet(string seq, string _Path, string _FolderTemplate, string _DownloadPath, string _excel_name, string export_type, Boolean report_all)
        {
            #region get data
            sqlstr = @"  select g.pha_request_name, format(g.target_start_date,'dd MMM yyyy') as target_start_date, g.descriptions
                         from EPHA_F_HEADER h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                         where h.seq = '" + seq + "' ";

            cls_conn = new ClassConnectionDb();
            DataTable dtHead = new DataTable();
            dtHead = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            sqlstr = @"   select tw.no, tw.workstep_no, tw.workstep, tw.taskdesc_no, tw.taskdesc, tw.potentailhazard_no, tw.potentailhazard, tw.possiblecase_no, tw.possiblecase  
                         , tw.category_no, tw.category_type, tw.ram_befor_security, tw.ram_befor_likelihood, tw.ram_befor_risk, tw.recommendations, tw.responder_action_by
                         , tw.ram_after_security, tw.ram_after_likelihood, tw.ram_after_risk
                         , g.id_ram
                         from EPHA_F_HEADER h 
                         inner join EPHA_T_GENERAL g on h.id = g.id_pha  
                         inner join EPHA_T_TASKS_WORKSHEET tw on h.id  = tw.id_pha 
                         where h.seq = '" + seq + "' order by tw.no  ";

            cls_conn = new ClassConnectionDb();
            DataTable dtWorksheet = new DataTable();
            dtWorksheet = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            sqlstr = @" select tr.user_type ,tr.no, tr.user_displayname, tr.user_title, tr.reviewer_date
                         from EPHA_F_HEADER h
                         inner join EPHA_T_TASKS_RELATEDPEOPLE tr on h.id  = tr.id_pha
                         where h.seq = '" + seq + "' order by tr.user_type ,tr.no  ";

            cls_conn = new ClassConnectionDb();
            DataTable dtRelatedPeople = new DataTable();
            dtRelatedPeople = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];


            FileInfo template = new FileInfo(_FolderTemplate + "JSEA Report Template.xlsx");
            if (report_all == true) { template = new FileInfo(_excel_name); }

            #endregion get data

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(template))
            {
                ExcelWorksheet sourceWorksheet = excelPackage.Workbook.Worksheets["WorksheetTemplate"];
                ExcelWorksheet worksheet = sourceWorksheet;

                //Worksheet
                #region Worksheet
                if (true)
                {
                    //header
                    worksheet.Cells["C4"].Value = (dtHead.Rows[0]["pha_request_name"] + "");
                    worksheet.Cells["I4"].Value = "วันที่ทำการประเมิน (Date): " + (dtHead.Rows[0]["target_start_date"] + "");
                    worksheet.Cells["I5"].Value = (dtHead.Rows[0]["descriptions"] + "");

                    int startRows = 12;
                    int icol_start = 1;
                    int icol_end = 6;
                    for (int i = 0; i < dtWorksheet.Rows.Count; i++)
                    {
                        worksheet.InsertRow(startRows, 1);
                        icol_start = 1;

                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (dtWorksheet.Rows[i]["workstep_no"] + "." + dtWorksheet.Rows[i]["workstep"]);
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (dtWorksheet.Rows[i]["taskdesc_no"] + "." + dtWorksheet.Rows[i]["taskdesc"]);
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (dtWorksheet.Rows[i]["potentailhazard_no"] + "." + dtWorksheet.Rows[i]["potentailhazard"]);
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (dtWorksheet.Rows[i]["possiblecase_no"] + "." + dtWorksheet.Rows[i]["possiblecase"]);
                        if ((dtWorksheet.Rows[i]["id_ram"] + "") == "5")
                        {
                            icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (dtWorksheet.Rows[i]["category_type"] + "");
                        }
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (dtWorksheet.Rows[i]["ram_befor_security"] + "");
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (dtWorksheet.Rows[i]["ram_befor_likelihood"] + "");
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (dtWorksheet.Rows[i]["ram_befor_risk"] + "");
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (dtWorksheet.Rows[i]["recommendations"] + "");
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (dtWorksheet.Rows[i]["responder_action_by"] + "");
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (dtWorksheet.Rows[i]["ram_after_security"] + "");
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (dtWorksheet.Rows[i]["ram_after_likelihood"] + "");
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (dtWorksheet.Rows[i]["ram_after_risk"] + "");

                        icol_end = icol_start;

                        // วาดเส้นตาราง โดยใช้เซลล์ XX ถึง XX
                        DrawTableBorders(worksheet, 1, 1, startRows - 1, icol_end - 1);


                        startRows++;
                    }

                    //RelatedPeople 
                    //attendees,specialist,reviewer,approver
                    int startRowsRP = startRows;
                    int endColRP = icol_end;
                    int iapprover_start_row = startRowsRP + 7;

                    DataRow[] drAttendees = dtRelatedPeople.Select("user_type = 'attendees'");
                    DataRow[] drSpecialist = dtRelatedPeople.Select("user_type = 'specialist'");
                    DataRow[] drReviewer = dtRelatedPeople.Select("user_type = 'reviewer'");
                    DataRow[] drApprover = dtRelatedPeople.Select("user_type = 'approver'");

                    //default row running  = 7 row
                    int iDefRow = 6;
                    for (int i = 0; i < drAttendees.Length; i++)
                    {
                        if (i >= iDefRow)
                        {
                            worksheet.InsertRow(startRows, 1);

                            iapprover_start_row = startRowsRP + 8;
                            ClearTableBorders(worksheet, iapprover_start_row, endColRP - 4, iapprover_start_row, endColRP - 3);
                        }
                        icol_start = 1;
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (drAttendees[i]["no"] + "");
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (drAttendees[i]["user_displayname"] + "");
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (drAttendees[i]["user_title"] + "");

                        startRows++;
                    }
                    //default row running  = 4 row
                    iDefRow = 3;
                    for (int i = 0; i < drSpecialist.Length; i++)
                    {
                        if (i >= iDefRow) { worksheet.InsertRow(startRows, 1); }
                        icol_start = 1;
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (drAttendees[i]["no"] + "");
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (drAttendees[i]["user_displayname"] + "");
                        icol_start += 1; worksheet.Cells[(i + startRows), icol_start].Value = (drAttendees[i]["user_title"] + "");

                        startRows++;
                    }

                    if (drReviewer.Length > 0)
                    {
                        worksheet.Cells[startRowsRP + 1, endColRP - 4].Value = (drReviewer[0]["user_displayname"] + "");
                        worksheet.Cells[startRowsRP + 1, endColRP - 3].Value = (drReviewer[0]["reviewer_date"] + "");
                    }

                    if (drApprover.Length > 0)
                    {
                        worksheet.Cells[iapprover_start_row, endColRP - 5].Value = ("AE or AGSI");
                        worksheet.Cells[iapprover_start_row, endColRP - 4].Value = (drApprover[0]["user_displayname"] + "");
                        worksheet.Cells[iapprover_start_row, endColRP - 3].Value = (drApprover[0]["reviewer_date"] + "");

                        DrawTableBorders(worksheet, iapprover_start_row, endColRP - 4, iapprover_start_row, endColRP - 3);
                        DrawTableBorders(worksheet, iapprover_start_row + 2, endColRP - 4, iapprover_start_row + 2, endColRP - 4);
                    }

                }
                #endregion Worksheet


                if (report_all == true)
                {
                    if (!Directory.Exists(_Path))
                    {
                        Directory.CreateDirectory(_Path);
                    }
                    excelPackage.Save();
                }
                else
                {
                    excelPackage.SaveAs(new FileInfo(_Path + _excel_name));

                    // Save the workbook as PDF
                    if (export_type == "pdf")
                    {
                        Workbook workbookPDF = new Workbook(_Path + _excel_name);
                        PdfSaveOptions options = new PdfSaveOptions
                        {
                            AllColumnsInOnePagePerSheet = true
                        };
                        workbookPDF.Save(_Path + _excel_name.Replace(".xlsx", ".pdf"), options);
                        //return _DownloadPath + _excel_name.Replace(".xlsx", ".pdf");
                        if (true)
                        {
                            #region move file to _temp  
                            string export_file_name = _Path + _excel_name.Replace(".xlsx", ".pdf");
                            File.Copy(export_file_name, export_file_name.Replace(@"/Jsea/", @"/_temp/"));
                            try { File.Delete(export_file_name.Replace(".pdf", ".xlsx")); } catch { }
                            try { File.Delete(export_file_name); } catch { }
                            #endregion move file to _temp
                        }
                        return (_DownloadPath + _excel_name.Replace(".xlsx", ".pdf")).Replace(@"/Jsea/", @"/_temp/");
                    }
                }
            }

            if (!report_all)
            {
                #region move file to _temp  
                string export_file_name = _Path + _excel_name;
                File.Copy(export_file_name, (export_file_name).Replace(@"/jsea/", @"/_temp/"));
                try { File.Delete(export_file_name); } catch { }
                #endregion move file to _temp
                return (_DownloadPath + _excel_name).Replace(@"/jsea/", @"/_temp/");
            }
            else { return (_DownloadPath + _excel_name); }
        }

        public string export_jsea_worksheet(ReportModel param)
        {
            string seq = param.seq;
            string export_type = param.export_type;

            DataTable dtdef = new DataTable();

            #region Determine whether the directory exists.
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.AcceptChanges();
            dtdef = dt.Clone(); dtdef.AcceptChanges();

            #endregion Determine whether the directory exists.

            string msg_error = "";
            string _DownloadPath = "/AttachedFileTemp/Jsea/";
            string _Folder = "/wwwroot/AttachedFileTemp/Jsea/";
            string _FolderTemplate = MapPathFiles("/wwwroot/AttachedFileTemp/");
            string _Path = MapPathFiles(_Folder);

            var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
            string export_file_name = "JSEA WORKSHEET & RECOMMENDATION RESPONSE SHEET " + datetime_run;
            string export_file_name_full = "";
            if (export_type == "excel" || export_type == "pdf")
            {
                export_file_name_full = excel_jsea_worksheet(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name + ".xlsx", export_type, false);
                if (export_type == "excel") { export_file_name += ".xlsx"; } else { export_file_name += ".pdf"; }

            }

            try
            {
                dtdef.Rows.Add(dtdef.NewRow()); dtdef.AcceptChanges();
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_NAME"] = export_file_name;
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_PATH"] = export_file_name_full;
                dtdef.Rows[dtdef.Rows.Count - 1]["IMPORT_DATA_MSG"] = msg_error;
                dtdef.AcceptChanges();
            }
            catch (Exception ex) { ex.Message.ToString(); }

            return cls_json.SetJSONresult(dtdef);
        }
        public string export_jsea_recommendation(ReportModel param)
        {
            string seq = param.seq;
            string export_type = param.export_type;

            DataTable dtdef = new DataTable();

            #region Determine whether the directory exists.
            DataTable dt = new DataTable();
            dt.Columns.Add("ATTACHED_FILE_NAME");
            dt.Columns.Add("ATTACHED_FILE_PATH");
            dt.Columns.Add("ATTACHED_FILE_OF");
            dt.Columns.Add("IMPORT_DATA_MSG");
            dt.AcceptChanges();
            dtdef = dt.Clone(); dtdef.AcceptChanges();

            #endregion Determine whether the directory exists.

            string msg_error = "";
            string _DownloadPath = "/AttachedFileTemp/Jsea/";
            string _Folder = "/wwwroot/AttachedFileTemp/Jsea/";
            string _FolderTemplate = MapPathFiles("/wwwroot/AttachedFileTemp/");
            string _Path = MapPathFiles(_Folder);

            var datetime_run = DateTime.Now.ToString("yyyyMMddHHmm");
            string export_file_name = "JSEA RECOMMENDATION RESPONSE SHEET & RECCOMENDATION STATUS TRACKING TABLE " + datetime_run;
            string export_file_name_full = "";
            if (export_type == "excel" || export_type == "pdf")
            {
                export_file_name_full = excel_hazop_recommendation(seq, _Path, _FolderTemplate, _DownloadPath, export_file_name + ".xlsx", export_type, false);
                if (export_type == "excel") { export_file_name += ".xlsx"; } else { export_file_name += ".pdf"; }
            }

            try
            {
                dtdef.Rows.Add(dtdef.NewRow()); dtdef.AcceptChanges();
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_NAME"] = export_file_name;
                dtdef.Rows[dtdef.Rows.Count - 1]["ATTACHED_FILE_PATH"] = export_file_name_full;
                dtdef.Rows[dtdef.Rows.Count - 1]["IMPORT_DATA_MSG"] = msg_error;
                dtdef.AcceptChanges();
            }
            catch (Exception ex) { ex.Message.ToString(); }

            return cls_json.SetJSONresult(dtdef);
        }

        #endregion export excel jsea

        #region export doc

        #endregion export doc

        #region save data
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

            string sub_expense_type = "";
            try
            {
                sub_expense_type = (dsData.Tables["general"].Rows[0]["sub_expense_type"] + "");
            }
            catch { }

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
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }


            jsper = param.json_session + "";
            if (jsper.Trim() == "") { msg = "No Data."; ret = "Error"; }
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
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

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
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

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
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

            jsper = param.json_ram_level + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "ram_level";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

            jsper = param.json_ram_master + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "ram_master";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

            jsper = param.json_flow_action + "";
            try
            {
                dt = new DataTable();
                dt = cls_json.ConvertJSONresult(jsper);
                if (dt != null)
                {
                    dt.TableName = "flow_action";
                    dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                }
            }
            catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; }

            //hazop
            if (true)
            {
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


            }

            //jsea
            if (true)
            {

                jsper = param.json_tasks_worksheet + "";
                try
                {
                    if (jsper.Trim() != "")
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "tasks_worksheet";
                            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                        }
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

                jsper = param.json_tasks_relatedpeople + "";
                try
                {
                    if (jsper.Trim() != "")
                    {
                        dt = new DataTable();
                        dt = cls_json.ConvertJSONresult(jsper);
                        if (dt != null)
                        {
                            dt.TableName = "tasks_relatedpeople";
                            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
                        }
                    }
                }
                catch (Exception ex) { msg = ex.Message.ToString() + ""; ret = "Error"; return; }

            }

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
            string seq_new = (param.token_doc + "");
            string pha_seq = "";
            string pha_no = "";


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

                #region check / update seq
                if (pha_status == "11" || dsData.Tables["header"].Rows.Count > 0)
                {
                    if ((dsData.Tables["header"].Rows[0]["action_type"] + "") == "insert")
                    {
                        sqlstr = @" select seq from EPHA_F_HEADER a where lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";
                        cls_conn = new ClassConnectionDb();
                        dt = new DataTable();
                        dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
                        if (dt.Rows.Count > 0)
                        {
                            for (int t = 0; t < dsData.Tables.Count; t++)
                            {
                                for (int i = 0; i < dsData.Tables[t].Rows.Count; i++)
                                {
                                    try
                                    {
                                        if (dsData.Tables[t].TableName == "header")
                                        {
                                            dsData.Tables[t].Rows[i]["seq"] = seq_header_now;
                                            dsData.Tables[t].Rows[i]["id"] = seq_header_now;
                                        }
                                        else
                                        {
                                            dsData.Tables[t].Rows[i]["id_pha"] = seq_header_now;
                                        }
                                    }
                                    catch { }
                                }
                            }
                            dsData.AcceptChanges();
                        }
                    }
                }
                #endregion check / update seq

            }

            ClassHazop cls_old = new ClassHazop();
            DataSet dsDataOld = new DataSet();

            string sub_expense_type = "";
            try
            {
                sub_expense_type = (dsData.Tables["general"].Rows[0]["sub_expense_type"] + "");
            }
            catch { }


            if (param.flow_action == "cancel")
            {
                if (pha_status == "11")
                {
                    cls = new ClassFunctions();
                    cls_conn = new ClassConnectionDb();
                    cls_conn.OpenConnection();
                    cls_conn.BeginTransaction();

                    int i = 0;
                    dt = new DataTable();
                    dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();

                    string pha_status_new = "81";

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
                }
                return cls_json.SetJSONresult(refMsg(ret, msg, seq_new));
            }

            if (dsData.Tables["header"].Rows.Count > 0)
            {
                #region ตรวจสอบ กรณีที่เป็นการเปลี่ยน session 
                //12	WP	PHA Conduct 
                try
                {
                    if (pha_status == "12" && dsData.Tables["session"].Rows.Count > 1)
                    {
                        dt = new DataTable();
                        dt = dsData.Tables["session"].Copy(); dt.AcceptChanges();

                        DataRow[] dr = dt.Select("action_type = 'insert'");
                        if (dr.Length > 0)
                        {
                            //กรณีที่มีมากกว่า 0 ให้ keep version เดิมและ new version ใหม่ 
                            //header,general,functional_audition,session,memberteam,drawing,node,nodedrawing,nodeguidwords,nodeworksheet

                            //update seq_header_now to id,seq or id_pha  
                            dsData.Tables["header"].Rows[0]["id"] = seq_header_now;
                            dsData.Tables["header"].Rows[0]["seq"] = seq_header_now;
                            dsData.Tables["header"].Rows[0]["action_type"] = "insert";
                            dsData.AcceptChanges();

                            //update seq_header_now to id_pha  
                            string[] xsplitTable = ("general,functional_audition,session,memberteam,drawing,node,nodedrawing,nodeguidwords,nodeworksheet").Split(',');
                            for (int t = 0; t < xsplitTable.Length; t++)
                            {
                                string table = xsplitTable[t].ToString();
                                dsData.Tables[table].Rows[0]["id_pha"] = seq_header_now;
                                dsData.Tables[table].Rows[0]["action_type"] = "insert";
                                dsData.AcceptChanges();
                            }
                        }
                    }
                }
                catch { }
                #endregion ตรวจสอบ กรณีที่เป็นการเปลี่ยน session 


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

                    if (pha_status == "11" || pha_status == "21")
                    {
                        if (pha_status == "21")
                        {
                            dsData.Tables["header"].Rows[0]["PHA_VERSION"] = Convert.ToInt32((dsData.Tables["header"].Rows[0]["PHA_VERSION"] + "")) + 1;
                            dsData.AcceptChanges();
                        }
                        ret = set_hazop_header(ref dsData, ref cls_conn_header, seq_header_now);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }
                    }
                    if (pha_status == "11" || pha_status == "12" || pha_status == "22")
                    {
                        #region update case SAFETY_CRITICAL_EQUIPMENT_SHOW
                        if (dsData.Tables["header"].Rows.Count > 0)
                        {
                            dt = new DataTable();
                            dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();
                            if ((dt.Rows[0]["action_type"] + "") == "update")
                            {
                                int i = 0;
                                sqlstr = "update  EPHA_F_HEADER set ";
                                sqlstr += " SAFETY_CRITICAL_EQUIPMENT_SHOW = " + cls.ChkSqlNum((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_SHOW"] + "").ToString(), "N");

                                sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                                sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                                sqlstr += " and YEAR = " + cls.ChkSqlNum((dt.Rows[i]["YEAR"] + "").ToString(), "N");
                                sqlstr += " and PHA_NO = " + cls.ChkSqlStr((dt.Rows[i]["PHA_NO"] + "").ToString(), 200);
                                ret = cls_conn_header.ExecuteNonQuery(sqlstr);
                                if (ret == "") { ret = "true"; }
                                if (ret != "true") { goto Next_Line; }
                            }
                        }
                        #endregion update case SAFETY_CRITICAL_EQUIPMENT_SHOW


                        ret = set_pha_parti(ref dsData, ref cls_conn_header, seq_header_now, dsDataOld);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }

                        ret = set_hazop_partii(ref dsData, ref cls_conn_node, seq_header_now);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }

                        ret = set_hazop_partiii(ref dsData, ref cls_conn_node, seq_header_now);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }

                        //ret = set_hazop_partiv(ref dsData, ref cls_conn_node, seq_header_now);
                        //if (ret == "") { ret = "true"; }
                        //if (ret != "true") { goto Next_Line; } 

                        if (dsData.Tables["ram_level"] != null)
                        {
                            DataTable dtDef = dsData.Tables["ram_level"].Copy(); dtDef.AcceptChanges();
                            ret = set_ram_level(dtDef, ref cls_conn_node, seq_header_now);
                            if (ret == "") { ret = "true"; }
                            if (ret != "true") { goto Next_Line; }
                        }
                        if (dsData.Tables["ram_master"] != null)
                        {
                            DataTable dtDef = dsData.Tables["ram_master"].Copy(); dtDef.AcceptChanges();
                            ret = set_ram_master(dtDef, ref cls_conn_node, seq_header_now);
                            if (ret == "") { ret = "true"; }
                            if (ret != "true") { goto Next_Line; }
                        }
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
                    //update seq new document
                    if (pha_status == "11") { seq_new = seq_header_now; }
                    pha_seq = seq_header_now;
                    pha_no = (dsData.Tables["header"].Rows[0]["PHA_NO"] + "");

                    //11	DF	Draft
                    //12	WP	PHA Conduct 
                    //21	WA	Waiting Approve Review
                    //22	AR	Approve Reject
                    //13	WF	Waiting Follow Up
                    //14	WR	Waiting Review Follow Up
                    //91	CL	Closed
                    //81	CN	Cancle

                    ClassEmail clsmail = new ClassEmail();
                    if ((param.flow_action == "submit" || param.flow_action == "submit_without") && sub_expense_type == "Normal")
                    {
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

                            //13	WF	Waiting Follow Up
                            string pha_status_new = "13";
                            if (dsData.Tables["header"].Rows[0]["request_approver"].ToString() == "1" ||
                              (dsData.Tables["general"].Rows[0]["expense_type"].ToString() == "CAPEX" && dsData.Tables["general"].Rows[0]["sub_expense_type"].ToString() == "Normal"))
                            {
                                //21	WA	Waiting Approve Review
                                pha_status_new = "21";
                            }

                            int i = 0;
                            dt = new DataTable();
                            dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();

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


                            if (param.flow_action == "submit")
                            {
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

                            string pha_status_new = "13";
                            if ((dt.Rows[0]["approve_status"] + "") == "reject") { pha_status_new = "22"; }

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

                            if (param.flow_action == "submit")
                            {
                                //13	WF	Waiting Follow Up
                                if (pha_status_new == "13")
                                {
                                    clsmail = new ClassEmail();
                                    clsmail.MailApprovByApprover((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");

                                    clsmail = new ClassEmail();
                                    clsmail.MailToActionOwner((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");
                                }
                                else if (pha_status_new == "22")
                                {
                                    //22	AR	Approve Reject
                                    clsmail = new ClassEmail();
                                    clsmail.MailRejectByApprover((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");
                                }
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
                    else if (param.flow_action == "submit" && sub_expense_type == "Study")
                    {
                        if (pha_status == "11")
                        {

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

                            clsmail = new ClassEmail();
                            clsmail.MailToAdminCaseStudy((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");

                        }
                    }

                }
                #endregion  flow action  submit 

            }

        Next_Line_Convert:;
            return cls_json.SetJSONresult(refMsgSave(ret, msg, seq_new, pha_seq, pha_no));
        }
        public string set_jsea(SetDocHazopModel param)
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
            string seq_new = (param.token_doc + "");
            string pha_seq = "";
            string pha_no = "";


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
                #region check / update seq
                if (pha_status == "11" || dsData.Tables["header"].Rows.Count > 0)
                {
                    if ((dsData.Tables["header"].Rows[0]["action_type"] + "") == "insert")
                    {
                        sqlstr = @" select seq from EPHA_F_HEADER a where lower(a.seq) = lower(" + cls.ChkSqlStr(seq, 50) + ")  ";
                        cls_conn = new ClassConnectionDb();
                        dt = new DataTable();
                        dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
                        if (dt.Rows.Count > 0)
                        {
                            for (int t = 0; t < dsData.Tables.Count; t++)
                            {
                                for (int i = 0; i < dsData.Tables[t].Rows.Count; i++)
                                {
                                    try
                                    {
                                        if (dsData.Tables[t].TableName == "header")
                                        {
                                            dsData.Tables[t].Rows[i]["seq"] = seq_header_now;
                                            dsData.Tables[t].Rows[i]["id"] = seq_header_now;
                                        }
                                        else
                                        {
                                            dsData.Tables[t].Rows[i]["id_pha"] = seq_header_now;
                                        }
                                    }
                                    catch { }
                                }
                            }
                            dsData.AcceptChanges();
                        }
                    }
                }
                #endregion check / update seq
            }

            ClassHazop cls_old = new ClassHazop();
            DataSet dsDataOld = new DataSet();

            string sub_expense_type = "";
            try
            {
                sub_expense_type = (dsData.Tables["general"].Rows[0]["sub_expense_type"] + "");
            }
            catch { }


            if (param.flow_action == "cancel")
            {
                if (pha_status == "11")
                {
                    cls = new ClassFunctions();
                    cls_conn = new ClassConnectionDb();
                    cls_conn.OpenConnection();
                    cls_conn.BeginTransaction();

                    int i = 0;
                    dt = new DataTable();
                    dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();

                    string pha_status_new = "81";

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
                }
                return cls_json.SetJSONresult(refMsg(ret, msg, seq_new));
            }

            if (dsData.Tables["header"].Rows.Count > 0)
            {
                #region ตรวจสอบ กรณีที่เป็นการเปลี่ยน session 
                //12	WP	PHA Conduct 
                try
                {
                    if (pha_status == "12" && dsData.Tables["session"].Rows.Count > 1)
                    {
                        dt = new DataTable();
                        dt = dsData.Tables["session"].Copy(); dt.AcceptChanges();

                        DataRow[] dr = dt.Select("action_type = 'insert'");
                        if (dr.Length > 0)
                        {
                            //กรณีที่มีมากกว่า 0 ให้ keep version เดิมและ new version ใหม่ 
                            //header,general,functional_audition,session,memberteam,drawing,node,nodedrawing,nodeguidwords,nodeworksheet

                            //update seq_header_now to id,seq or id_pha  
                            dsData.Tables["header"].Rows[0]["id"] = seq_header_now;
                            dsData.Tables["header"].Rows[0]["seq"] = seq_header_now;
                            dsData.Tables["header"].Rows[0]["action_type"] = "insert";
                            dsData.AcceptChanges();

                            //update seq_header_now to id_pha  
                            string[] xsplitTable = ("general,functional_audition,session,memberteam,drawing,node,nodedrawing,nodeguidwords,nodeworksheet").Split(',');
                            for (int t = 0; t < xsplitTable.Length; t++)
                            {
                                string table = xsplitTable[t].ToString();
                                dsData.Tables[table].Rows[0]["id_pha"] = seq_header_now;
                                dsData.Tables[table].Rows[0]["action_type"] = "insert";
                                dsData.AcceptChanges();
                            }
                        }
                    }
                }
                catch { }
                #endregion ตรวจสอบ กรณีที่เป็นการเปลี่ยน session 


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

                    if (pha_status == "11" || pha_status == "21")
                    {
                        if (pha_status == "21")
                        {
                            dsData.Tables["header"].Rows[0]["PHA_VERSION"] = Convert.ToInt32((dsData.Tables["header"].Rows[0]["PHA_VERSION"] + "")) + 1;
                            dsData.AcceptChanges();
                        }
                        ret = set_hazop_header(ref dsData, ref cls_conn_header, seq_header_now);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }
                    }
                    if (pha_status == "11" || pha_status == "12" || pha_status == "22")
                    {
                        #region update case SAFETY_CRITICAL_EQUIPMENT_SHOW
                        if (dsData.Tables["header"].Rows.Count > 0)
                        {
                            dt = new DataTable();
                            dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();
                            if ((dt.Rows[0]["action_type"] + "") == "update")
                            {
                                int i = 0;
                                sqlstr = "update  EPHA_F_HEADER set ";
                                sqlstr += " SAFETY_CRITICAL_EQUIPMENT_SHOW = " + cls.ChkSqlNum((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_SHOW"] + "").ToString(), "N");

                                sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                                sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                                sqlstr += " and YEAR = " + cls.ChkSqlNum((dt.Rows[i]["YEAR"] + "").ToString(), "N");
                                sqlstr += " and PHA_NO = " + cls.ChkSqlStr((dt.Rows[i]["PHA_NO"] + "").ToString(), 200);
                                ret = cls_conn_header.ExecuteNonQuery(sqlstr);
                                if (ret == "") { ret = "true"; }
                                if (ret != "true") { goto Next_Line; }
                            }
                        }
                        #endregion update case SAFETY_CRITICAL_EQUIPMENT_SHOW

                        ret = set_pha_parti(ref dsData, ref cls_conn_header, seq_header_now, dsDataOld);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }

                        ret = set_jsea_partiii(ref dsData, ref cls_conn_node, seq_header_now);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }


                        if (dsData.Tables["ram_level"] != null)
                        {
                            DataTable dtDef = dsData.Tables["ram_level"].Copy(); dtDef.AcceptChanges();
                            ret = set_ram_level(dtDef, ref cls_conn_node, seq_header_now);
                            if (ret == "") { ret = "true"; }
                            if (ret != "true") { goto Next_Line; }
                        }
                        if (dsData.Tables["ram_master"] != null)
                        {
                            DataTable dtDef = dsData.Tables["ram_master"].Copy(); dtDef.AcceptChanges();
                            ret = set_ram_master(dtDef, ref cls_conn_node, seq_header_now);
                            if (ret == "") { ret = "true"; }
                            if (ret != "true") { goto Next_Line; }
                        }
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
                    //update seq new document
                    if (pha_status == "11") { seq_new = seq_header_now; }
                    pha_seq = seq_header_now;
                    pha_no = (dsData.Tables["header"].Rows[0]["PHA_NO"] + "");

                    //11	DF	Draft
                    //12	WP	PHA Conduct 
                    //21	WA	Waiting Approve Review
                    //22	AR	Approve Reject
                    //13	WF	Waiting Follow Up
                    //14	WR	Waiting Review Follow Up
                    //91	CL	Closed
                    //81	CN	Cancle

                    ClassEmail clsmail = new ClassEmail();
                    if ((param.flow_action == "submit" || param.flow_action == "submit_without") && sub_expense_type == "Normal")
                    {
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

                            //13	WF	Waiting Follow Up
                            string pha_status_new = "13";
                            if (dsData.Tables["header"].Rows[0]["request_approver"].ToString() == "1" ||
                              (dsData.Tables["general"].Rows[0]["expense_type"].ToString() == "CAPEX" && dsData.Tables["general"].Rows[0]["sub_expense_type"].ToString() == "Normal"))
                            {
                                //21	WA	Waiting Approve Review
                                pha_status_new = "21";
                            }

                            int i = 0;
                            dt = new DataTable();
                            dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();

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


                            if (param.flow_action == "submit")
                            {
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

                            string pha_status_new = "13";
                            if ((dt.Rows[0]["approve_status"] + "") == "reject") { pha_status_new = "22"; }

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

                            if (param.flow_action == "submit")
                            {
                                //13	WF	Waiting Follow Up
                                if (pha_status_new == "13")
                                {
                                    clsmail = new ClassEmail();
                                    clsmail.MailApprovByApprover((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");

                                    clsmail = new ClassEmail();
                                    clsmail.MailToActionOwner((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");
                                }
                                else if (pha_status_new == "22")
                                {
                                    //22	AR	Approve Reject
                                    clsmail = new ClassEmail();
                                    clsmail.MailRejectByApprover((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");
                                }
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
                    else if (param.flow_action == "submit" && sub_expense_type == "Study")
                    {
                        if (pha_status == "11")
                        {

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

                            clsmail = new ClassEmail();
                            clsmail.MailToAdminCaseStudy((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");

                        }
                    }

                }
                #endregion  flow action  submit 

            }

        Next_Line_Convert:;
            return cls_json.SetJSONresult(refMsgSave(ret, msg, seq_new, pha_seq, pha_no));
        }
        public string set_whatif(SetDocHazopModel param)
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
            string seq_new = (param.token_doc + "");
            string pha_seq = "";
            string pha_no = "";


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

                    for (int i = 0; i < dsData.Tables["list"].Rows.Count; i++) { dsData.Tables["list"].Rows[0]["id_pha"] = seq_header_now; dsData.Tables["list"].Rows[0]["action_by"] = "insert"; }
                    for (int i = 0; i < dsData.Tables["listworksheet"].Rows.Count; i++) { dsData.Tables["listworksheet"].Rows[0]["id_pha"] = seq_header_now; dsData.Tables["listworksheet"].Rows[0]["action_by"] = "insert"; }
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

            string sub_expense_type = "";
            try
            {
                sub_expense_type = (dsData.Tables["general"].Rows[0]["sub_expense_type"] + "");
            }
            catch { }


            if (param.flow_action == "cancel")
            {
                if (pha_status == "11")
                {
                    cls = new ClassFunctions();
                    cls_conn = new ClassConnectionDb();
                    cls_conn.OpenConnection();
                    cls_conn.BeginTransaction();

                    int i = 0;
                    dt = new DataTable();
                    dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();

                    string pha_status_new = "81";

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
                }
                return cls_json.SetJSONresult(refMsg(ret, msg, seq_new));
            }

            if (dsData.Tables["header"].Rows.Count > 0)
            {
                #region ตรวจสอบ กรณีที่เป็นการเปลี่ยน session 
                //12	WP	PHA Conduct 
                try
                {
                    if (pha_status == "12" && dsData.Tables["session"].Rows.Count > 1)
                    {
                        dt = new DataTable();
                        dt = dsData.Tables["session"].Copy(); dt.AcceptChanges();

                        DataRow[] dr = dt.Select("action_type = 'insert'");
                        if (dr.Length > 0)
                        {
                            //กรณีที่มีมากกว่า 0 ให้ keep version เดิมและ new version ใหม่ 
                            //header,general,functional_audition,session,memberteam,drawing,list,listdrawing,listguidwords,listworksheet

                            //update seq_header_now to id,seq or id_pha  
                            dsData.Tables["header"].Rows[0]["id"] = seq_header_now;
                            dsData.Tables["header"].Rows[0]["seq"] = seq_header_now;
                            dsData.Tables["header"].Rows[0]["action_type"] = "insert";
                            dsData.AcceptChanges();

                            //update seq_header_now to id_pha  
                            string[] xsplitTable = ("general,functional_audition,session,memberteam,drawing,list,listdrawing,listguidwords,listworksheet").Split(',');
                            for (int t = 0; t < xsplitTable.Length; t++)
                            {
                                string table = xsplitTable[t].ToString();
                                dsData.Tables[table].Rows[0]["id_pha"] = seq_header_now;
                                dsData.Tables[table].Rows[0]["action_type"] = "insert";
                                dsData.AcceptChanges();
                            }
                        }
                    }
                }
                catch { }
                #endregion ตรวจสอบ กรณีที่เป็นการเปลี่ยน session 


                #region connection transaction
                cls = new ClassFunctions();
                ClassConnectionDb cls_conn_header = new ClassConnectionDb();
                ClassConnectionDb cls_conn_list = new ClassConnectionDb();
                ClassConnectionDb cls_conn_worksheet = new ClassConnectionDb();
                ClassConnectionDb cls_conn_managerecom = new ClassConnectionDb();

                cls_conn = new ClassConnectionDb();
                cls_conn_header = new ClassConnectionDb();
                cls_conn_list = new ClassConnectionDb();
                cls_conn_worksheet = new ClassConnectionDb();
                cls_conn_managerecom = new ClassConnectionDb();

                cls_conn.OpenConnection();
                cls_conn_header.OpenConnection();
                cls_conn_list.OpenConnection();
                cls_conn_worksheet.OpenConnection();
                cls_conn_managerecom.OpenConnection();

                cls_conn.BeginTransaction();
                cls_conn_header.BeginTransaction();
                cls_conn_list.BeginTransaction();
                cls_conn_worksheet.BeginTransaction();
                cls_conn_managerecom.BeginTransaction();

                #endregion connection transaction
                try
                {

                    if (pha_status == "11" || pha_status == "21")
                    {
                        if (pha_status == "21")
                        {
                            dsData.Tables["header"].Rows[0]["PHA_VERSION"] = Convert.ToInt32((dsData.Tables["header"].Rows[0]["PHA_VERSION"] + "")) + 1;
                            dsData.AcceptChanges();
                        }
                        ret = set_hazop_header(ref dsData, ref cls_conn_header, seq_header_now);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }
                    }
                    if (pha_status == "11" || pha_status == "12" || pha_status == "22")
                    {
                        #region update case SAFETY_CRITICAL_EQUIPMENT_SHOW
                        if (dsData.Tables["header"].Rows.Count > 0)
                        {
                            dt = new DataTable();
                            dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();
                            if ((dt.Rows[0]["action_type"] + "") == "update")
                            {
                                int i = 0;
                                sqlstr = "update  EPHA_F_HEADER set ";
                                sqlstr += " SAFETY_CRITICAL_EQUIPMENT_SHOW = " + cls.ChkSqlNum((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_SHOW"] + "").ToString(), "N");

                                sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                                sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                                sqlstr += " and YEAR = " + cls.ChkSqlNum((dt.Rows[i]["YEAR"] + "").ToString(), "N");
                                sqlstr += " and PHA_NO = " + cls.ChkSqlStr((dt.Rows[i]["PHA_NO"] + "").ToString(), 200);
                                ret = cls_conn_header.ExecuteNonQuery(sqlstr);
                                if (ret == "") { ret = "true"; }
                                if (ret != "true") { goto Next_Line; }
                            }
                        }
                        #endregion update case SAFETY_CRITICAL_EQUIPMENT_SHOW

                        ret = set_pha_parti(ref dsData, ref cls_conn_header, seq_header_now, dsDataOld);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }

                        ret = set_hazop_partii(ref dsData, ref cls_conn_list, seq_header_now);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }

                        ret = set_hazop_partiii(ref dsData, ref cls_conn_list, seq_header_now);
                        if (ret == "") { ret = "true"; }
                        if (ret != "true") { goto Next_Line; }

                        if (dsData.Tables["ram_level"] != null)
                        {
                            DataTable dtDef = dsData.Tables["ram_level"].Copy(); dtDef.AcceptChanges();
                            ret = set_ram_level(dtDef, ref cls_conn_list, seq_header_now);
                            if (ret == "") { ret = "true"; }
                            if (ret != "true") { goto Next_Line; }
                        }
                        if (dsData.Tables["ram_master"] != null)
                        {
                            DataTable dtDef = dsData.Tables["ram_master"].Copy(); dtDef.AcceptChanges();
                            ret = set_ram_master(dtDef, ref cls_conn_list, seq_header_now);
                            if (ret == "") { ret = "true"; }
                            if (ret != "true") { goto Next_Line; }
                        }
                    }

                }
                catch (Exception ex) { ret = ex.Message.ToString(); goto Next_Line; }

            Next_Line:;

                #region connection transaction
                if (ret == "") { ret = "true"; }
                if (ret == "true")
                {
                    cls_conn_header.CommitTransaction();
                    cls_conn_list.CommitTransaction();
                    cls_conn_worksheet.CommitTransaction();
                    cls_conn_managerecom.CommitTransaction();

                    cls_conn.CommitTransaction();
                }
                else
                {
                    cls_conn_header.RollbackTransaction();
                    cls_conn_list.RollbackTransaction();
                    cls_conn_worksheet.RollbackTransaction();
                    cls_conn_managerecom.RollbackTransaction();

                    cls_conn.RollbackTransaction();
                }
                cls_conn_header.CloseConnection();
                cls_conn_list.CloseConnection();
                cls_conn_worksheet.CloseConnection();
                cls_conn_managerecom.CloseConnection();

                cls_conn.CloseConnection();
                #endregion connection transaction

                #region  flow action  submit  
                if (ret == "true")
                {
                    //update seq new document
                    if (pha_status == "11") { seq_new = seq_header_now; }
                    pha_seq = seq_header_now;
                    pha_no = (dsData.Tables["header"].Rows[0]["PHA_NO"] + "");

                    //11	DF	Draft
                    //12	WP	PHA Conduct 
                    //21	WA	Waiting Approve Review
                    //22	AR	Approve Reject
                    //13	WF	Waiting Follow Up
                    //14	WR	Waiting Review Follow Up
                    //91	CL	Closed
                    //81	CN	Cancle

                    ClassEmail clsmail = new ClassEmail();
                    if ((param.flow_action == "submit" || param.flow_action == "submit_without") && sub_expense_type == "Normal")
                    {
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
                            sqlstr = "update EPHA_F_HEADER set ";
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

                            //13	WF	Waiting Follow Up
                            string pha_status_new = "13";
                            if (dsData.Tables["header"].Rows[0]["request_approver"].ToString() == "1" ||
                              (dsData.Tables["general"].Rows[0]["expense_type"].ToString() == "CAPEX" && dsData.Tables["general"].Rows[0]["sub_expense_type"].ToString() == "Normal"))
                            {
                                //21	WA	Waiting Approve Review
                                pha_status_new = "21";
                            }

                            int i = 0;
                            dt = new DataTable();
                            dt = dsData.Tables["header"].Copy(); dt.AcceptChanges();

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


                            if (param.flow_action == "submit")
                            {
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

                            string pha_status_new = "13";
                            if ((dt.Rows[0]["approve_status"] + "") == "reject") { pha_status_new = "22"; }

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

                            if (param.flow_action == "submit")
                            {
                                //13	WF	Waiting Follow Up
                                if (pha_status_new == "13")
                                {
                                    clsmail = new ClassEmail();
                                    clsmail.MailApprovByApprover((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");

                                    clsmail = new ClassEmail();
                                    clsmail.MailToActionOwner((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");
                                }
                                else if (pha_status_new == "22")
                                {
                                    //22	AR	Approve Reject
                                    clsmail = new ClassEmail();
                                    clsmail.MailRejectByApprover((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");
                                }
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
                                        from EPHA_T_list_WORKSHEET a where a.id_pha = " + (dt.Rows[i]["SEQ"] + "").ToString();

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
                                                     from EPHA_T_list_WORKSHEET a 
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
                    else if (param.flow_action == "submit" && sub_expense_type == "Study")
                    {
                        if (pha_status == "11")
                        {

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

                            clsmail = new ClassEmail();
                            clsmail.MailToAdminCaseStudy((dt.Rows[i]["SEQ"] + "").ToString(), "hazop");

                        }
                    }

                }
                #endregion  flow action  submit 

            }

        Next_Line_Convert:;
            return cls_json.SetJSONresult(refMsgSave(ret, msg, seq_new, pha_seq, pha_no));
        }
        public string set_approve(SetDocApproveModel param)
        {
            string msg = "";
            string ret = "";
            cls_json = new ClassJSON();

            string role_type = (param.role_type + "");
            string user_name = (param.user_name + "");
            string pha_seq = (param.token_doc + "");
            string action = (param.action + "");


            #region ตรวจสอบว่าเป็นผู้อนุมัติหรือป่าว หรือมีการอนุมัติไปแล้วหรือยัง
            sqlstr = @" select distinct approver_user_name,approve_action_type from EPHA_F_HEADER where seq = " + pha_seq;
            if (role_type != "admin")
            {
                sqlstr += @" and lower(approver_user_name) = lower('" + user_name + "')";
            }
            cls_conn = new ClassConnectionDb();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count == 0)
            {
                return cls_json.SetJSONresult(refMsgSave("false", "", "", pha_seq, ""));
            }
            else
            {
                if ((dt.Rows[0]["approve_action_type"] + "") == "2")
                {
                    return cls_json.SetJSONresult(refMsgSave("false", "", "", pha_seq, ""));
                }
            }
            #endregion ตรวจสอบว่าเป็นผู้อนุมัติหรือป่าว หรือมีการอนุมัติไปแล้วหรือยัง


            //11	DF	Draft
            //12	WP	PHA Conduct 
            //21	WA	Waiting Approve Review
            //22	AR	Approve Reject
            //13	WF	Waiting Follow Up
            //14	WR	Waiting Review Follow Up
            //91	CL	Closed
            //81	CN	Cancle
            string pha_status_new = "13";
            if (action == "approve") { }
            else if (action == "reject" || action == "reject_no_comment") { pha_status_new = "22"; }

            cls = new ClassFunctions();
            cls_conn = new ClassConnectionDb();
            cls_conn.OpenConnection();
            cls_conn.BeginTransaction();

            #region update
            //APPROVER_USER_NAME, APPROVE_STATUS, APPROVE_COMMENT, UPDATE_BY, UPDATE_DATE
            sqlstr = "update  EPHA_F_HEADER set ";
            sqlstr += " PHA_STATUS = " + cls.ChkSqlNum((pha_status_new).ToString(), "N");
            sqlstr += " ,APPROVE_ACTION_TYPE = 2";//1=save, 2=submit
            sqlstr += " ,APPROVE_STATUS = " + cls.ChkSqlStr((action.ToLower() + "").ToString(), 200);

            sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((user_name.ToLower() + "").ToString(), 200);
            sqlstr += " ,UPDATE_DATE = getdate()";

            sqlstr += " where SEQ = " + cls.ChkSqlNum((pha_seq + "").ToString(), "N");
            if (role_type != "admin") { sqlstr += " and lower(APPROVER_USER_NAME) = lower(" + cls.ChkSqlStr((user_name + "").ToString(), 200) + ")"; }


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

            if (ret == "true")
            {
                ClassEmail clsmail = new ClassEmail();
                if (pha_status_new == "13")
                {
                    //13	WF	Waiting Follow Up
                    clsmail = new ClassEmail();
                    clsmail.MailToActionOwner(pha_seq, "hazop");
                }
                else if (pha_status_new == "22")
                {
                    //22	AR	Approve Reject
                    clsmail = new ClassEmail();
                    clsmail.MailRejectByApprover(pha_seq, "hazop");
                }

            }

            return cls_json.SetJSONresult(refMsgSave(ret, msg, "", pha_seq, ""));
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
                        ",REQUEST_USER_NAME,REQUEST_USER_DISPLAYNAME,SAFETY_CRITICAL_EQUIPMENT_SHOW" +
                        ",FLOW_MAIL_TO_MEMBER" +
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
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["APPROVE_STATUS"] + "").ToString(), 200);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["APPROVE_COMMENT"] + "").ToString(), 200);

                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["REQUEST_USER_NAME"] + "").ToString(), 50);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["REQUEST_USER_DISPLAYNAME"] + "").ToString(), 4000);
                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_SHOW"] + "").ToString(), "N");
                    sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["FLOW_MAIL_TO_MEMBER"] + "").ToString(), "N");

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
                        sqlstr += " ,APPROVE_STATUS = " + cls.ChkSqlStr((dt.Rows[i]["APPROVE_STATUS"] + "").ToString(), 200);
                        sqlstr += " ,APPROVE_COMMENT = " + cls.ChkSqlStr((dt.Rows[i]["APPROVE_COMMENT"] + "").ToString(), 200);
                    }

                    sqlstr += " ,REQUEST_USER_NAME = " + cls.ChkSqlStr((dt.Rows[i]["REQUEST_USER_NAME"] + "").ToString(), 50);
                    sqlstr += " ,REQUEST_USER_DISPLAYNAME = " + cls.ChkSqlStr((dt.Rows[i]["REQUEST_USER_DISPLAYNAME"] + "").ToString(), 4000);
                    sqlstr += " ,SAFETY_CRITICAL_EQUIPMENT_SHOW = " + cls.ChkSqlNum((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_SHOW"] + "").ToString(), "N");
                    sqlstr += " ,FLOW_MAIL_TO_MEMBER = " + cls.ChkSqlNum((dt.Rows[i]["FLOW_MAIL_TO_MEMBER"] + "").ToString(), "N");


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
        public string set_pha_parti(ref DataSet dsData, ref ClassConnectionDb cls_conn, string seq_header_now, DataSet dsDataOld)
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
                        ",ID_AREA,ID_APU,ID_BUSINESS_UNIT,ID_UNIT_NO,OTHER_AREA,OTHER_APU,OTHER_BUSINESS_UNIT,OTHER_UNIT_NO,OTHER_FUNCTIONAL_LOCATION,FUNCTIONAL_LOCATION  " +
                        ",PHA_REQUEST_NAME,TARGET_START_DATE,TARGET_END_DATE,ACTUAL_START_DATE,ACTUAL_END_DATE  " +
                        ",DESCRIPTIONS,WORK_SCOPE" +
                        ",ID_TOC,ID_TAGID,INPUT_TYPE_EXCEL,TYPES_OF_HAZARD,FILE_UPLOAD_SIZE,FILE_UPLOAD_NAME,FILE_UPLOAD_PATH" +
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
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["OTHER_FUNCTIONAL_LOCATION"] + "").ToString(), 4000);

                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["FUNCTIONAL_LOCATION"] + "").ToString(), 4000);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["PHA_REQUEST_NAME"] + "").ToString(), 4000);
                    sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["TARGET_START_DATE"] + "").ToString());
                    sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["TARGET_END_DATE"] + "").ToString());
                    sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ACTUAL_START_DATE"] + "").ToString());
                    sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ACTUAL_END_DATE"] + "").ToString());

                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DESCRIPTIONS"] + "").ToString(), 4000);
                    sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["WORK_SCOPE"] + "").ToString(), 4000);

                    #region jsea
                    //ID_TOC,ID_TAGID,INPUT_TYPE_EXCEL,TYPES_OF_HAZARD,FILE_UPLOAD_SIZE,FILE_UPLOAD_NAME,FILE_UPLOAD_PATH
                    try
                    {
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_TOC"] + "").ToString(), "N");
                    }
                    catch { sqlstr += " ,null"; }
                    try
                    {
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_TAGID"] + "").ToString(), "N");
                    }
                    catch { sqlstr += " ,null"; } 
                    try
                    {
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["INPUT_TYPE_EXCEL"] + "").ToString(), "N");
                    }
                    catch { sqlstr += " ,null"; }
               
                    try
                    {
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["TYPES_OF_HAZARD"] + "").ToString(), "N");
                    }
                    catch { sqlstr += " ,null"; }
                    try
                    {
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["FILE_UPLOAD_SIZE"] + "").ToString(), "N");
                    }
                    catch { sqlstr += " ,null"; }
                    try
                    {
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["FILE_UPLOAD_NAME"] + "").ToString(), 4000);
                    }
                    catch { sqlstr += " ,null"; }
                    try
                    {
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["FILE_UPLOAD_PATH"] + "").ToString(), 4000);
                    }
                    catch { sqlstr += " ,null"; }
                    #endregion jsea

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
                    sqlstr += " ,OTHER_FUNCTIONAL_LOCATION = " + cls.ChkSqlStr((dt.Rows[i]["OTHER_FUNCTIONAL_LOCATION"] + "").ToString(), 4000);

                    sqlstr += " ,FUNCTIONAL_LOCATION = " + cls.ChkSqlStr((dt.Rows[i]["FUNCTIONAL_LOCATION"] + "").ToString(), 4000);
                    sqlstr += " ,PHA_REQUEST_NAME = " + cls.ChkSqlStr((dt.Rows[i]["PHA_REQUEST_NAME"] + "").ToString(), 4000);
                    sqlstr += " ,TARGET_START_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["TARGET_START_DATE"] + "").ToString());
                    sqlstr += " ,TARGET_END_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["TARGET_END_DATE"] + "").ToString());
                    sqlstr += " ,ACTUAL_START_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ACTUAL_START_DATE"] + "").ToString());
                    sqlstr += " ,ACTUAL_END_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ACTUAL_END_DATE"] + "").ToString());

                    sqlstr += " ,DESCRIPTIONS = " + cls.ChkSqlStr((dt.Rows[i]["DESCRIPTIONS"] + "").ToString(), 4000);
                    sqlstr += " ,WORK_SCOPE = " + cls.ChkSqlStr((dt.Rows[i]["WORK_SCOPE"] + "").ToString(), 4000);

                    sqlstr += " ,UPDATE_DATE = getdate()";
                    sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);


                    #region jsea
                    //ID_TOC,ID_TAGID,INPUT_TYPE_EXCEL,TYPES_OF_HAZARD,FILE_UPLOAD_SIZE,FILE_UPLOAD_NAME,FILE_UPLOAD_PATH
                    try
                    {
                        sqlstr += " ,ID_TOC = " + cls.ChkSqlNum((dt.Rows[i]["ID_TOC"] + "").ToString(), "N");
                    }
                    catch { }
                    try
                    {
                        sqlstr += " ,ID_TAGID = " + cls.ChkSqlNum((dt.Rows[i]["ID_TAGID"] + "").ToString(), "N");
                    }
                    catch { } 
                    try
                    {
                        sqlstr += " ,INPUT_TYPE_EXCEL = " + cls.ChkSqlNum((dt.Rows[i]["INPUT_TYPE_EXCEL"] + "").ToString(), "N");
                    }
                    catch { }
                    try
                    {
                        sqlstr += " ,TYPES_OF_HAZARD = " + cls.ChkSqlNum((dt.Rows[i]["TYPES_OF_HAZARD"] + "").ToString(), "N");
                    }
                    catch { }
                    try
                    {
                        sqlstr += " ,FILE_UPLOAD_SIZE = " + cls.ChkSqlNum((dt.Rows[i]["FILE_UPLOAD_SIZE"] + "").ToString(), "N");
                    }
                    catch { }
                    try
                    {
                        sqlstr += " ,FILE_UPLOAD_NAME = " + cls.ChkSqlStr((dt.Rows[i]["FILE_UPLOAD_NAME"] + "").ToString(), 4000);
                    }
                    catch { }
                    try
                    {
                        sqlstr += " ,FILE_UPLOAD_PATH = " + cls.ChkSqlStr((dt.Rows[i]["FILE_UPLOAD_PATH"] + "").ToString(), 4000);
                    }
                    catch { }
                    #endregion jsea


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
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");

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
                    sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");

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
                            "SEQ,ID,ID_PHA,NO,DOCUMENT_NAME,DOCUMENT_NO,DOCUMENT_FILE_NAME,DOCUMENT_FILE_PATH,DOCUMENT_FILE_SIZE,DESCRIPTIONS" +
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
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["DOCUMENT_FILE_SIZE"] + "").ToString(), "N");
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
                        sqlstr += " ,DOCUMENT_FILE_SIZE = " + cls.ChkSqlNum((dt.Rows[i]["DOCUMENT_FILE_SIZE"] + "").ToString(), "N");
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
                            "SEQ,ID,ID_PHA,NO,NODE,DESIGN_INTENT,DESIGN_CONDITIONS,OPERATING_CONDITIONS,NODE_BOUNDARY,DESCRIPTIONS" +
                            ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";
                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["NODE"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DESIGN_INTENT"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DESIGN_CONDITIONS"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["OPERATING_CONDITIONS"] + "").ToString(), 4000);
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
                        sqlstr += " ,DESIGN_CONDITIONS = " + cls.ChkSqlStr((dt.Rows[i]["DESIGN_CONDITIONS"] + "").ToString(), 4000);
                        sqlstr += " ,OPERATING_CONDITIONS = " + cls.ChkSqlStr((dt.Rows[i]["OPERATING_CONDITIONS"] + "").ToString(), 4000);
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
                            ",CATEGORY_NO,CATEGORY_TYPE,RAM_BEFOR_SECURITY,RAM_BEFOR_LIKELIHOOD,RAM_BEFOR_RISK,MAJOR_ACCIDENT_EVENT,SAFETY_CRITICAL_EQUIPMENT,SAFETY_CRITICAL_EQUIPMENT_TAG,EXISTING_SAFEGUARDS" +
                            ",RAM_AFTER_SECURITY,RAM_AFTER_LIKELIHOOD,RAM_AFTER_RISK,RECOMMENDATIONS,RESPONDER_USER_NAME,RESPONDER_USER_DISPLAYNAME" +
                            ",ESTIMATED_START_DATE,ESTIMATED_END_DATE,ACTION_STATUS" +
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
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_TAG"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["EXISTING_SAFEGUARDS"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_SECURITY"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_LIKELIHOOD"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_RISK"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RECOMMENDATIONS"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_NAME"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_DISPLAYNAME"] + "").ToString(), 4000);

                        sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_START_DATE"] + "").ToString());
                        sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_END_DATE"] + "").ToString());

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
                        sqlstr += " ,SAFETY_CRITICAL_EQUIPMENT = " + cls.ChkSqlStr((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT"] + "").ToString(), 10);
                        sqlstr += " ,SAFETY_CRITICAL_EQUIPMENT_TAG = " + cls.ChkSqlStr((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_TAG"] + "").ToString(), 4000);
                        sqlstr += " ,EXISTING_SAFEGUARDS = " + cls.ChkSqlStr((dt.Rows[i]["EXISTING_SAFEGUARDS"] + "").ToString(), 4000);
                        sqlstr += " ,RAM_AFTER_SECURITY = " + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_SECURITY"] + "").ToString(), 10);
                        sqlstr += " ,RAM_AFTER_LIKELIHOOD = " + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_LIKELIHOOD"] + "").ToString(), 10);
                        sqlstr += " ,RAM_AFTER_RISK = " + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_RISK"] + "").ToString(), 10);
                        sqlstr += " ,RECOMMENDATIONS = " + cls.ChkSqlStr((dt.Rows[i]["RECOMMENDATIONS"] + "").ToString(), 4000);
                        sqlstr += " ,RESPONDER_USER_NAME = " + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_NAME"] + "").ToString(), 50);
                        sqlstr += " ,RESPONDER_USER_DISPLAYNAME = " + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_DISPLAYNAME"] + "").ToString(), 4000);

                        sqlstr += " ,ESTIMATED_START_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_START_DATE"] + "").ToString());
                        sqlstr += " ,ESTIMATED_END_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_END_DATE"] + "").ToString());

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

        public string set_jsea_partiii(ref DataSet dsData, ref ClassConnectionDb cls_conn, string seq_header_now)
        {
            string ret = "";
            #region update data tasksworksheet
            if (dsData.Tables["tasks_worksheet"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["tasks_worksheet"].Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    if (action_type == "insert")
                    {
                        #region insert
                        //SEQ Auto running
                        sqlstr = "insert into EPHA_T_TASKS_WORKSHEET (" +
                            "SEQ,ID,ID_PHA" +
                            ",NO,ROW_TYPE,WORKSTEP_NO,WORKSTEP,TASKDESC_NO,TASKDESC,POTENTAILHAZARD_NO,POTENTAILHAZARD,POSSIBLECASE_NO,POSSIBLECASE,CATEGORY_NO,CATEGORY_TYPE" +
                            ",RAM_BEFOR_SECURITY,RAM_BEFOR_LIKELIHOOD,RAM_BEFOR_RISK,MAJOR_ACCIDENT_EVENT,SAFETY_CRITICAL_EQUIPMENT,EXISTING_SAFEGUARDS,RAM_AFTER_SECURITY,RAM_AFTER_LIKELIHOOD,RAM_AFTER_RISK" +
                            ",RECOMMENDATIONS_NO,RECOMMENDATIONS,SAFETY_CRITICAL_EQUIPMENT_TAG,RESPONDER_ACTION_BY,RESPONDER_USER_NAME,RESPONDER_USER_DISPLAYNAME,ACTION_STATUS,ESTIMATED_START_DATE,ESTIMATED_END_DATE" +
                            ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";


                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["ROW_TYPE"] + "").ToString(), 50);

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["WORKSTEP_NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["WORKSTEP"] + "").ToString(), 4000);

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["TASKDESC_NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["TASKDESC"] + "").ToString(), 4000);

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["POTENTAILHAZARD_NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["POTENTAILHAZARD"] + "").ToString(), 4000);

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["POSSIBLECASE_NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["POSSIBLECASE"] + "").ToString(), 4000);

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["CATEGORY_NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CATEGORY_TYPE"] + "").ToString(), 4000);

                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_BEFOR_SECURITY"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_BEFOR_LIKELIHOOD"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_BEFOR_RISK"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["MAJOR_ACCIDENT_EVENT"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["EXISTING_SAFEGUARDS"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_SECURITY"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_LIKELIHOOD"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_RISK"] + "").ToString(), 10);

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["RECOMMENDATIONS_NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RECOMMENDATIONS"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_TAG"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_ACTION_BY"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_NAME"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_DISPLAYNAME"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["ACTION_STATUS"] + "").ToString(), 50);


                        sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_START_DATE"] + "").ToString());
                        sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_END_DATE"] + "").ToString());


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

                        sqlstr = "update EPHA_T_TASKS_WORKSHEET set ";

                        sqlstr += " NO = " + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");


                        sqlstr += " ,WORKSTEP_NO = " + cls.ChkSqlNum((dt.Rows[i]["WORKSTEP_NO"] + "").ToString(), "N");
                        sqlstr += " ,WORKSTEP = " + cls.ChkSqlStr((dt.Rows[i]["WORKSTEP"] + "").ToString(), 4000);
                        sqlstr += " ,TASKDESC_NO = " + cls.ChkSqlNum((dt.Rows[i]["TASKDESC_NO"] + "").ToString(), "N");
                        sqlstr += " ,TASKDESC = " + cls.ChkSqlStr((dt.Rows[i]["TASKDESC"] + "").ToString(), 4000);
                        sqlstr += " ,POTENTAILHAZARD_NO = " + cls.ChkSqlNum((dt.Rows[i]["POTENTAILHAZARD_NO"] + "").ToString(), "N");
                        sqlstr += " ,POTENTAILHAZARD = " + cls.ChkSqlStr((dt.Rows[i]["POTENTAILHAZARD"] + "").ToString(), 4000);
                        sqlstr += " ,POSSIBLECASE_NO = " + cls.ChkSqlNum((dt.Rows[i]["POSSIBLECASE_NO"] + "").ToString(), "N");
                        sqlstr += " ,POSSIBLECASE = " + cls.ChkSqlStr((dt.Rows[i]["POSSIBLECASE"] + "").ToString(), 4000);
                        sqlstr += " ,CATEGORY_NO = " + cls.ChkSqlNum((dt.Rows[i]["CATEGORY_NO"] + "").ToString(), "N");
                        sqlstr += " ,CATEGORY_TYPE = " + cls.ChkSqlStr((dt.Rows[i]["CATEGORY_TYPE"] + "").ToString(), 4000);

                        sqlstr += " ,RAM_BEFOR_SECURITY = " + cls.ChkSqlStr((dt.Rows[i]["RAM_BEFOR_SECURITY"] + "").ToString(), 10);
                        sqlstr += " ,RAM_BEFOR_LIKELIHOOD = " + cls.ChkSqlStr((dt.Rows[i]["RAM_BEFOR_LIKELIHOOD"] + "").ToString(), 10);
                        sqlstr += " ,RAM_BEFOR_RISK = " + cls.ChkSqlStr((dt.Rows[i]["RAM_BEFOR_RISK"] + "").ToString(), 10);
                        sqlstr += " ,MAJOR_ACCIDENT_EVENT = " + cls.ChkSqlStr((dt.Rows[i]["MAJOR_ACCIDENT_EVENT"] + "").ToString(), 10);
                        sqlstr += " ,SAFETY_CRITICAL_EQUIPMENT = " + cls.ChkSqlStr((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT"] + "").ToString(), 10);
                        sqlstr += " ,EXISTING_SAFEGUARDS = " + cls.ChkSqlStr((dt.Rows[i]["EXISTING_SAFEGUARDS"] + "").ToString(), 4000);
                        sqlstr += " ,RAM_AFTER_SECURITY = " + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_SECURITY"] + "").ToString(), 10);
                        sqlstr += " ,RAM_AFTER_LIKELIHOOD = " + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_LIKELIHOOD"] + "").ToString(), 10);
                        sqlstr += " ,RAM_AFTER_RISK = " + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_RISK"] + "").ToString(), 10);
                        sqlstr += " ,RECOMMENDATIONS_NO = " + cls.ChkSqlNum((dt.Rows[i]["CATEGORY_NO"] + "").ToString(), "N");
                        sqlstr += " ,RECOMMENDATIONS = " + cls.ChkSqlStr((dt.Rows[i]["RECOMMENDATIONS"] + "").ToString(), 4000);
                        sqlstr += " ,SAFETY_CRITICAL_EQUIPMENT_TAG = " + cls.ChkSqlStr((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_TAG"] + "").ToString(), 4000);
                        sqlstr += " ,RESPONDER_ACTION_BY = " + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_ACTION_BY"] + "").ToString(), 4000);
                        sqlstr += " ,RESPONDER_USER_NAME = " + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_NAME"] + "").ToString(), 50);
                        sqlstr += " ,RESPONDER_USER_DISPLAYNAME = " + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_DISPLAYNAME"] + "").ToString(), 4000);

                        sqlstr += " ,ESTIMATED_START_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_START_DATE"] + "").ToString());
                        sqlstr += " ,ESTIMATED_END_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_END_DATE"] + "").ToString());

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
                        sqlstr = "delete from EPHA_T_TASKS_WORKSHEET ";

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
            #endregion update data tasksworksheet

            #region update data tasks_relatedpeople
            if (dsData.Tables["tasks_relatedpeople"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["tasks_relatedpeople"].Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    if (action_type == "insert")
                    {
                        #region insert
                        //SEQ Auto running
                        sqlstr = "insert into EPHA_T_TASKS_RELATEDPEOPLE (" +
                            "SEQ,ID,ID_PHA,ID_TASKS,NO,USER_TYPE,APPROVER_TYPE,USER_NAME,USER_DISPLAYNAME,USER_TITLE,REVIEWER_DATE,CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";

                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_TASKS"] + "").ToString(), "N");

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["USER_TYPE"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["APPROVER_TYPE"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["USER_NAME"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["USER_DISPLAYNAME"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["USER_TITLE"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["REVIEWER_DATE"] + "").ToString());

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

                        sqlstr = "update EPHA_T_TASKS_RELATEDPEOPLE set ";

                        sqlstr += " NO = " + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");

                        sqlstr += " ,APPROVER_TYPE = " + cls.ChkSqlStr((dt.Rows[i]["APPROVER_TYPE"] + "").ToString(), 50);
                        sqlstr += " ,USER_NAME = " + cls.ChkSqlStr((dt.Rows[i]["USER_NAME"] + "").ToString(), 50);
                        sqlstr += " ,USER_DISPLAYNAME = " + cls.ChkSqlStr((dt.Rows[i]["USER_DISPLAYNAME"] + "").ToString(), 4000);
                        sqlstr += " ,USER_TITLE = " + cls.ChkSqlStr((dt.Rows[i]["USER_TITLE"] + "").ToString(), 50);
                        sqlstr += " ,REVIEWER_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["REVIEWER_DATE"] + "").ToString());

                        sqlstr += " ,UPDATE_DATE = getdate()";
                        sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        //sqlstr += " and ID_TASKS = " + cls.ChkSqlNum((dt.Rows[i]["ID_TASKS"] + "").ToString(), "N");
                        sqlstr += " and USER_TYPE = " + cls.ChkSqlStr((dt.Rows[i]["USER_TYPE"] + "").ToString(), 50);

                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "delete from EPHA_T_TASKS_RELATEDPEOPLE ";

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        //sqlstr += " and ID_TASKS = " + cls.ChkSqlNum((dt.Rows[i]["ID_TASKS"] + "").ToString(), "N");
                        sqlstr += " and USER_TYPE = " + cls.ChkSqlStr((dt.Rows[i]["USER_TYPE"] + "").ToString(), 50);
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
            #endregion update data tasks_relatedpeople

            return ret;

        }

        public string set_whatif_partii(ref DataSet dsData, ref ClassConnectionDb cls_conn, string seq_header_now)
        {
            string ret = "";
            #region update data list
            if (dsData.Tables["list"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["list"].Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    if (action_type == "insert")
                    {
                        #region insert
                        //SEQ Auto running
                        sqlstr = "insert into EPHA_T_LIST (" +
                            "SEQ,ID,ID_PHA,NO,LIST,DESIGN_INTENT,DESIGN_CONDITIONS,OPERATING_CONDITIONS,LIST_BOUNDARY,DESCRIPTIONS" +
                            ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";
                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["LIST"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DESIGN_INTENT"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DESIGN_CONDITIONS"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["OPERATING_CONDITIONS"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["LIST_BOUNDARY"] + "").ToString(), 4000);
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

                        sqlstr = "update EPHA_T_LIST set ";

                        sqlstr += " NO = " + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");
                        sqlstr += " ,LIST = " + cls.ChkSqlStr((dt.Rows[i]["LIST"] + "").ToString(), 4000);
                        sqlstr += " ,DESIGN_INTENT = " + cls.ChkSqlStr((dt.Rows[i]["DESIGN_INTENT"] + "").ToString(), 4000);
                        sqlstr += " ,DESIGN_CONDITIONS = " + cls.ChkSqlStr((dt.Rows[i]["DESIGN_CONDITIONS"] + "").ToString(), 4000);
                        sqlstr += " ,OPERATING_CONDITIONS = " + cls.ChkSqlStr((dt.Rows[i]["OPERATING_CONDITIONS"] + "").ToString(), 4000);
                        sqlstr += " ,LIST_BOUNDARY = " + cls.ChkSqlStr((dt.Rows[i]["LIST_BOUNDARY"] + "").ToString(), 4000);
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
                        sqlstr = "delete from EPHA_T_LIST ";

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
            #endregion update data list

            #region update data listdrawing
            if (dsData.Tables["listdrawing"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["listdrawing"].Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    if (action_type == "insert")
                    {
                        #region insert
                        //SEQ Auto running
                        sqlstr = "insert into EPHA_T_LIST_DRAWING (" +
                            "SEQ,ID,ID_PHA,ID_LIST,ID_DRAWING,NO,PAGE_START_FIRST,PAGE_END_FIRST,PAGE_START_SECOND,PAGE_END_SECOND,PAGE_START_THIRD,PAGE_END_THIRD,DESCRIPTIONS" +
                            ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";
                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_LIST"] + "").ToString(), "N");
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

                        sqlstr = "update EPHA_T_LIST_DRAWING set ";

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
                        sqlstr += " and ID_LIST = " + cls.ChkSqlNum((dt.Rows[i]["ID_LIST"] + "").ToString(), "N");

                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "delete from EPHA_T_LIST_DRAWING ";

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and ID_LIST = " + cls.ChkSqlNum((dt.Rows[i]["ID_LIST"] + "").ToString(), "N");
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
            #endregion update data listdrawing

            return ret;

        }
        public string set_whatif_partiii(ref DataSet dsData, ref ClassConnectionDb cls_conn, string seq_header_now)
        {
            string ret = "";
            #region update data listworksheet
            if (dsData.Tables["listworksheet"] != null)
            {
                dt = new DataTable();
                dt = dsData.Tables["listworksheet"].Copy(); dt.AcceptChanges();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                    if (action_type == "insert")
                    {
                        #region insert
                        //SEQ Auto running
                        sqlstr = "insert into EPHA_T_LIST_WORKSHEET (" +
                            "SEQ,ID,ID_PHA,ROW_TYPE,ID_NODE,NO,LIST_SYSTEM_NO,LIST_SYSTEM,LIST_SUB_SYSTEM_NO,LIST_SUB_SYSTEM,CAUSES_NO,CAUSES,CONSEQUENCES_NO,CONSEQUENCES" +
                            ",CATEGORY_NO,CATEGORY_TYPE,RAM_BEFOR_SECURITY,RAM_BEFOR_LIKELIHOOD,RAM_BEFOR_RISK,MAJOR_ACCIDENT_EVENT,SAFETY_CRITICAL_EQUIPMENT,SAFETY_CRITICAL_EQUIPMENT_TAG,EXISTING_SAFEGUARDS" +
                            ",RAM_AFTER_SECURITY,RAM_AFTER_LIKELIHOOD,RAM_AFTER_RISK,RECOMMENDATIONS,RESPONDER_USER_NAME,RESPONDER_USER_DISPLAYNAME" +
                            ",ESTIMATED_START_DATE,ESTIMATED_END_DATE,ACTION_STATUS" +
                            ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                            ") values ";
                        sqlstr += " ( ";
                        sqlstr += " " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlNum(seq_header_now, "N");

                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["ROW_TYPE"] + "").ToString(), 50);

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ID_NODE"] + "").ToString(), "N");

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["LIST_SYSTEM_NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["LIST_SYSTEM"] + "").ToString(), 4000);

                        sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["LIST_SUB_SYSTEM_NO"] + "").ToString(), "N");
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["LIST_SUB_SYSTEM"] + "").ToString(), 4000);

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
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_TAG"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["EXISTING_SAFEGUARDS"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_SECURITY"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_LIKELIHOOD"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_RISK"] + "").ToString(), 10);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RECOMMENDATIONS"] + "").ToString(), 4000);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_NAME"] + "").ToString(), 50);
                        sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_DISPLAYNAME"] + "").ToString(), 4000);

                        sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_START_DATE"] + "").ToString());
                        sqlstr += " ," + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_END_DATE"] + "").ToString());

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

                        sqlstr = "update EPHA_T_LIST_WORKSHEET set ";

                        sqlstr += " NO = " + cls.ChkSqlNum((dt.Rows[i]["NO"] + "").ToString(), "N");

                        sqlstr += " ,LIST_SYSTEM_NO = " + cls.ChkSqlNum((dt.Rows[i]["LIST_SYSTEM_NO"] + "").ToString(), "N");
                        sqlstr += " ,LIST_SYSTEM = " + cls.ChkSqlStr((dt.Rows[i]["LIST_SYSTEM"] + "").ToString(), 4000);

                        sqlstr += " ,LIST_SUB_SYSTEM_NO = " + cls.ChkSqlNum((dt.Rows[i]["LIST_SYSTEM_NO"] + "").ToString(), "N");
                        sqlstr += " ,LIST_SUB_SYSTEM = " + cls.ChkSqlStr((dt.Rows[i]["LIST_SYSTEM"] + "").ToString(), 4000);

                        sqlstr += " ,CAUSES_NO = " + cls.ChkSqlNum((dt.Rows[i]["CAUSES_NO"] + "").ToString(), "N");
                        sqlstr += " ,CAUSES = " + cls.ChkSqlStr((dt.Rows[i]["CAUSES"] + "").ToString(), 4000);

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
                        sqlstr += " ,SAFETY_CRITICAL_EQUIPMENT = " + cls.ChkSqlStr((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT"] + "").ToString(), 10);
                        sqlstr += " ,SAFETY_CRITICAL_EQUIPMENT_TAG = " + cls.ChkSqlStr((dt.Rows[i]["SAFETY_CRITICAL_EQUIPMENT_TAG"] + "").ToString(), 4000);
                        sqlstr += " ,EXISTING_SAFEGUARDS = " + cls.ChkSqlStr((dt.Rows[i]["EXISTING_SAFEGUARDS"] + "").ToString(), 4000);
                        sqlstr += " ,RAM_AFTER_SECURITY = " + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_SECURITY"] + "").ToString(), 10);
                        sqlstr += " ,RAM_AFTER_LIKELIHOOD = " + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_LIKELIHOOD"] + "").ToString(), 10);
                        sqlstr += " ,RAM_AFTER_RISK = " + cls.ChkSqlStr((dt.Rows[i]["RAM_AFTER_RISK"] + "").ToString(), 10);
                        sqlstr += " ,RECOMMENDATIONS = " + cls.ChkSqlStr((dt.Rows[i]["RECOMMENDATIONS"] + "").ToString(), 4000);
                        sqlstr += " ,RESPONDER_USER_NAME = " + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_NAME"] + "").ToString(), 50);
                        sqlstr += " ,RESPONDER_USER_DISPLAYNAME = " + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_USER_DISPLAYNAME"] + "").ToString(), 4000);

                        sqlstr += " ,ESTIMATED_START_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_START_DATE"] + "").ToString());
                        sqlstr += " ,ESTIMATED_END_DATE = " + cls.ChkSqlDateYYYYMMDD((dt.Rows[i]["ESTIMATED_END_DATE"] + "").ToString());

                        sqlstr += " ,UPDATE_DATE = getdate()";
                        sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and ID_LIST = " + cls.ChkSqlNum((dt.Rows[i]["ID_LIST"] + "").ToString(), "N");

                        #endregion update
                    }
                    else if (action_type == "delete")
                    {
                        #region delete
                        sqlstr = "delete from EPHA_T_LIST_WORKSHEET ";

                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and ID_LIST = " + cls.ChkSqlNum((dt.Rows[i]["ID_LIST"] + "").ToString(), "N");
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
            #endregion update data listworksheet
            return ret;

        }

        public string set_ram_level(DataTable _dtDef, ref ClassConnectionDb cls_conn, string seq_header_now)
        {
            string ret = "";
            #region update data ram level
            dt = new DataTable();
            dt = _dtDef.Copy(); dt.AcceptChanges();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                if (action_type == "update")
                {
                    #region update
                    cls = new ClassFunctions();
                    sqlstr = "update EPHA_M_RAM_LEVEL set ";
                    sqlstr += " people = " + cls.ChkSqlStr((dt.Rows[i]["people"] + "").ToString(), 4000);
                    sqlstr += " ,assets = " + cls.ChkSqlStr((dt.Rows[i]["assets"] + "").ToString(), 4000);
                    sqlstr += " ,enhancement = " + cls.ChkSqlStr((dt.Rows[i]["enhancement"] + "").ToString(), 4000);
                    sqlstr += " ,reputation = " + cls.ChkSqlStr((dt.Rows[i]["reputation"] + "").ToString(), 4000);
                    sqlstr += " ,product_quality = " + cls.ChkSqlStr((dt.Rows[i]["product_quality"] + "").ToString(), 4000);
                    sqlstr += " ,security_level = " + cls.ChkSqlStr((dt.Rows[i]["security_level"] + "").ToString(), 4000);

                    for (int c = 1; c < 11; c++)
                    {
                        sqlstr += " ,likelihood" + c + "_level = " + cls.ChkSqlStr((dt.Rows[i]["likelihood" + c + "_level"] + "").ToString(), 4000);
                        sqlstr += " ,likelihood" + c + "_text = " + cls.ChkSqlStr((dt.Rows[i]["likelihood" + c + "_text"] + "").ToString(), 4000);
                        sqlstr += " ,likelihood" + c + "_desc = " + cls.ChkSqlStr((dt.Rows[i]["likelihood" + c + "_desc"] + "").ToString(), 4000);
                        sqlstr += " ,likelihood" + c + "_criterion = " + cls.ChkSqlStr((dt.Rows[i]["likelihood" + c + "_criterion"] + "").ToString(), 4000);
                        sqlstr += " ,ram" + c + "_text = " + cls.ChkSqlStr((dt.Rows[i]["ram" + c + "_text"] + "").ToString(), 4000);
                        sqlstr += " ,ram" + c + "_priority = " + cls.ChkSqlStr((dt.Rows[i]["ram" + c + "_priority"] + "").ToString(), 4000);
                        sqlstr += " ,ram" + c + "_desc = " + cls.ChkSqlStr((dt.Rows[i]["ram" + c + "_desc"] + "").ToString(), 4000);
                        sqlstr += " ,ram" + c + "_color = " + cls.ChkSqlStr((dt.Rows[i]["ram" + c + "_color"] + "").ToString(), 4000);
                    }

                    sqlstr += " ,opportunity_level = " + cls.ChkSqlNum((dt.Rows[i]["opportunity_level"] + "").ToString(), "N");
                    sqlstr += " ,opportunity_desc = " + cls.ChkSqlStr((dt.Rows[i]["opportunity_desc"] + "").ToString(), 4000);
                    sqlstr += " ,security_text = " + cls.ChkSqlStr((dt.Rows[i]["security_text"] + "").ToString(), 4000);


                    sqlstr += " ,UPDATE_DATE = getdate()";
                    //sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                    sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                    sqlstr += " and ID_RAM = " + cls.ChkSqlNum((dt.Rows[i]["ID_RAM"] + "").ToString(), "N");

                    #endregion update

                }
                if (action_type != "")
                {
                    ret = cls_conn.ExecuteNonQuery(sqlstr);
                    if (ret != "true") { break; }
                }
            }
            if (ret == "") { ret = "true"; }
            if (ret != "true") { return ret; }

            #endregion update data ram level
            return ret;

        }
        public string set_ram_master(DataTable _dtDef, ref ClassConnectionDb cls_conn, string seq_header_now)
        {
            string ret = "";
            #region update data ram level
            dt = new DataTable();
            dt = _dtDef.Copy(); dt.AcceptChanges();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                if (action_type == "update")
                {
                    #region update
                    cls = new ClassFunctions();
                    sqlstr = "update EPHA_M_RAM set ";

                    sqlstr += " DOCUMENT_FILE_NAME = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_NAME"] + "").ToString(), 4000);
                    sqlstr += " ,DOCUMENT_FILE_PATH = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_PATH"] + "").ToString(), 4000);
                    sqlstr += " ,ROWS_LEVEL = " + cls.ChkSqlNum((dt.Rows[i]["ROWS_LEVEL"] + "").ToString(), "N");
                    sqlstr += " ,COLUMNS_LEVEL = " + cls.ChkSqlNum((dt.Rows[i]["COLUMNS_LEVEL"] + "").ToString(), "N");

                    sqlstr += " ,UPDATE_DATE = getdate()";
                    //sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                    sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                    sqlstr += " and ID_RAM = " + cls.ChkSqlNum((dt.Rows[i]["ID_RAM"] + "").ToString(), "N");

                    #endregion update

                }
                if (action_type != "")
                {
                    ret = cls_conn.ExecuteNonQuery(sqlstr);
                    if (ret != "true") { break; }
                }
            }
            if (ret == "") { ret = "true"; }
            if (ret != "true") { return ret; }

            #endregion update data ram level
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
            string sub_software = param.sub_software;
            string sqlstr_check = "";

            //$scope.flow_role_type = "admin";//admin,request,responder,approver
            string role_type = ("admin");
            Boolean bOwnerAction = true;//เป็น Owner Action ด้วยหรือป่าว

            string table_name = "";
            if (sub_software == "hazop") { table_name = "EPHA_T_NODE_WORKSHEET"; }
            if (sub_software == "jsea") { table_name = "EPHA_T_TASKS_WORKSHEET"; }
            if (sub_software == "whatif") { table_name = "EPHA_T_LIST_WORKSHEET"; }

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
                        sqlstr = "update " + table_name + " set ";
                        sqlstr += " DOCUMENT_FILE_NAME = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_NAME"] + "").ToString(), 4000);
                        sqlstr += " ,DOCUMENT_FILE_PATH = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_PATH"] + "").ToString(), 4000);
                        sqlstr += " ,ACTION_STATUS = " + cls.ChkSqlStr((dt.Rows[i]["ACTION_STATUS"] + "").ToString(), 50);

                        sqlstr += " ,RESPONDER_COMMENT = " + cls.ChkSqlStr((dt.Rows[i]["RESPONDER_COMMENT"] + "").ToString(), 4000);

                        //RAM_ACTION_SECURITY, RAM_ACTION_LIKELIHOOD, RAM_ACTION_RISK
                        sqlstr += " ,RAM_ACTION_SECURITY = " + cls.ChkSqlStr((dt.Rows[i]["RAM_ACTION_SECURITY"] + "").ToString(), 50);
                        sqlstr += " ,RAM_ACTION_LIKELIHOOD = " + cls.ChkSqlStr((dt.Rows[i]["RAM_ACTION_LIKELIHOOD"] + "").ToString(), 50);
                        sqlstr += " ,RAM_ACTION_RISK = " + cls.ChkSqlStr((dt.Rows[i]["RAM_ACTION_RISK"] + "").ToString(), 50);

                        //sqlstr += " ,RESPONDER_ACTION_TYPE = 1";//0,1,2-> 2 = ห้ามแก้ไข
                        sqlstr += " ,RESPONDER_ACTION_TYPE = " + cls.ChkSqlNum((dt.Rows[i]["RESPONDER_ACTION_TYPE"] + "").ToString(), "N");
                        if ((dt.Rows[i]["RESPONDER_ACTION_TYPE"] + "").ToString() == "2")
                        {
                            sqlstr += " ,RESPONDER_ACTION_DATE = getdate()";
                        }

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
                    if (flow_action == "update")
                    {
                        #region กรณีที่ update รายการเดียว
                        cls = new ClassFunctions();
                        cls_conn = new ClassConnectionDb();
                        cls_conn.OpenConnection();
                        cls_conn.BeginTransaction();

                        #region update responder_action_type ให้เป็น responder_action_type = 2 ห้ามแก้ไข
                        sqlstr = "update " + table_name + " set responder_action_type = 2 ";
                        sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[0]["SEQ"] + "").ToString(), "N");
                        sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[0]["ID_PHA"] + "").ToString(), "N");
                        sqlstr += " and RESPONDER_USER_NAME = " + cls.ChkSqlStr((dt.Rows[0]["RESPONDER_USER_NAME"] + "").ToString(), 50);
                        ret = cls_conn.ExecuteNonQuery(sqlstr);
                        #endregion update responder_action_type ให้เป็น responder_action_type = 2 ห้ามแก้ไข

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
                        #endregion กรณีที่ update รายการเดียว

                        #region check pha no - Action Owner update action items closed all  -> คนสุดท้ายของใบงาน
                        if (true)
                        {
                            string id_pha = (dt.Rows[0]["SEQ"] + "").ToString();
                            string responder_user_name = (dt.Rows[0]["RESPONDER_USER_NAME"] + "").ToString();
                            sqlstr = @" select t.* from ( select nw.id_pha, count(isnull(nw.responder_action_type,0)) - sum(case when isnull(nw.responder_action_type,0) = 2 then 1 else 0 end)  check_action_type
                                        from  " + table_name + "  nw where nw.responder_user_name is not null and nw.id_pha = " + id_pha + " group by nw.id_pha ) t where t.check_action_type =  0 ";
                            cls_conn = new ClassConnectionDb();
                            DataTable dtaction = new DataTable();
                            dtaction = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
                            if (dtaction.Rows.Count > 0)
                            {
                                //mail not admin กรณีที่ Action Owner Update status Closed All 
                                ClassEmail clsmail = new ClassEmail();
                                clsmail.MailToAdminReviewAll(id_pha, "hazop");

                                #region update pha status 
                                string pha_status_new = "14";

                                cls = new ClassFunctions();
                                cls_conn = new ClassConnectionDb();
                                cls_conn.OpenConnection();
                                cls_conn.BeginTransaction();

                                #region update
                                sqlstr = "update EPHA_F_HEADER set ";
                                sqlstr += " PHA_STATUS = " + cls.ChkSqlNum((pha_status_new).ToString(), "N");
                                sqlstr += " where SEQ = " + cls.ChkSqlNum(id_pha, "N");
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
                            else
                            {
                                //ไว้สำหรับส่ง mail แจ้งเตือน admin ข้อมูลต้องเรียงตาม id_pha, responder_user_name 
                                ClassEmail clsmail = new ClassEmail();
                                clsmail.MailNotificationToAdminOwnerUpdateAction(id_pha, responder_user_name, sub_software);
                            }

                        }
                        #endregion check pha no - Action Owner update action items closed all   -> คนสุดท้ายของใบงาน
                    }
                }
                #endregion  flow action  submit 

            }

        Next_Line_Convert:;
            return cls_json.SetJSONresult(refMsg(ret, msg));
        }

        public string set_follow_up_review(SetDocHazopModel param)
        {
            string msg = "";
            string ret = "";
            cls_json = new ClassJSON();

            DataSet dsData = new DataSet();
            string user_name = (param.user_name + "");
            string flow_action = param.flow_action;
            string sub_software = param.sub_software;
            string sqlstr_check = "";

            //$scope.flow_role_type = "admin";//admin,request,responder,approver
            string role_type = ("admin");

            string table_name = "";
            if (sub_software == "hazop") { table_name = "EPHA_T_NODE_WORKSHEET"; }
            if (sub_software == "jsea") { table_name = "EPHA_T_TASKS_WORKSHEET"; }
            if (sub_software == "whatif") { table_name = "EPHA_T_LIST_WORKSHEET"; }

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

            jsper = param.json_general + "";
            try
            {
                if (jsper.Trim() != "")
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "general";
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
                        sqlstr = "update EPHA_T_" + table_name + "_WORKSHEET set ";
                        sqlstr += " DOCUMENT_FILE_NAME = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_NAME"] + "").ToString(), 4000);
                        sqlstr += " ,DOCUMENT_FILE_PATH = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_PATH"] + "").ToString(), 4000);
                        sqlstr += " ,ACTION_STATUS = " + cls.ChkSqlStr((dt.Rows[i]["ACTION_STATUS"] + "").ToString(), 50);

                        //RAM_ACTION_SECURITY, RAM_ACTION_LIKELIHOOD, RAM_ACTION_RISK
                        sqlstr += " ,RAM_ACTION_SECURITY = " + cls.ChkSqlStr((dt.Rows[i]["RAM_ACTION_SECURITY"] + "").ToString(), 50);
                        sqlstr += " ,RAM_ACTION_LIKELIHOOD = " + cls.ChkSqlStr((dt.Rows[i]["RAM_ACTION_LIKELIHOOD"] + "").ToString(), 50);
                        sqlstr += " ,RAM_ACTION_RISK = " + cls.ChkSqlStr((dt.Rows[i]["RAM_ACTION_RISK"] + "").ToString(), 50);

                        //sqlstr += " ,RESPONDER_ACTION_TYPE = 1";//0,1,2-> 2 = ห้ามแก้ไข
                        sqlstr += " ,RESPONDER_ACTION_TYPE = " + cls.ChkSqlNum((dt.Rows[i]["RESPONDER_ACTION_TYPE"] + "").ToString(), "N");

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


                    dt = new DataTable();
                    dt = dsData.Tables["general"].Copy(); dt.AcceptChanges();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                        if (action_type == "insert")
                        {

                            #region update 
                            sqlstr = "update EPHA_T_GENERAL set ";

                            sqlstr += " REVIEW_FOLOW_COMMENT = " + cls.ChkSqlStr((dt.Rows[i]["REVIEW_FOLOW_COMMENT"] + "").ToString(), 4000);

                            sqlstr += " ,UPDATE_DATE = getdate()";
                            sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr(user_name, 50);

                            sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                            sqlstr += " and ID = " + cls.ChkSqlNum((dt.Rows[i]["ID"] + "").ToString(), "N");
                            sqlstr += " and ID_PHA = " + cls.ChkSqlNum((dt.Rows[i]["ID_PHA"] + "").ToString(), "N");
                            #endregion update

                            ret = cls_conn.ExecuteNonQuery(sqlstr);
                            if (ret == "") { ret = "true"; }
                            if (ret != "true") { goto Next_Line; }
                        }

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
                    if (flow_action == "update")
                    {
                        string seq = (dt.Rows[0]["ID_PHA"] + "").ToString();
                        //mail not admin กรณีที่ Action Owner Update status Closed All 
                        ClassEmail clsmail = new ClassEmail();
                        clsmail.MailClosedAll(seq, sub_software);

                        #region update pha status 
                        string pha_status_new = "91";//Closed

                        cls = new ClassFunctions();
                        cls_conn = new ClassConnectionDb();
                        cls_conn.OpenConnection();
                        cls_conn.BeginTransaction();

                        #region update
                        sqlstr = "update EPHA_F_HEADER set ";
                        sqlstr += " PHA_STATUS = " + cls.ChkSqlNum((pha_status_new).ToString(), "N");
                        sqlstr += " where SEQ = " + cls.ChkSqlNum(seq, "N");
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
                #endregion  flow action  submit 

            }

        Next_Line_Convert:;
            return cls_json.SetJSONresult(refMsg(ret, msg));
        }

        public string set_master_ram(SetDocHazopModel param)
        {
            string msg = "";
            string ret = "";
            cls_json = new ClassJSON();

            DataSet dsData = new DataSet();
            string user_name = (param.user_name + "");

            jsper = param.json_ram_master + "";
            try
            {
                if (jsper.Trim() != "")
                {
                    dt = new DataTable();
                    dt = cls_json.ConvertJSONresult(jsper);
                    if (dt != null)
                    {
                        dt.TableName = "ram_master";
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
                int seq_now = 0; int seq_level_now = 0;

                sqlstr = @" select max(a.seq)+1 as max_seq from EPHA_M_RAM a ";
                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
                if (dt.Rows.Count > 0)
                {
                    seq_now = Convert.ToInt32(dt.Rows[0]["max_seq"]);
                }
                sqlstr = @" select max(a.seq)+1 as max_seq from EPHA_M_RAM_LEVEL a ";
                cls_conn = new ClassConnectionDb();
                dt = new DataTable();
                dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];
                if (dt.Rows.Count > 0)
                {
                    seq_level_now = Convert.ToInt32(dt.Rows[0]["max_seq"]);
                }

                cls = new ClassFunctions();

                cls_conn = new ClassConnectionDb();
                cls_conn.OpenConnection();
                cls_conn.BeginTransaction();
                #endregion connection transaction
                try
                {
                    dt = new DataTable();
                    dt = dsData.Tables["ram_master"].Copy(); dt.AcceptChanges();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                        if (action_type == "insert")
                        {
                            #region insert
                            //SEQ Auto running
                            sqlstr = "insert into EPHA_M_RAM (" +
                                "SEQ,ID,NAME,DESCRIPTIONS,ACTIVE_TYPE,CATEGORY_TYPE,DOCUMENT_FILE_NAME,DOCUMENT_FILE_PATH,DOCUMENT_FILE_SIZE,ROWS_LEVEL,COLUMNS_LEVEL" +
                                ",CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY" +
                                ") values ";
                            sqlstr += " ( ";
                            sqlstr += " " + cls.ChkSqlNum(seq_now, "N");
                            sqlstr += " ," + cls.ChkSqlNum(seq_now, "N");
                            sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["NAME"] + "").ToString(), 4000);
                            sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DESCRIPTIONS"] + "").ToString(), 4000);
                            sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ACTIVE_TYPE"] + "").ToString(), "N");
                            sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["CATEGORY_TYPE"] + "").ToString(), "N");

                            sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_NAME"] + "").ToString(), 4000);
                            sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_PATH"] + "").ToString(), 4000);
                            sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["DOCUMENT_FILE_SIZE"] + "").ToString(), "N");
                            sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["ROWS_LEVEL"] + "").ToString(), "N");
                            sqlstr += " ," + cls.ChkSqlNum((dt.Rows[i]["COLUMNS_LEVEL"] + "").ToString(), "N");

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

                            sqlstr = "update EPHA_M_RAM set ";

                            sqlstr += " NAME = " + cls.ChkSqlStr((dt.Rows[i]["NAME"] + "").ToString(), 4000);
                            sqlstr += " ,DESCRIPTIONS = " + cls.ChkSqlStr((dt.Rows[i]["DESCRIPTIONS"] + "").ToString(), 4000);
                            sqlstr += " ,ACTIVE_TYPE = " + cls.ChkSqlNum((dt.Rows[i]["ACTIVE_TYPE"] + "").ToString(), "N");
                            sqlstr += " ,CATEGORY_TYPE = " + cls.ChkSqlNum((dt.Rows[i]["CATEGORY_TYPE"] + "").ToString(), "N");

                            sqlstr += " ,DOCUMENT_FILE_NAME = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_NAME"] + "").ToString(), 4000);
                            sqlstr += " ,DOCUMENT_FILE_PATH = " + cls.ChkSqlStr((dt.Rows[i]["DOCUMENT_FILE_PATH"] + "").ToString(), 4000);
                            sqlstr += " ,DOCUMENT_FILE_SIZE = " + cls.ChkSqlNum((dt.Rows[i]["DOCUMENT_FILE_SIZE"] + "").ToString(), "N");

                            sqlstr += " ,ROWS_LEVEL = " + cls.ChkSqlNum((dt.Rows[i]["ROWS_LEVEL"] + "").ToString(), "N");
                            sqlstr += " ,COLUMNS_LEVEL = " + cls.ChkSqlNum((dt.Rows[i]["COLUMNS_LEVEL"] + "").ToString(), "N");


                            sqlstr += " ,UPDATE_DATE = getdate()";
                            sqlstr += " ,UPDATE_BY = " + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);

                            sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");

                            #endregion update
                        }
                        else if (action_type == "delete")
                        {
                            #region delete
                            sqlstr = "delete from EPHA_M_RAM ";

                            sqlstr += " where SEQ = " + cls.ChkSqlNum((dt.Rows[i]["SEQ"] + "").ToString(), "N");
                            #endregion delete
                        }
                        if (action_type != "")
                        {
                            ret = cls_conn.ExecuteNonQuery(sqlstr);
                            if (ret == "") { ret = "true"; }
                            if (ret != "true") { break; }
                        }
                    }

                    #region genarate ram level  
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string action_type = (dt.Rows[i]["action_type"] + "").ToString();
                        if (action_type == "insert")
                        {
                            int rows_level = Convert.ToInt32((dt.Rows[i]["rows_level"] + "").ToString());
                            int columns_level = Convert.ToInt32((dt.Rows[i]["columns_level"] + "").ToString());
                            for (int ir = 0; ir < rows_level; ir++)
                            {
                                sqlstr = "";
                                sqlstr += @" insert into EPHA_M_RAM_LEVEL (SEQ, ID, ID_RAM, SORT_BY, SECURITY_LEVEL";
                                for (int ic = 1; ic < (columns_level + 1); ic++)
                                {
                                    //likelihood1_text
                                    sqlstr += @" ,LIKELIHOOD" + ic + "_TEXT";
                                }
                                sqlstr += @" ,CREATE_DATE,UPDATE_DATE,CREATE_BY,UPDATE_BY ) values ";
                                sqlstr += " ( ";
                                sqlstr += " " + cls.ChkSqlNum(seq_level_now, "N");
                                sqlstr += " ," + cls.ChkSqlNum(seq_level_now, "N");
                                sqlstr += " ," + cls.ChkSqlNum(seq_now, "N");
                                sqlstr += " ," + cls.ChkSqlNum((ir + 1), "N");
                                sqlstr += " ," + cls.ChkSqlNum(rows_level - ir, "N");
                                for (int ic = 1; ic < (columns_level + 1); ic++)
                                {
                                    //likelihood1_text
                                    sqlstr += " ," + cls.ChkSqlStr(ic, 4000);
                                }
                                sqlstr += " ,getdate()";
                                sqlstr += " ,null";
                                sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["CREATE_BY"] + "").ToString(), 50);
                                sqlstr += " ," + cls.ChkSqlStr((dt.Rows[i]["UPDATE_BY"] + "").ToString(), 50);
                                sqlstr += @" )";

                                ret = cls_conn.ExecuteNonQuery(sqlstr);
                                if (ret == "") { ret = "true"; }
                                if (ret != "true") { break; }

                            }

                            break;
                        }
                    }
                    #endregion genarate ram level  
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

            }

        Next_Line_Convert:;


            dsData = new DataSet();
            dt = new DataTable();
            dt = refMsg(ret, msg);
            dt.TableName = "msg";
            dsData.Tables.Add(dt.Copy()); dsData.AcceptChanges();
            if (ret == "true")
            {
                ClassHazop cls = new ClassHazop();
                cls.get_master_ram(ref dsData);
            }
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(dsData, Newtonsoft.Json.Formatting.Indented);
            return json;
        }

        #endregion save data

        #region email_member_review
        public string set_email_member_review_stamp(SetDocHazopModel param)
        {
            string user_name = (param.user_name + "");
            string id_pha = (param.token_doc + "");
            string sub_software = ("hazop");

            //กรณีที่ Member ได้รับแจ้งเตือนให้ stamp ค่าไว้ว่าเข้ามาดูข้อมูลแล้ว
            string msg = "";
            string ret = "";
            cls_json = new ClassJSON();

            DataSet dsData = new DataSet();
            sqlstr = @" select c.id_session, c.seq 
                        from EPHA_F_HEADER a 
                        inner join EPHA_T_SESSION b  on a.id  = b.id_pha 
                        inner join (select max(id) as id, id_pha from EPHA_T_SESSION group by id_pha ) b2 on b.id = b2.id and b.id_pha = b2.id_pha
                        inner join EPHA_T_MEMBER_TEAM c on a.id  = c.id_pha and b.id  = c.id_session
                        inner join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha ) c2 on c.id_session = c2.id_session and c.id_pha = c2.id_pha ";
            sqlstr += " where lower(a.seq) = lower(" + cls.ChkSqlStr(id_pha, 50) + ") and isnull(b.action_to_review,0) <> 0 and isnull(c.action_review,0) = 0 ";
            if (user_name == "") { sqlstr += " and lower(a.user_name) = lower(" + cls.ChkSqlStr(user_name, 50) + ")  "; }

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count > 0)
            {
                string seq = (dt.Rows[0]["seq"] + "").ToString();
                string id_session = (dt.Rows[0]["id_session"] + "").ToString();

                cls_conn = new ClassConnectionDb();
                cls_conn.OpenConnection();
                cls_conn.BeginTransaction();

                //update table EPHA_T_SESSION  
                sqlstr = "update EPHA_T_SESSION set ";

                sqlstr += " ACTION_TO_REVIEW = 1";//0 = null, 1 = waiting, 2 = send
                sqlstr += " ,DATE_REVIEW = null";

                sqlstr += " where SEQ = " + cls.ChkSqlNum((id_session + "").ToString(), "N");
                sqlstr += " and ID_PHA = " + cls.ChkSqlNum((id_pha + "").ToString(), "N");

                ret = cls_conn.ExecuteNonQuery(sqlstr);
                if (ret != "true") { goto Next_Line; }

            Next_Line:;

                if (ret == "") { ret = "true"; } else { msg = ret; }
                if (ret == "true")
                {
                    cls_conn.CommitTransaction();
                }
                else
                {
                    cls_conn.RollbackTransaction();
                }
                cls_conn.CloseConnection();

            }

            return cls_json.SetJSONresult(refMsg(ret, msg, ""));
        }

        public string set_email_member_review(SetDocHazopModel param)
        {
            string msg = "";
            string ret = "";
            cls_json = new ClassJSON();

            DataSet dsData = new DataSet();
            string user_name = (param.user_name + "");
            string id_pha = (param.token_doc + "");
            string id_session = "";
            string sub_software = "hazop";

            ClassEmail clsmail = new ClassEmail();
            ret = clsmail.MailToMemberReviewPHAConduct(user_name, id_pha, sub_software, ref id_session);

            msg = ret;
            ret = (ret == "" ? "true" : "false");
            if (ret == "true")
            {
                cls_conn = new ClassConnectionDb();
                cls_conn.OpenConnection();
                cls_conn.BeginTransaction();

                //update table EPHA_T_SESSION  
                sqlstr = "update EPHA_T_SESSION set ";

                sqlstr += " DATE_TO_REVIEW = getdate()";
                sqlstr += " ,ACTION_TO_REVIEW = 2";//0 = null, 1 = waiting, 2 = send

                sqlstr += " where SEQ = " + cls.ChkSqlNum((id_session + "").ToString(), "N");
                sqlstr += " and ID_PHA = " + cls.ChkSqlNum((id_pha + "").ToString(), "N");

                ret = cls_conn.ExecuteNonQuery(sqlstr);
                if (ret != "true") { goto Next_Line; }

                //update table EPHA_T_MEMBER_TEAM  
                sqlstr = "update EPHA_T_MEMBER_TEAM set ACTION_REVIEW = 0, DATE_REVIEW = null";//0 = null, 1 = see doc
                sqlstr += " where ID_PHA = " + cls.ChkSqlNum((id_pha + "").ToString(), "N");
                sqlstr += " and ID_SESSION = " + cls.ChkSqlNum((id_session + "").ToString(), "N");

                ret = cls_conn.ExecuteNonQuery(sqlstr);
                if (ret != "true") { goto Next_Line; }

            Next_Line:;

                if (ret == "") { ret = "true"; } else { msg = ret; }
                if (ret == "true")
                {
                    cls_conn.CommitTransaction();
                }
                else
                {
                    cls_conn.RollbackTransaction();
                }
                cls_conn.CloseConnection();
            }

            return cls_json.SetJSONresult(refMsg(ret, msg, ""));
        }
        public string set_member_review(string user_name, string id_pha, string sub_software)
        {
            //กรณีที่ Member ได้รับแจ้งเตือนให้ stamp ค่าไว้ว่าเข้ามาดูข้อมูลแล้ว
            string msg = "";
            string ret = "";
            cls_json = new ClassJSON();

            DataSet dsData = new DataSet();
            //string user_name = (param.user_name + "");
            //string id_pha = (param.token_doc + "");
            //string sub_software = "hazop";

            sqlstr = @" select c.id_session, c.seq 
                        from EPHA_F_HEADER a 
                        inner join EPHA_T_SESSION b  on a.id  = b.id_pha 
                        inner join (select max(id) as id, id_pha from EPHA_T_SESSION group by id_pha ) b2 on b.id = b2.id and b.id_pha = b2.id_pha
                        inner join EPHA_T_MEMBER_TEAM c on a.id  = c.id_pha and b.id  = c.id_session
                        inner join (select max(id_session) as id_session, id_pha from EPHA_T_MEMBER_TEAM group by id_pha ) c2 on c.id_session = c2.id_session and c.id_pha = c2.id_pha ";
            sqlstr += " where lower(a.seq) = lower(" + cls.ChkSqlStr(id_pha, 50) + ") and isnull(b.action_to_review,0) <> 0 and isnull(c.action_review,0) = 0 ";
            if (user_name == "") { sqlstr += " and lower(a.user_name) = lower(" + cls.ChkSqlStr(user_name, 50) + ")  "; }

            cls_conn = new ClassConnectionDb();
            dt = new DataTable();
            dt = cls_conn.ExecuteAdapterSQL(sqlstr).Tables[0];

            if (dt.Rows.Count > 0)
            {
                cls_conn = new ClassConnectionDb();
                cls_conn.OpenConnection();
                cls_conn.BeginTransaction();

                string seq = (dt.Rows[0]["seq"] + "").ToString();
                string id_session = (dt.Rows[0]["id_session"] + "").ToString();

                sqlstr = "update EPHA_T_MEMBER_TEAM set ACTION_REVIEW = 1, DATE_REVIEW = getdate()";

                sqlstr += " where SEQ = " + cls.ChkSqlNum((seq + "").ToString(), "N");
                sqlstr += " and ID_PHA = " + cls.ChkSqlNum((id_pha + "").ToString(), "N");
                sqlstr += " and ID_SESSION = " + cls.ChkSqlNum((id_session + "").ToString(), "N");
                sqlstr += " and USER_NAME = " + cls.ChkSqlStr((user_name + "").ToString(), 50);

                ret = cls_conn.ExecuteNonQuery(sqlstr);
                if (ret != "true") { goto Next_Line; }


            Next_Line:;

                if (ret == "") { ret = "true"; } else { msg = ret; }
                if (ret == "true")
                {
                    cls_conn.CommitTransaction();
                }
                else
                {
                    cls_conn.RollbackTransaction();
                }
                cls_conn.CloseConnection();


            }

            return cls_json.SetJSONresult(refMsg(ret, msg, ""));
        }

        #endregion email_member_review
    }
}
