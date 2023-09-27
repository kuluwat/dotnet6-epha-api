using System.ComponentModel.DataAnnotations;

namespace Model
{
    public class uploadFile
    {
        public IFormFileCollection file_obj { get; set; }
        public string? file_of { get; set; }
        public string? file_name { get; set; }
        public string? file_seq { get; set; }
    }
    public class ReportModel
    {
        public string? user_name { get; set; } = "";
        public string? seq { get; set; } = "";
        public string? export_type { get; set; } = "";

    }
    public class CopyFileModel
    {
        //page_start_first,page_start_second,page_end_first,page_end_second
        public string? file_name { get; set; } = "";
        public string? file_path { get; set; } = "";
        public string? page_start_first { get; set; } = "";
        public string? page_start_second { get; set; } = "";
        public string? page_end_first { get; set; } = "";
        public string? page_end_second { get; set; } = "";

    }
    public class EmailConfigModel
    {
        public string? user_name { get; set; } = "";
        public string? user_email { get; set; } = "";

    }
    public class LoadDocModel
    {
        public string? user_name { get; set; } = "";
        public string? token_doc { get; set; } = "";
        public string? sub_software { get; set; } = "";
        public string? type_doc { get; set; } = "";

    }
    public class LoadDocFollowModel
    {
        public string? user_name { get; set; } = "";
        public string? token_doc { get; set; } = "";
        public string? sub_software { get; set; } = "";
        public string? type_doc { get; set; } = "";
        public string? pha_no { get; set; } = "";
        public string? responder_user_name { get; set; } = "";
    }
    public class SetDocHazopModel
    {
        public string? user_name { get; set; } = "";
        public string? token_doc { get; set; } = "";
        public string? pha_status { get; set; } = "";
        public string? pha_version { get; set; } = "";
        public string? action_part { get; set; } = "";
        public string? json_header { get; set; } = "";
        public string? json_general { get; set; } = "";
        public string? json_functional_audition { get; set; } = "";
        public string? json_session { get; set; } = "";
        public string? json_memberteam { get; set; } = "";
        public string? json_drawing { get; set; } = "";
        public string? json_node { get; set; } = "";
        public string? json_nodedrawing { get; set; } = "";
        public string? json_nodeguidwords { get; set; } = "";
        public string? json_nodeworksheet { get; set; } = "";
        public string? json_managerecom { get; set; } = "";
        public string? json_ram_level { get; set; } = "";
        public string? json_ram_master { get; set; } = "";

        public string? flow_action { get; set; } = "";

    }

    public class HeaderModel
    {
        public int? seq { get; set; } = 0;
        public int? id { get; set; } = 0;

        public int? year { get; set; } = 0;
        public string? pha_no { get; set; } = "";
        public int? pha_version { get; set; } = 0;
        public int? pha_status { get; set; } = 0;
        public string? pha_request_by { get; set; } = "";
        public string? pha_request_user_name { get; set; } = "";
        public string? pha_request_user_displayname { get; set; } = "";
        public string? pha_sub_software { get; set; } = "";

        public int? request_approver { get; set; } = 0;
        public string? approver_user_name { get; set; } = "";
        public string? approver_user_displayname { get; set; } = "";
        public int? approve_action_type { get; set; } = 0;
        public int? approve_status { get; set; } = 0;
        public string? approve_comment { get; set; } = "";

        //[create_date] date NULL,
        //[update_date] date NULL, 
        //[create_by] [varchar](4000) NULL,
        //[update_by] [varchar](4000) NULL 
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";


        public string? action_type { get; set; } = "";

    }
    public class GeneralModel
    {
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;
        public int? id_ram { get; set; } = 0;
        public string? ram { get; set; } = "";
        public string? expense_type { get; set; } = "";
        public string? sub_expense_type { get; set; } = "";
        public string? user_name { get; set; } = "";
        public string? approver_user_name { get; set; } = "";
        public string? approver_user_displayname { get; set; } = "";
        public string? approver_user_img { get; set; } = "";
        public string? reference_moc { get; set; } = "";
        public int? id_area { get; set; } = 0;
        public int? id_business_unit { get; set; } = 0;
        public int? id_unit_no { get; set; } = 0;
        public string? functional_location { get; set; } = "";

        //[target_start_date] date NULL,
        //[target_end_date] date NULL, 
        //[actual_start_date] date NULL,
        //[actual_end_date] date NULL,   
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime target_start_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime target_end_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime actual_start_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime actual_end_date { get; set; }

        public string? descriptions { get; set; } = "";


        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";

    }
    public class FunctionalAuditionModel
    {
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;
        public string? functional_location { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }
    public class SessionModel
    {
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;


        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime meeting_date { get; set; }
        public string? meeting_start_time { get; set; } = "";
        public string? meeting_end_time { get; set; } = "";

        public int? request_approver { get; set; } = 0;
        public string? approver_user_name { get; set; } = "";
        public string? approver_user_displayname { get; set; } = "";
        public int? approve_action_type { get; set; } = 0;
        public int? approve_status { get; set; } = 0;
        public string? approve_comment { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";

    }
    public class MemberTeamModel
    {
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;

        public string? user_name { get; set; } = "";
        public string? user_displayname { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";

    }
    public class DrawingModel
    {
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;

        public string? document_name { get; set; } = "";
        public string? document_no { get; set; } = "";
        public string? document_file_name { get; set; } = "";
        public string? document_file_path { get; set; } = "";
        public string? descriptions { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }

    public class NodeModel
    {
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;

        public string? node { get; set; } = "";
        public string? design_intent { get; set; } = "";
        public string? design_conditions { get; set; } = "";
        public string? operating_conditions { get; set; } = "";
        public string? node_boundary { get; set; } = "";
        public string? descriptions { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }
    public class NodeDrawingModel
    {
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;

        public int? id_node { get; set; } = 0;
        public int? id_drawing { get; set; } = 0;
        public int? page_start_first { get; set; } = 0;
        public int? page_end_first { get; set; } = 0;
        public int? page_start_second { get; set; } = 0;
        public int? page_end_second { get; set; } = 0;
        public int? page_start_third { get; set; } = 0;
        public int? page_end_third { get; set; } = 0;

        public string? descriptions { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }

    public class NodeGuidWordsModel
    {
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;

        public int? id_node { get; set; } = 0;
        public int? id_guide_word { get; set; } = 0;

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }
    public class NodeWorksheetModel
    {
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;

        public int? id_node { get; set; } = 0;
        public int? id_guide_word { get; set; } = 0;


        public int? causes_no { get; set; } = 0;
        public string? causes { get; set; } = "";
        public int? consequences_no { get; set; } = 0;
        public string? consequences { get; set; } = "";
        public int? category_no { get; set; } = 0;
        public string? category_type { get; set; } = "";


        public string? security_befor { get; set; } = "";
        public string? likelihood_befor { get; set; } = "";
        public string? risk_befor { get; set; } = "";
        public string? major_accident_event { get; set; } = "";
        public string? existing_safeguards { get; set; } = "";
        public string? security_after { get; set; } = "";
        public string? likelihood_after { get; set; } = "";
        public string? risk_after { get; set; } = "";
        public string? recommendations { get; set; } = "";
        public string? responder_user_name { get; set; } = "";
        public string? responder_user_displayname { get; set; } = "";
        public string? action_status { get; set; } = "";


        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }

    public class ManageRecomModel
    {
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;

        public string? responder_user_name { get; set; } = "";
        public string? responder_user_displayname { get; set; } = "";
        public string? risk_befor { get; set; } = "";
        public string? risk_after { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime meeting_date { get; set; }
        public string? estimated_start_time { get; set; } = "";
        public string? estimated_end_time { get; set; } = "";
        public string? document_file_name { get; set; } = "";
        public string? document_file_path { get; set; } = "";
        public string? action_status { get; set; } = "";
        public int? responder_action_type { get; set; } = 0;

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }


    public class StorageLocationModel
    {
        public int? selected_type { get; set; } = 0;

        public int? id_company { get; set; } = 0;
        public string? name_company { get; set; } = "";
        public int? id_area { get; set; } = 0;
        public string? name_area { get; set; } = "";
        public int? id_apu { get; set; } = 0;
        public string? name_apu { get; set; } = "";
        public int? id_toc { get; set; } = 0;
        public string? name_toc { get; set; } = "";
        public int? id_business_unit { get; set; } = 0;
        public string? name_business_unit { get; set; } = "";
        public int? id_unit_no { get; set; } = 0;
        public string? name_unit_no { get; set; } = "";

    }
    public class GuideWordsModel
    {
        //deviations, guide_words, process_deviation, area_applocation, 0 as selected_type
        public int? selected_type { get; set; } = 0;
        public string? parameter { get; set; } = "";
        public string? deviations { get; set; } = "";
        public string? guide_words { get; set; } = "";
        public string? process_deviation { get; set; } = "";
        public string? area_application { get; set; } = "";

    }
    public class SuggestionCausesModel
    {
        public int? selected_type { get; set; } = 0;

        public string? deviations { get; set; } = "";
        public string? guide_words { get; set; } = "";

        public string? causes { get; set; } = "";

    }
    public class SuggestionRecommendationsModel
    {
        public int? selected_type { get; set; } = 0;

        public string? deviations { get; set; } = "";
        public string? guide_words { get; set; } = "";

        public string? causes { get; set; } = "";
        public string? recommendations { get; set; } = "";

    }
    public class SetDocApproveModel
    {
        public string? role_type { get; set; } = "";
        public string? user_name { get; set; } = "";
        public string? token_doc { get; set; } = ""; 
        public string? action { get; set; } = "";

    }
}