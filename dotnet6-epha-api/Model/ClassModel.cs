namespace Model
{
    public class ClassModel
    {
    }

    public class LoginUserModel
    {
        public string? emp_id { get; set; } = "";  
        public string? user_name { get; set; } = "";  
        public string? user_pass { get; set; } = "";  
    }
    public class LoginUserDetailModel
    {
        public string? user_name { get; set; } = "";
        public string? user_id { get; set; } = "";
        public string? user_email { get; set; } = "";
        public string? user_displayname { get; set; } = "";
        public string? user_img { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? active_type { get; set; } = "";
        public string? input_type { get; set; } = "";
    }

    public class PageUserModel
    {
        public string? user_name { get; set; } = "";
        public string? user_id { get; set; } = "";
        public string? user_email { get; set; } = "";
        public string? user_displayname { get; set; } = "";
        public string? user_decriptions { get; set; } = "";
        public string? user_img { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? active_type { get; set; } = "";
        public string? input_type { get; set; } = "";
    }
    public class RegisterAccountModel
    {
        public string? user_active { get; set; } = ""; 
        public string? user_displayname { get; set; } = ""; 
        public string? user_email { get; set; } = "";
        public string? user_password { get; set; } = "";
        public string? user_password_confirm { get; set; } = ""; 
        public string? accept_status { get; set; } = ""; 
    }


}
