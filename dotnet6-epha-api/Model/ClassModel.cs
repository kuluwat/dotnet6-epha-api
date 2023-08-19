namespace Model
{
    public class ClassModel
    {
    }
    public class LoginUserModel
    {
        public string? user_name { get; set; } = "";  
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


}
