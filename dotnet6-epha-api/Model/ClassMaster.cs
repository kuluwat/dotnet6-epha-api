using System.ComponentModel.DataAnnotations;

namespace Model
{ 

    public class AreaModel
    {
        public int? seq { get; set; } = 0;
        public int? id { get; set; } = 0;

        public string? name { get; set; } = "";
        public string? descriptions { get; set; } = "";

        public int? active_type { get; set; } = 0;

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }
    public class BusinessUnitModel
    {
        public int? seq { get; set; } = 0;
        public int? id { get; set; } = 0;
        public int? id_area { get; set; } = 0;

        public string? name { get; set; } = "";
        public string? descriptions { get; set; } = "";

        public int? active_type { get; set; } = 0;

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }
    public class UnitNoModel
    {
        public int? seq { get; set; } = 0;
        public int? id { get; set; } = 0;
        public int? id_area { get; set; } = 0;
        public int? id_business_unit { get; set; } = 0;

        public string? name { get; set; } = "";
        public string? descriptions { get; set; } = "";

        public int? active_type { get; set; } = 0;

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }


}
