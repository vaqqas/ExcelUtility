using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using Vqs.Excel;

namespace Vqs.ExcelTest.Models.Horizontal
{
    [ExcelMapper(Header = 1, UseDisplayName = true)]
    public class TeamAttributes : BaseExcelModel
    {
        [Required(ErrorMessage = "Name is required.")]
        [StringLength(100, ErrorMessage = "Name should be 3 to 100 characters long.")]
        //[ExcelMap(ColumnName = "Name")]
        [DisplayName("Name")]
        public string Name { get; set; }

        //[ExcelMap(ColumnName = "Designation")]
        [DisplayName("Designation")]
        public string Designation { get; set; }
        
        [Range(0, 100, ErrorMessage = "Points Scored must be between 0 to 100")]
        //[ExcelMap(ColumnName = "Points Scored")]
        [DisplayName("Points Scored")]
        public int? Points { get; set; }

        //[ExcelMap(ColumnName = "Date Of Birth")]
        [DataType(DataType.Date, ErrorMessage = "Date Of Birth must be a valid date.")]
        [Required(ErrorMessage = "Date of Birth can not be empty.")]
        [Range(typeof(DateTime), "1/1/1980", "12/31/2000", ErrorMessage = "Date of Birth must be between 1/1/1980 to 12/31/2000")]
        [DisplayName("Date Of Birth")]
        public DateTime DOB { get; set; }
    }
}