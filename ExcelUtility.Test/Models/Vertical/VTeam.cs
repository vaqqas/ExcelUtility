using Vqs.Excel;
using System;

namespace Vqs.ExcelTest.Models.Vertical
{
    [ExcelMapper(MappingDirection = ExcelMappingDirection.Vertical)]
    public class VTeam
    {
        public string Name { get; set; }
        public string Designation { get; set; }
        public int? Points { get; set; }
        public DateTime DOB { get; set; }
    }
}
