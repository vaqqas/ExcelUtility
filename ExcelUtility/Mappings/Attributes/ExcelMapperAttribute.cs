using System;

namespace Vqs.Excel
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelMapperAttribute : Attribute
    {
        public ExcelMappingDirection MappingDirection { get; set; } = ExcelMappingDirection.Horizontal;

        public int Header { get; set; } = 1;

        public bool UseDisplayName { get; set; } = false;

    }
}