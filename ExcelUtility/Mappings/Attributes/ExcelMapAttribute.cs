using System;

namespace Vqs.Excel
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelMapAttribute : Attribute
    {
        /// <summary>
        /// Use column with HeaderDirection.Horizontal
        /// </summary>
        public int ColumnIndex { get; set; } = -1;

        /// <summary>
        ///  Use column name with HeaderDirection.Horizontal, and Header = 1
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// Use row with HeaderDirection.Vertical
        /// </summary>
        public int RowIndex { get; set; } = -1;

        /// <summary>
        ///  Use row name with HeaderDirection.Vertical, and Header = 1
        /// </summary>
        public string RowName { get; set; }
    }
}