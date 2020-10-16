using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;
using System.ComponentModel;

namespace Vqs.Excel
{
    public class ExcelMap<TItem> where TItem : class
    {
        /// <summary>
        /// Gives direction whether the sheet has columns on the top row or on the left hand side in the first column
        /// </summary>
        public virtual ExcelMappingDirection MappingDirection { get; set; } = ExcelMappingDirection.Horizontal;

        /// <summary>
        /// Row or column number of the header
        /// </summary>
        public virtual int Header { get; set; } = 1;

        /// <summary>
        /// Map of the column number with the model propertly (which column has values for which properties)
        /// </summary>
        public Dictionary<int, PropertyInfo> Mapping { get; private set; } = new Dictionary<int, PropertyInfo>();

        /// <summary>
        /// Create the Map (which column has values for which properties) for the given sheet and the provided model type
        /// </summary>
        /// <typeparam name="TMap"></typeparam>
        /// <param name="sheet"></param>
        /// <returns></returns>
        protected static TMap CreateMap<TMap>(ExcelWorksheet sheet) where TMap : ExcelMap<TItem>
        {
            //we need to use the reflection here as the ExcelMap<TItem> is a generic class and we can't pass TItem as an argument to it.
            var map = Activator.CreateInstance<TMap>();
            var type = typeof(TItem);

            // Check if we map by attributes or by column header name
            var mapper = type.GetCustomAttribute<ExcelMapperAttribute>();

            //if the model class is decorated with the ExcelMapperAttribute but the user hasn't specified the DisplayNames for the properties
            if (mapper != null && !mapper.UseDisplayName)
            {
                // Map by attribute
                map.MappingDirection = mapper.MappingDirection;
                map.Header = mapper.Header;

                type.GetProperties()
                    .Select(x => new { Property = x, Attribute = x.GetCustomAttribute<ExcelMapAttribute>() })
                    .Where(x => x.Attribute != null)
                    .ToList()
                    .ForEach(prop =>
                    {
                        var key = map.MappingDirection == ExcelMappingDirection.Horizontal
                            ? prop.Attribute.ColumnIndex
                            : prop.Attribute.RowIndex;

                        if (key < 0)
                        {
                            throw new ArgumentNullException($"No/invalid Column/Row sequence found on type {typeof(TItem)}");
                        }

                        map.Mapping.Add(key, prop.Property);
                    });
            }

            //if not able to create the map till now than try to map
            if (!map.Mapping.Any())
            {
                // Map by column / row header name
                var props = type.GetProperties().ToList();

                // Determine end dimension for the header
                var endDimension = map.MappingDirection == ExcelMappingDirection.Horizontal
                    ? sheet.Dimension.End.Column
                    : sheet.Dimension.End.Row;
                for (var rowOrColumn = 1; rowOrColumn <= endDimension; rowOrColumn++)
                {
                    //get the row/column header name
                    var headerName = map.MappingDirection == ExcelMappingDirection.Horizontal
                        ? sheet.GetValue<string>(map.Header, rowOrColumn)
                        : sheet.GetValue<string>(rowOrColumn, map.Header);

                    //throw an error if headerName is blank.
                    if (string.IsNullOrWhiteSpace(headerName))
                    {
                        var message = map.MappingDirection == ExcelMappingDirection.Horizontal
                            ? $"Column {rowOrColumn} has no parameter name"
                            : $"Row {rowOrColumn} has no parameter name";
                        throw new ArgumentNullException(nameof(headerName), message);
                    }

                    PropertyInfo prop = null;

                    headerName = headerName.Trim();

                    //If each property has DisplayName
                    if (mapper != null && mapper.UseDisplayName)
                    {
                        //find the property with the specified column name
                        prop = props.FirstOrDefault(
                            x => x.GetCustomAttribute<DisplayNameAttribute>() != null
                            && StringComparer.OrdinalIgnoreCase.Equals(x.GetCustomAttribute<DisplayNameAttribute>().DisplayName, headerName)
                        );
                    }
                    else
                    {
                        // Remove spaces
                        headerName = headerName.Replace(" ", string.Empty).Trim();
                        // Map to property
                        prop = props.FirstOrDefault(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, headerName));
                    }

                    if (prop == null)
                    {
                        throw new ArgumentNullException(nameof(headerName), $"No property {headerName} found on type {typeof(TItem)}");
                    }
                    map.Mapping.Add(rowOrColumn, prop);
                }
            }

            return map;
        }
    }
}