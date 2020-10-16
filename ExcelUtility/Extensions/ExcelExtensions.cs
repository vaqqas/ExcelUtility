using System;
using System.Collections.Generic;
using System.Reflection;
using OfficeOpenXml;
using System.Linq;

namespace Vqs.Excel
{
    public static class ExcelExtensions
    {
        /// <summary>
        /// Get a row's values from a sheet into a model object
        /// </summary>
        /// <typeparam name="TItem">Type of the model</typeparam>
        /// <param name="sheet">input sheet</param>
        /// <param name="rowOrColumn">SN of the row or column being read</param>
        /// <param name="map">Excel map of the sheet and the model class</param>
        /// <returns>Model object filled with the values from the sheet</returns>
        public static TItem GetRecord<TItem>(this ExcelWorksheet sheet, int rowOrColumn, ExcelMap<TItem> map = null)
            where TItem : class
        {
            if (sheet == null)
            {
                return null;
            }

            if (map == null)
            {
                map = GetMap<TItem>(sheet);
            }

            if (rowOrColumn <= map.Header ||
                (map.MappingDirection == ExcelMappingDirection.Horizontal && rowOrColumn > sheet.Dimension.End.Row) ||
                (map.MappingDirection == ExcelMappingDirection.Vertical && rowOrColumn > sheet.Dimension.End.Column))
            {
                return null;
            }

            return GetItem(sheet, rowOrColumn, map);
        }

        /// <summary>
        /// Get a row's values from a sheet into a model object
        /// </summary>
        /// <typeparam name="TItem">Type of the model</typeparam>
        /// <param name="sheet">input sheet</param>
        /// <param name="rowOrColumn">SN of the row or column being read</param>
        /// <param name="map">Excel map of the sheet and the model class</param>
        /// <returns>Model object filled with the values from the sheet</returns>
        private static TItem GetItem<TItem>(ExcelWorksheet sheet, int rowOrColumn, ExcelMap<TItem> map)
            where TItem : class
        {
            var item = Activator.CreateInstance<TItem>();
            foreach (var mapping in map.Mapping)
            {
                //make sure we don't fall over the edge of the world
                if ((map.MappingDirection == ExcelMappingDirection.Horizontal && mapping.Key > sheet.Dimension.End.Column) ||
                    (map.MappingDirection == ExcelMappingDirection.Vertical && mapping.Key > sheet.Dimension.End.Row))
                {
                    throw new ArgumentOutOfRangeException(nameof(rowOrColumn),
                        $"Key {mapping.Key} is outside of the sheet dimension using direction {map.MappingDirection}");
                }

                //get the value of the cell
                var value = (map.MappingDirection == ExcelMappingDirection.Horizontal)
                    ? sheet.GetValue(rowOrColumn, mapping.Key)
                    : sheet.GetValue(mapping.Key, rowOrColumn);

                //
                if (value != null)
                {
                    // if type is nullable then get the underlying data type
                    var type = mapping.Value.PropertyType.IsValueType
                        ? Nullable.GetUnderlyingType(mapping.Value.PropertyType) ?? mapping.Value.PropertyType
                        : mapping.Value.PropertyType;

                    //convert the value into the target type and set it into the field as per the mapping
                    if (type == typeof(string))
                    {
                        var convertedValue = Convert.ToString(value).Trim();
                        mapping.Value.SetValue(item, convertedValue);
                    }

                    //numeric values throw exception. we need to suppress that and use 0 as the default value.
                    try
                    {
                        string trimmedValue = Convert.ToString(value).Trim();
                        var convertedValue = Convert.ChangeType(trimmedValue, type);
                        mapping.Value.SetValue(item, convertedValue);
                    }
                    catch
                    {
                        var convertedValue = Convert.ChangeType("0", type);
                        mapping.Value.SetValue(item, convertedValue);
                    }                    
                }
                else
                {
                    // Explicitly set null values to prevent properties being initialized with their default values
                    mapping.Value.SetValue(item, null);
                }
            }

            return item;
        }

        /// <summary>
        /// Get the Excel mapping for the Model type and the provided sheet
        /// </summary>
        /// <typeparam name="TItem">Model type</typeparam>
        /// <param name="sheet">input Excel Sheet</param>
        /// <returns>Excel mapping for the Model type and the provided sheet</returns>
        private static ExcelMap<TItem> GetMap<TItem>(ExcelWorksheet sheet)
            where TItem : class
        {
            //We need to use reflection to call CretemMap on the ExcelMap<TItem> as TItem can't be passed to a method
            var method = typeof(ExcelMap<TItem>).GetMethod("CreateMap", BindingFlags.Static | BindingFlags.NonPublic);
            if (method == null)
            {
                throw new ArgumentNullException(nameof(method), $"Method CreateMap not found on type {typeof(ExcelMap<TItem>)}");
            }

            method = method.MakeGenericMethod(typeof(ExcelMap<TItem>));
            var map = method.Invoke(null, new object[] { sheet }) as ExcelMap<TItem>;

            if (map == null)
            {
                throw new ArgumentNullException(nameof(map), $"Map {typeof(ExcelMap<TItem>)} could not be created");
            }

            return map;
        }

        /// <summary>
        /// Get rows from the sheet as per the map
        /// </summary>
        /// <typeparam name="TItem">Type of the model class</typeparam>
        /// <param name="sheet">input ExcelWorkSheet</param>
        /// <param name="map">map. If no map provided it will try to generate one on the basis of annotations.</param>
        /// <returns></returns>
        public static List<TItem> GetRecords<TItem>(this ExcelWorksheet sheet, ExcelMap<TItem> map = null)
            where TItem : class
        {
            if (sheet == null)
            {
                return new List<TItem>();
            }

            if (map == null)
            {
                map = GetMap<TItem>(sheet);
            }

            var items = new List<TItem>();
            var start = map.Header + 1;
            var endDimension = map.MappingDirection == ExcelMappingDirection.Horizontal
                ? sheet.Dimension.End.Row
                : sheet.Dimension.End.Column;
            for (var rowOrColumn = start; rowOrColumn <= endDimension; rowOrColumn++)
            {
                var item = GetItem(sheet, rowOrColumn, map);
                items.Add(item);
            }

            return items;
        }

        /// <summary>
        /// Add a new sheet to the collection. Replace if already exists.
        /// </summary>
        /// <param name="sheets"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static ExcelWorksheet AddOrReplace(this ExcelWorksheets sheets, string name)
        {
            if (sheets.Any(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, name)))
            {
                sheets.Delete(name);
            }

            return sheets.Add(name);
        }
    }
}