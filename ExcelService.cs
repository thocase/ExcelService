using FluentValidation;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using VIDIShopper.SharedKernel.Interfaces;
using VIDIShopper.SharedKernel.Utils.Attributes;

namespace VIDIShopper.Excel
{
    public class ExcelService : IExcelService
    {
        public List<T> ImportToClass<T>(byte[] file, int worksheetIndex = 0, int headerRowsSize = 1) where T : new()
        {
            var result = new List<T>();
            var excelStructure = GetExcelStructure<T>();
            var columnsNames = excelStructure.Select(x => x.ColumnName).ToArray();

            using (IVIDIExcel excel = VIDIExcelFactory.GetVIDIExcel(file))
            {
                string[] colunasInvalidas = excel.ValidateHeaderNames(worksheetIndex, columnsNames);
                if (colunasInvalidas.Length != 0) return result;

                DataTable dataTable = excel.GetWorkSheet(worksheetIndex, headerRowsSize);
                foreach (DataRow row in dataTable.Rows)
                {
                    var obj = new T();
                    foreach (var propertyT in obj.GetType().GetProperties())
                    {
                        if (!Attribute.GetCustomAttributes(propertyT).Any(x => x is ExcelColumnAttribute))
                            continue;

                        var structure = excelStructure.First(x => x.PropertyName == propertyT.Name);
                        object value = null;

                        if (structure.Type == typeof(string))
                        {
                            value = row.GetString(structure.ColumnName);
                        }

                        else if (structure.Type == typeof(int) || structure.Type == typeof(int?))
                        {
                            value = row.GetInt(structure.ColumnName);
                        }

                        else if (structure.Type == typeof(long) || structure.Type == typeof(long?))
                        {
                            value = row.GetLong(structure.ColumnName);
                        }

                        else if (structure.Type == typeof(float) || structure.Type == typeof(float?))
                        {
                            value = row.GetFloat(structure.ColumnName);
                        }

                        else if (structure.Type == typeof(decimal) || structure.Type == typeof(decimal?))
                        {
                            value = row.GetDecimal(structure.ColumnName);
                        }

                        else if (structure.Type == typeof(double) || structure.Type == typeof(double?))
                        {
                            value = row.GetDouble(structure.ColumnName);
                        }

                        else if (structure.Type == typeof(bool) || structure.Type == typeof(bool?))
                        {
                            value = row.GetBool(structure.ColumnName);
                        }

                        if (value != null)
                        {
                            propertyT.SetValue(obj, value);
                        }
                    }
                    result.Add(obj);
                }
            }

            return result;
        }

        public byte[] ExportToFile<TList>(IEnumerable<TList> list, string sheetName = "Planilha", AbstractValidator<TList> validation = null)
        {
            var excelStructure = GetExcelStructure<TList>();
            var columnsName = excelStructure.Select(x => x.ColumnName).ToArray();

            using (IVIDIExcel excel = VIDIExcelFactory.GetVIDIExcel())
            {
                excel.CreateSheet(sheetName);

                excel.AddHeaders(sheetName, columnsName);
                foreach (var structure in excelStructure)
                {
                    excel.SetColumnFormat(sheetName, structure.Order - 1, structure.Format);
                    excel.SetCellFontColor(sheetName, 0, structure.Order - 1, Color.White);
                    excel.SetCellBackgroundColor(sheetName, 0, structure.Order - 1, Color.FromArgb(112, 134, 207));
                }

                foreach (var item in list)
                {
                    var properties = item.GetType().GetProperties();
                    var orderValues = new Dictionary<int, object>();

                    foreach (var property in properties)
                    {
                        if (!Attribute.GetCustomAttributes(property).Any(x => x is ExcelColumnAttribute))
                            continue;

                        var orderProperty = excelStructure.First(x => x.PropertyName == property.Name).Order;
                        var value = property.GetValue(item);
                        orderValues.Add(orderProperty, value);
                    }

                    var values = orderValues.OrderBy(x => x.Key).Select(x => x.Value).ToArray();
                    var currentRow = excel.AddRow(sheetName, values);

                    if (validation != null)
                    {
                        var validationResult = validation.Validate(item);
                        if (validationResult.IsValid == false)
                        {
                            var errors = validationResult.Errors.GroupBy(x => x.PropertyName).ToList();
                            foreach (var error in errors)
                            {
                                excel.SetRowFontColor(sheetName, currentRow, Color.Red);
                                var structure = excelStructure.FirstOrDefault(x => x.PropertyName == error.Key);
                                var columnId = excelStructure.Count;
                                var errorFromProperty = error.Select(x => x.ErrorMessage).ToArray();
                                var errorMessage = string.Join(" / ", errorFromProperty);
                                if (structure != null)
                                {
                                    columnId = structure.Order - 1;
                                    excel.AddComment(sheetName, currentRow, columnId, errorMessage);
                                }
                                else
                                {
                                    excel.SetCellValue(sheetName, currentRow, columnId, errorMessage);
                                }
                            }
                        }
                    }
                }

                excel.AutoFitColumns(0);
                return excel.GetBytes();
            }
        }

        #region Privates

        private static List<ExcelStructure> GetExcelStructure<T>()
        {
            var result = new List<ExcelStructure>();

            foreach (var property in typeof(T).GetProperties())
            {
                var attributes = Attribute.GetCustomAttributes(property).Where(x => x is ExcelColumnAttribute).ToList();
                foreach (var attribute in attributes)
                {
                    var excelColumnAttribute = (ExcelColumnAttribute)attribute;
                    result.Add(new ExcelStructure
                    {
                        Type = property.PropertyType,
                        PropertyName = property.Name,
                        ColumnName = excelColumnAttribute.GetColumnName(),
                        Order = excelColumnAttribute.GetColumnOrder(),
                        Format = excelColumnAttribute.GetFormat()
                    });
                }
            }

            return result.OrderBy(x => x.Order).ToList();
        }

        class ExcelStructure
        {
            public Type Type { get; set; }
            public string PropertyName { get; set; }
            public string ColumnName { get; set; }
            public int Order { get; set; }
            public string Format { get; set; }
        }

        #endregion
    }
}
