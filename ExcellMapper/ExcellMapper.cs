using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcellMapper
{
    public class ExcelMapper<T>
    {
        public List<T> MapExcel(FileInfo excelFileInfo, int startRow)
        {
            using (var fileStream = excelFileInfo.OpenRead())
            {
                return MapExcel(fileStream, startRow);
            }
        }
        public List<T> MapExcel(IFormFile excelFile, int startRow)
        {
            using (var fileStream = new MemoryStream())
            {
                excelFile.CopyTo(fileStream);
                return MapExcel(fileStream, startRow);
            }
        }
        public List<T> MapExcel(Stream fileStream, int startRow)
        {
            List<T> mappedData = new List<T>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileStream, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                foreach (var worksheetPart in workbookPart.WorksheetParts.ToList())
                {
                    Worksheet worksheet = worksheetPart.Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                    foreach (Row row in sheetData.Elements<Row>().Skip(startRow - 1))
                    {
                        T mappedObject = Activator.CreateInstance<T>();

                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            string columnName = GetColumnNameFromCellReference(cell.CellReference);
                            int actualCellIndex = CellReferenceToIndex(cell);

                            PropertyInfo property = null;

                            if (actualCellIndex >= 0)
                            {
                                // If column index is provided, ignore column name and use index-based mapping
                                property = typeof(T).GetProperties()
                                    .FirstOrDefault(p => p.GetCustomAttribute<ExcelColumnAttribute>()?.ColumnIndex - 1 == actualCellIndex);
                            }
                            if (property is null && actualCellIndex >= 0)
                            {
                                // If column index is not provided, use column name for mapping
                                property = typeof(T).GetProperties().FirstOrDefault(p =>
                                {
                                    var excelColumnAttribute = p.GetCustomAttribute<ExcelColumnAttribute>();
                                    return excelColumnAttribute != null && excelColumnAttribute.ColumnName == columnName;
                                });
                            }
                            if (property != null)
                            {
                                if (!IsPropertyIgnored(property))
                                {
                                    try
                                    {
                                        object cellValue = GetCellValue(cell, workbookPart);
                                        object convertedValue = ConvertValue(cellValue, property.PropertyType);
                                        property.SetValue(mappedObject, convertedValue);
                                    }
                                    catch (Exception) { }
                                }
                            }
                        }

                        mappedData.Add(mappedObject);
                    }
                }
            }

            return mappedData;
        }
        private object ConvertValue(object value, Type targetType)
        {
            if (value == null || targetType.IsInstanceOfType(value))
            {
                return value;
            }
            try
            {
                return Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                // Handle specific conversions dynamically using reflection
                MethodInfo tryParseMethod = targetType.GetMethod("TryParse",
                    BindingFlags.Public | BindingFlags.Static,
                    null,
                    new Type[] { typeof(string), targetType.MakeByRefType() },
                    null);

                try
                {
                    if (tryParseMethod != null)
                    {
                        object[] parameters = { value, null };
                        if ((bool)tryParseMethod.Invoke(null, parameters))
                        {
                            return parameters[1];
                        }
                    }
                }
                catch (Exception) { }

            }
            return value;
        }
        private string GetColumnNameFromCellReference(string cellReference)
        {
            // Extracts the column name from the cell reference (e.g., "A1" -> "A")
            return Regex.Replace(cellReference, "[^A-Z]", string.Empty);
        }
        private int CellReferenceToIndex(Cell cell)
        {
            int index = 0;
            string reference = cell.CellReference.ToString().ToUpper();
            foreach (char ch in reference)
            {
                if (Char.IsLetter(ch))
                {
                    int value = (int)ch - (int)'A';
                    index = (index == 0) ? value : ((index + 1) * 26) + value;
                }
                else
                {
                    return index;
                }
            }
            return index;
        }
        private object GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            // Retrieve the value from the cell based on its data type
            string cellValue = cell.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int sharedStringIndex = int.Parse(cellValue);
                SharedStringItem sharedStringItem = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex);
                return sharedStringItem.Text?.Text;
            }

            return cellValue;
        }

        private bool IsPropertyIgnored(PropertyInfo property)
        {
            // Check if the property is marked with the IgnoreExcelColumnAttribute and its Ignore property is set to true
            return property.GetCustomAttribute<IgnoreExcelColumnAttribute>()?.Ignore == true;
        }

    }
}
