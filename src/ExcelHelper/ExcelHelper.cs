namespace ExcelHelper
{
    using System.Linq;
    using System.IO;
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using ClosedXML.Excel;

    //https://closedxml.codeplex.com/
    public class ExcelHelper : IExcelHelper
    {
        public virtual XLWorkbook ListToExcel<T>(List<T> list) where T : class
        {
            using (var xl = new XLWorkbook())
            {
                using (var sheet = xl.AddWorksheet($"ListOf{typeof(T).Name}"))
                {
                    var mappings = new List<Tuple<int, PropertyInfo, XLCellValues>>();
                    var headerRow = sheet.FirstRow();

                    int colIndex = 1;
                    foreach (var prop in typeof(T).GetProperties())
                    {
                        mappings.Add(new Tuple<int, PropertyInfo, XLCellValues>(colIndex, prop, this.GetExcelDataTypeFromCType(prop.PropertyType)));
                        headerRow.Cell(colIndex).Value = this.GetColumnName(prop);
                        headerRow.Style.Font.Bold = true;
                        headerRow.Style.Font.Underline = XLFontUnderlineValues.Single;
                        colIndex++;
                    }

                    var rowCount = 2;
                    foreach (var item in list)
                    {
                        var row = sheet.Row(rowCount);
                        foreach (var map in mappings)
                        {
                            var cell = row.Cell(map.Item1);
                            cell.Value = map.Item2.GetValue(item).ToString();
                            cell.DataType = map.Item3;
                        }
                        rowCount++;
                    }
                    return xl;
                }
            }
        }

        public virtual XLWorkbook BuildExcelTemplate<T>()
        {
            var xl = new XLWorkbook();
            using (var ws = xl.AddWorksheet($"{nameof(T)}Template"))
            {
                var cellCount = ws.FirstCell().WorksheetColumn().ColumnNumber();
                foreach (var prop in typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public).Where(p => p.CanRead))
                {
                    var cell = ws.Cell(row: 1, column: cellCount);
                    cell.Value = this.GetColumnName(prop);
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.Underline = XLFontUnderlineValues.Single;
                    cellCount++;
                }
            }
            return xl;
        }

        public virtual SheetParseResult<T> ParseSheet<T>(IXLWorksheet sheet, Func<T, List<string>> validateT = null) where T : class, new()
        {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));

            var result = new SheetParseResult<T>();
            var headerRow = sheet.FirstRowUsed();
            var lastRow = sheet.LastRowUsed();
            var mappings = GetDefaultPropertyMapParsers<T>(headerRow);

            //foreach over the rows and parse to T
            foreach (var _row in sheet.Rows(firstRow: headerRow.RowBelow().RowNumber(), lastRow: lastRow.RowNumber()))
            {
                result.TotalRecordCount++;
                var row = _row;//modified closure
                var runningValidation = new List<string>();//use to give feedback on parse and validation
                var t = new T();
                foreach (var m in mappings)
                {
                    object val;
                    var cell = row.Cell(m.ExcelColumnIndex);
                    var didParse = m.TryGetProperty(propertyInfo: m.ObjectPropertyInfo, input: cell.GetString(), outVal: out val);
                    if (didParse)
                    {
                        m.ObjectPropertyInfo.SetValue(t, val);
                    }
                    else
                    {
                        runningValidation.Add($"{m.ObjectPropertyInfo.Name} did not parse.");
                    }

                    this.FillCellBackground(cell: ref cell, isValid: didParse);
                }

                if (runningValidation.Count == 0 && validateT != null)
                {
                    runningValidation.AddRange(validateT(t));
                }

                if (runningValidation.Count == 0)
                {
                    result.ValidList.Add(t);
                }

                this.FillRowBackgroundWithValidationMessage(row: ref row, isValid: runningValidation.Count == 0, validationMessages: runningValidation);
            }

            return result;
        }

        public virtual ExcelParseResult<T> ParseExcel<T>(Stream excelStream, Func<T, List<string>> validateT = null) where T : class, new()
        {
            var result = new ExcelParseResult<T>();
            using (var xl = new XLWorkbook(excelStream))
            {
                using (var sheet = xl.Worksheet(1))
                {
                    var sheetResult = this.ParseSheet(sheet: sheet, validateT: validateT);
                    result.ValidList = sheetResult.ValidList;
                    result.TotalRecordCount = sheetResult.TotalRecordCount;
                }
                result.XlWorkbook = xl;
            }
            return result;
        }

        public virtual List<PropertyMapParser> GetDefaultPropertyMapParsers<T>(IXLRow headerRow) where T : class
        {
            var result = new List<PropertyMapParser>();

            var headerCells = headerRow.Cells(
                firstColumn: headerRow.FirstCellUsed().WorksheetColumn().ColumnNumber(),
                lastColumn: headerRow.LastCellUsed().WorksheetColumn().ColumnNumber());

            foreach (var p in typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public).Where(p => p.CanWrite && p.CanRead))
            {
                var cell = headerCells.SingleOrDefault(
                    c => String.Compare(
                             strA: c.GetString(),
                             strB: this.GetColumnName(p),
                             comparisonType: StringComparison.InvariantCultureIgnoreCase) == 0);

                if (cell != null)
                {
                    var mapping = new PropertyMapParser { ObjectPropertyInfo = p, ExcelColumnIndex = cell.WorksheetColumn().ColumnNumber(), TryGetProperty = this.TryParseProperty };
                    result.Add(mapping);
                }
                else
                {
                    throw new Exception($"Excel did not provide required column: {p.Name}");
                }
            }

            return result;
        }

        public virtual bool TryParseProperty(PropertyInfo propertyInfo, string input, out object outVal)
        {
            var result = false;
            outVal = null;

            if (propertyInfo.PropertyType == typeof(string))
            {
                outVal = input;
                result = true;
            }
            else
            {
                var isNullable = (propertyInfo.PropertyType.IsGenericType && propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>));

                //nullable doesn't have try parse, so get the underlying type that does, even though the end value can still be null
                var tryParseType = isNullable
                    ? Nullable.GetUnderlyingType(propertyInfo.PropertyType)
                    : propertyInfo.PropertyType;

                var methodParameterTypes = new Type[] { typeof(string), tryParseType.MakeByRefType() };
                var method = tryParseType.GetMethod("TryParse", methodParameterTypes);

                var didParse = false;
                if (method != null)
                {
                    var args = new object[] { input, null };
                    didParse = (bool)method.Invoke(null, args);
                    if (didParse) outVal = args[1];
                }

                result = didParse || (isNullable && String.IsNullOrWhiteSpace(input));//deciding arbitrarily to not accept "abc" as valid null for int, as an example.
            }

            return result;
        }

        public virtual void FillCellBackground(ref IXLCell cell, bool isValid)
        {
            cell.Style.Fill.BackgroundColor = isValid ? XLColor.NoColor : XLColor.Salmon;
        }

        public virtual void FillRowBackgroundWithValidationMessage(ref IXLRow row, bool isValid, List<string> validationMessages)
        {
            row.Style.Fill.BackgroundColor = isValid ? XLColor.NoColor : XLColor.Salmon;
            if (!isValid && validationMessages != null && validationMessages.Any())//only add messages to invalid rows
            {
                var messageCell = row.Cell(row.LastCellUsed().CellRight().WorksheetColumn().ColumnNumber());
                messageCell.Value = String.Join("->", validationMessages);
            }
        }

        public virtual XLCellValues GetExcelDataTypeFromCType(Type propertyType)
        {
            //just copied from msdn, might not be complete
            XLCellValues result = XLCellValues.Text;
            switch (propertyType.ToString())
            {
                case "System.Int16":
                case "System.Nullable`1[System.Int16]":
                case "System.Int32":
                case "System.Nullable`1[System.Int32]":
                case "System.Int64":
                case "System.Nullable`1[System.Int64]":
                case "System.Decimal":
                case "System.Nullable`1[System.Decimal]":
                case "System.Byte":
                case "System.Nullable`1[System.Byte]":
                case "System.Double":
                case "System.Nullable`1[System.Double]":
                case "System.Single":
                case "System.Nullable`1[System.Single]":
                case "System.IntPtr":
                case "System.Nullable`1[IntPtr]":
                case "System.SByte":
                case "System.Nullable`1[System.SByte]":
                case "System.UInt16":
                case "System.Nullable`1[System.UInt16]":
                case "System.UInt32":
                case "System.Nullable`1[System.UInt32]":
                case "System.UInt64":
                case "System.Nullable`1[UInt64]":
                    result = XLCellValues.Number;
                    break;
                case "System.Boolean":
                case "System.Nullable`1[System.Boolean]":
                    result = XLCellValues.Boolean;
                    break;
                case "System.DateTime":
                case "System.Nullable`1[System.DateTime]":
                    result = XLCellValues.DateTime;
                    break;
                case "System.Char":
                case "System.Nullable`1[System.Char]":
                case "System.String":
                default:
                    result = XLCellValues.Text;
                    break;
            }

            return result;
        }

        /// <summary>
        /// Attempts to get the Column Name from attrubet then property name
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public virtual string GetColumnName(PropertyInfo p)
        {
            return p.GetCustomAttributes(true)
                .OfType<ExcelColumnNameAttribute>()
                .SingleOrDefault()?
                .ColumnName ?? p.Name;
        }
    }
}
