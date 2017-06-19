namespace ExcelHelper
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;
    using ClosedXML.Excel;

    public interface IExcelHelper
    {
        XLWorkbook ListToExcel<T>(List<T> list) where T : class;
        XLWorkbook BuildExcelTemplate<T>();
        ExcelParseResult<T> ParseExcel<T>(Stream excelStream, Func<T, List<string>> validateT = null) where T : class, new();
        List<PropertyMapParser> GetDefaultPropertyMapParsers<T>(IXLRow headerRow) where T : class;
        bool TryParseProperty(PropertyInfo propertyInfo, string input, out object outVal);
        void FillCellBackground(ref IXLCell cell, bool isValid);
        void FillRowBackgroundWithValidationMessage(ref IXLRow row, bool isValid, List<string> validationMessages);
    }
}