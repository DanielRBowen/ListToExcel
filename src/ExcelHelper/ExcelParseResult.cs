namespace ExcelHelper
{
    using System;
    using System.Collections.Generic;
    using ClosedXML.Excel;
    public class ExcelParseResult<T> : IDisposable where T : class
    {
        public ExcelParseResult()
        {
            this.ValidList = new List<T>();
        }
        public int TotalRecordCount { get; set; }
        public List<T> ValidList { get; set; }
        public XLWorkbook XlWorkbook { get; set; }

        public void Dispose()
        {
            this.XlWorkbook?.Dispose();
        }
    }
}