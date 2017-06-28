namespace ExcelHelper
{
    using System.Collections.Generic;

    public class SheetParseResult<T>
    {
        public SheetParseResult()
        {
            this.ValidList = new List<T>();
        }
        public int TotalRecordCount { get; set; }
        public List<T> ValidList { get; set; }
    }
}
