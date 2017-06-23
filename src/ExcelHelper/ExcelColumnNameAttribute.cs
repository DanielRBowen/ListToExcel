namespace ExcelHelper
{
    using System;

    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelColumnNameAttribute : Attribute
    {
        public string ColumnName { get; private set; }
        public ExcelColumnNameAttribute(string columnName)
        {
            this.ColumnName = columnName;
        }
    }
}
