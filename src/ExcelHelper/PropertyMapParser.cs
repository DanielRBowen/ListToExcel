namespace ExcelHelper
{
    using System.Reflection;

    public delegate bool TryGetProperty(PropertyInfo propertyInfo, string input, out object outVal);
    public class PropertyMapParser
    {
        public PropertyInfo ObjectPropertyInfo { get; set; }
        public int ExcelColumnIndex { get; set; }
        public TryGetProperty TryGetProperty { get; set; }
    }
}