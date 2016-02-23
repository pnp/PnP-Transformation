using System;
using System.Globalization;

namespace Transformation.PowerShell.Base
{
    public interface IFileExport
    {
    }

    public class Elementbase
    {
        public string WebApplication { get; set; }
        public string SiteCollection { get; set; }
        public string WebUrl { get; set; }
    }

    public class Inputbase
    {
        public string WebApplicationUrl { get; set; }
        public string SiteCollectionUrl { get; set; }
        public string WebUrl { get; set; }
    }

    public class CustomFieldBase
    {
        public string SiteCollectionUrl { get; set; }
    }

    public class ExportAttribute : Attribute
    {
        public int FieldOrder { get; set; }
        public int FieldLength { get; set; }
    }

    public class CsvColumnAttribute : Attribute
    {
        internal const int McDefaultFieldIndex = Int32.MaxValue;

        public CsvColumnAttribute()
        {
            Name = "";
            FieldIndex = McDefaultFieldIndex;
            CanBeNull = true;
            NumberStyle = NumberStyles.Any;
            OutputFormat = "G";
        }

        public CsvColumnAttribute(
            string name,
            int fieldIndex,
            bool canBeNull,
            string outputFormat,
            NumberStyles numberStyle,
            int charLength)
        {
            Name = name;
            FieldIndex = fieldIndex;
            CanBeNull = canBeNull;
            NumberStyle = numberStyle;
            OutputFormat = outputFormat;

            CharLength = charLength;
        }

        public string Name { get; set; }
        public bool CanBeNull { get; set; }
        public int FieldIndex { get; set; }
        public NumberStyles NumberStyle { get; set; }
        public string OutputFormat { get; set; }
        public int CharLength { get; set; }
    }
}