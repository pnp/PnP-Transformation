using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PeoplePickerRemediation.Console
{
    /// <summary>
    /// This class holds all constants used by the program.  No code
    /// </summary>
    public class Constants
    {
        public const string Quote = "\"";
        public const string EscapedQuote = "\"\"";
        public static char[] CharactersThatMustBeQuoted = { ',', '"', '\n' };
        public static readonly string CsvDelimeter = ",";
        public static readonly string NotApplicable = "N/A";
        public static readonly long ExceptionFileSizeinKb = 4096; //4 MB
        public static readonly bool Logging = true;
        public static readonly string Exception = "Exception.csv";
        public static readonly string TraceLogFileSuffix = "TraceLog";
        public static readonly int MaxListRecordsToExportCountForPeriodic = 50000;
        public static readonly int MaxListRecordsToExportCount = 10000;
        ///// <summary>
        //// Excel/CSV Cell CharacterLimit. According to Microsoft's documentation: 
        //// https://support.office.com/en-us/article/Excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
        //// Excel cannot read more than 32767 characters in a single cell
        //// Total number of characters that a cell can contain: 32,767 characters
        ///// </summary>
        public static readonly int CharacterLimitForCsvCell = 32758;
        public static readonly long OutputFileSizeinKb = 8192; //8 MB

        public static readonly string PeopplePickerReportOutput = "PeoplePickerDataFix_" + DateTime.Now.ToString("dd_MM_yyyy_hh_mm_ss") + ".csv";
        public static readonly string SuccessStatus = "Success";
        public static readonly string NoUpdateRequired = "NoupdateRequired";
        public static readonly string ErrorStatus = "Error";
        public static readonly string OutPutreportSeparator = ";";
    }
}
