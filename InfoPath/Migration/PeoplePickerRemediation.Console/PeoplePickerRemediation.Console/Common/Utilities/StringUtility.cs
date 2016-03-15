using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PeoplePickerRemediation.Console.Common.Utilities
{
    public static class StringUtility
    {
        public static string Escape(string s)
        {
            if (s != null)
            {
                if (s.Contains(Constants.Quote))
                    s = s.Replace(Constants.Quote, Constants.EscapedQuote);

                if (s.IndexOfAny(Constants.CharactersThatMustBeQuoted) > -1)
                    s = Constants.Quote + s + Constants.Quote;

                return s;
            }
            else
            {
                return string.Empty;
            }
        }

        public static string Unescape(string s)
        {
            if (s.StartsWith(Constants.Quote) && s.EndsWith(Constants.Quote))
            {
                s = s.Substring(1, s.Length - 2);

                if (s.Contains(Constants.EscapedQuote))
                    s = s.Replace(Constants.EscapedQuote, Constants.Quote);
            }

            return s;
        }

        public static bool CompareCommaDelimitedString(string strValue, string CommaDelimitedStringValues)
        {
            try
            {
                if (CommaDelimitedStringValues.Trim() == string.Empty)
                {
                    return false;
                }

                if (CommaDelimitedStringValues.ToLower().Contains(Constants.CsvDelimeter))
                {
                    string[] delimiterChars = { Constants.CsvDelimeter };
                    string[] strExcludeWebApp = CommaDelimitedStringValues.Split(delimiterChars,
                        StringSplitOptions.RemoveEmptyEntries);
                    for (int index = 0; index < strExcludeWebApp.Length; index++)
                    {
                        string str = strExcludeWebApp[index];
                        if (strValue.ToLower().Equals(str.ToLower()))
                        {
                            return true;
                        }
                    }

                    return false;
                }
                if (strValue.ToLower().Equals(CommaDelimitedStringValues.ToLower()))
                {
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        public static DateTime ConvertDateTime(string datetime)
        {
            try
            {
                return Convert.ToDateTime(datetime);
            }
            catch
            {
                return DateTime.Now;
            }
        }

        public static string TotalExcelutionTime(TimeSpan _TimeSpan)
        {
            // Format and display the TimeSpan value.
            string elapsedTime = "Total Time of Execution in HH:MM:SS:MS Format " + String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                _TimeSpan.Hours, _TimeSpan.Minutes, _TimeSpan.Seconds,
                _TimeSpan.Milliseconds / 10);

            return elapsedTime;
        }

    }
}
