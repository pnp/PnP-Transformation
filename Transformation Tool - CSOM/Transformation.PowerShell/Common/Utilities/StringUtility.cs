using System;

namespace Transformation.PowerShell.Common.Utilities
{
    public static class StringUtility
    {
        public static string Escape(string s)
        {
            if (s.Contains(Constants.Quote))
                s = s.Replace(Constants.Quote, Constants.EscapedQuote);

            if (s.IndexOfAny(Constants.CharactersThatMustBeQuoted) > -1)
                s = Constants.Quote + s + Constants.Quote;

            return s;
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

        public static bool ExcludeWebApplication(string webApplicationUrl, string excludeWebApp)
        {
            try
            {
                if (excludeWebApp.Trim() == string.Empty)
                {
                    return false;
                }

                if (excludeWebApp.ToLower().Contains(Constants.CsvDelimeter))
                {
                    string[] delimiterChars = {Constants.CsvDelimeter};
                    string[] strExcludeWebApp = excludeWebApp.Split(delimiterChars,
                        StringSplitOptions.RemoveEmptyEntries);
                    for (int index = 0; index < strExcludeWebApp.Length; index++)
                    {
                        string str = strExcludeWebApp[index];
                        if (webApplicationUrl.ToLower().Contains(str.ToLower()))
                        {
                            return true;
                        }
                    }

                    return false;
                }
                if (webApplicationUrl.ToLower().Contains(excludeWebApp.ToLower()))
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
    }
}