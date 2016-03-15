using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO.Compression;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;

namespace PeoplePickerRemediation.Console.Common.Utilities
{
    public static class CommonUtility
    {
        
        /// <summary>
        // Excel/CSV Cell CharacterLimit. According to Microsoft's documentation: 
        // https://support.office.com/en-us/article/Excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
        // Excel cannot read more than 32767 characters in a single cell
        // Total number of characters that a cell can contain: 32,767 characters
        // This function checks if the value of a column is more than 32,767 characters, and if it finds any it splits the data into rows so as to save it in csv/excel
        /// </summary>
        /// <param name="stringToSplit"></param>
        /// <returns></returns>
        public static IEnumerable<string> SplitToLines(string stringToSplit)
        {
            stringToSplit = stringToSplit.Trim();
            var lines = new List<string>();

            //If stringToSplit is Blank then return Constants.NotApplicable
            if (stringToSplit.Length != 0)
            {
                while (stringToSplit.Length > 0)
                {
                    if (stringToSplit.Length <= Constants.CharacterLimitForCsvCell)
                    {
                        lines.Add(stringToSplit);
                        break;
                    }

                    int indexOfLastSpaceInLine = stringToSplit.Substring(0, Constants.CharacterLimitForCsvCell).LastIndexOf(' ');

                    lines.Add(stringToSplit.Substring(0, indexOfLastSpaceInLine >= 0 ? indexOfLastSpaceInLine : Constants.CharacterLimitForCsvCell).Trim());

                    stringToSplit = stringToSplit.Substring(indexOfLastSpaceInLine >= 0 ? indexOfLastSpaceInLine + 1 : Constants.CharacterLimitForCsvCell);
                }
            }
            else
            {
                //If stringToSplit is Blank then return Constants.NotApplicable
                lines.Add(Constants.NotApplicable);
            }

            return lines.ToArray();
        }

        /// <summary>
        /// Returns a string Url by combining the WebAplicationUrl and the Input Url (Which Can be - SiteCollectionUrl, WebUrl and PageUrl)
        /// </summary>
        /// <param name="WebApplicationUrl"></param>
        /// <param name="Url"></param>
        /// <returns></returns>
        public static string GetUrl(string WebApplicationUrl, string Url)
        {
            string newUrl = string.Empty;

            if (WebApplicationUrl.EndsWith("/"))
            {
                if (Url.StartsWith("/"))
                    newUrl = Regex.Match(WebApplicationUrl, "^(.*).{1}", RegexOptions.IgnoreCase).Groups[1].Value + Url;
                else
                    newUrl = WebApplicationUrl + Url;
            }
            else if (Url.StartsWith("/"))
                newUrl = WebApplicationUrl + Url;
            else
                newUrl = WebApplicationUrl + "/" + Url;

            return newUrl;
        }
    }
}
