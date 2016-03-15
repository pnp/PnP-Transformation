using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace PeoplePickerRemediation.Console
{
    public class Helper
    {

        public class MasterPageInfo
        {
            public string MasterPageUrl;
            public bool InheritMaster;
            public string CustomMasterPageUrl;
            public bool InheritCustomMaster;
        }

        public static ClientContext CreateClientContextBasedOnAuthMode(bool useAppModel, string domain, string username, SecureString password, string siteUrl)
        {
            if (useAppModel)
            {
                return CreateAppOnlyClientContext(siteUrl);
            }
            else
            {
                return CreateAuthenticatedUserContext(domain, username, password, siteUrl);
            }
        }

        private static ClientContext CreateAppOnlyClientContext(string siteUrl)
        {
            var parentSiteUri = new Uri(siteUrl);
            string realm = TokenHelper.GetRealmFromTargetUrl(parentSiteUri);
            var token = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, parentSiteUri.Authority, realm).AccessToken;
            var clientContext = TokenHelper.GetClientContextWithAccessToken(parentSiteUri.ToString(), token);

            return clientContext;
        }
        private static ClientContext CreateAuthenticatedUserContext(string domain, string username, SecureString password, string siteUrl)
        {
            ClientContext userContext = new ClientContext(siteUrl);
            if (String.IsNullOrEmpty(domain))
            {
                // use o365 authentication (SPO-MT or vNext)
                userContext.Credentials = new SharePointOnlineCredentials(username, password);
            }
            else
            {
                // use Windows authentication (SPO-D or On-Prem) 
                userContext.Credentials = new NetworkCredential(username, password, domain);
            }

            return userContext;
        }

        /// <summary>
        /// Creates a Secure String
        /// </summary>
        /// <param name="data">string to be converted</param>
        /// <returns>secure string instance</returns>
        public static SecureString CreateSecureString(string data)
        {
            if (data == null || string.IsNullOrEmpty(data))
            {
                return null;
            }

            SecureString secureString = new SecureString();

            char[] charArray = data.ToCharArray();

            foreach (char ch in charArray)
            {
                secureString.AppendChar(ch);
            }

            return secureString;
        }






        public static string[] ReadInputFile(string inputFileSpec, bool hasHeader)
        {
            try
            {
                if (hasHeader == true)
                {
                    // remove the header row from the resulting string array...
                    List<string> temp = new List<string>(System.IO.File.ReadAllLines(inputFileSpec));
                    temp.RemoveAt(0);
                    return temp.ToArray();
                }
                else
                {
                    return System.IO.File.ReadAllLines(inputFileSpec);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ReadInputFile() failed for {0}: Error={1}", inputFileSpec, ex.Message), true);
                return new string[0];
            }
        }

        // Parses the CSV input line and returns a string array where each entry corresponds to a field parsed from the line
        // Strips double quotes from those fields that contain commas: e.g., 123,"abc,def",456
        public static string[] ParseInputLine(string inputLine)
        {
            try
            {
                if (inputLine.Contains('"'))
                {
                    int pos1 = inputLine.IndexOf('"');
                    int pos2 = inputLine.IndexOf('"', pos1 + 1);

                    string left = inputLine.Substring(0, pos1);
                    left = left.Trim(new char[] { '"', ',' });

                    string center = inputLine.Substring(pos1, pos2 - pos1);
                    center = center.TrimStart(new char[] { '"', ',' });

                    string right = inputLine.Substring(pos2);
                    right = right.TrimStart(new char[] { '"', ',' });

                    List<string> result = new List<string>();
                    string[] temp = null;

                    temp = ParseInputLine(left);
                    if (temp.Length > 0) result.AddRange(temp);

                    result.Add(center);

                    temp = ParseInputLine(right);
                    if (temp.Length > 0) result.AddRange(temp);

                    return result.ToArray();
                }
                else
                {
                    inputLine = inputLine.Trim(new char[] { ',' });
                    if (String.IsNullOrEmpty(inputLine))
                    {
                        return new string[0];
                    }
                    return inputLine.Split(',');
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ParseInputLine() failed for [{0}]: Error={1}", inputLine, ex.Message), false);
                return new string[0];
            }
        }
    }
}
