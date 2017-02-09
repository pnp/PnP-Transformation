using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;

using JDP.Remediation.Console.Common.Utilities;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Base;

namespace JDP.Remediation.Console
{
    /*
        This operation reads the SiteUsers property of each site collection and reports any instances of the specified principals.
        •	https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.web.siteusers.aspx 
        
        The SiteUsers property maps to the SiteUsers list, which basically acts as an intermediate representation of all principals who, 
        at one time or another, have been granted explicitly granted access to the site collection.

        Behavior of the Site Users list:
        •	When direct access has been explicitly granted for a principal:
            o	Permission is granted for the principal
            o	The principal is also added to the list
        •	When direct access has been explicitly removed for a principal:
            o	Permission is remove for the principal
            o	HOWEVER, the principal does not get removed from the list

        In order to remove the principal from the list, one must also:
        •	visit either of the following pages:
            o	the People and Groups page (/_layouts/15/groups.aspx)
            o	the All People page (/_layouts/15/people.aspx?membershipGroup=0)
        •	click on the Principal to remove
        •	select Delete user from site collection
    */
    public class GenerateSecurityGroupReport
    {
        private static string csvOutputFileSpec = String.Empty;
        private static bool csvOutputFileHasHeader = false;

        public static void DoWork()
        {
            string timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");

            csvOutputFileSpec = Environment.CurrentDirectory + "\\GenerateSecurityGroupReport-" + timeStamp + Constants.CSVExtension;
            csvOutputFileHasHeader = System.IO.File.Exists(csvOutputFileSpec);

            Logger.OpenLog("GenerateSecurityGroupReport", timeStamp);
            Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);

            string[] securityGroups = new string[0];
            string securityGroupsInputFileSpec = Environment.CurrentDirectory + "\\" + Constants.SecurityGroupsInputFileName;
            if (!System.IO.File.Exists(securityGroupsInputFileSpec))
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage(String.Format("Input file {0} is not available", securityGroupsInputFileSpec), true);
                Logger.LogInfoMessage(String.Format("Scan aborted {0}", DateTime.Now.ToString()), true);
                Logger.CloseLog();
                System.Console.ResetColor();
                return;
            }
            securityGroups = Helper.ReadInputFile(securityGroupsInputFileSpec, false);
            if (securityGroups.Length == 0)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage(String.Format("Input file {0} is empty", securityGroupsInputFileSpec), true);
                Logger.LogInfoMessage(String.Format("Scan aborted {0}", DateTime.Now.ToString()), true);
                Logger.CloseLog();
                System.Console.ResetColor();
                return;
            }
            Logger.LogInfoMessage(String.Format("Loaded a total of {0} security groups ...", securityGroups.Length), true);

            string inputFileSpec = Environment.CurrentDirectory + "\\" + Constants.UsageReport_SitesInputFileName;
            if (!System.IO.File.Exists(inputFileSpec))
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage(String.Format("Input file {0} is not available", inputFileSpec), true);
                Logger.LogInfoMessage(String.Format("Scan aborted {0}", DateTime.Now.ToString()), true);
                Logger.CloseLog();
                System.Console.ResetColor();
                return;
            }

            string[] siteUrls = Helper.ReadInputFile(inputFileSpec, false);
            Logger.LogInfoMessage(String.Format("Preparing to scan a total of {0} sites ...", siteUrls.Length), true);

            foreach (string siteUrl in siteUrls)
            {
                ProcessSite(siteUrl, securityGroups);
            }
            Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        private static void ProcessSite(string siteUrl, string [] securityGroups)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Processing Site: {0} ...", siteUrl), true);

                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, siteUrl))
                {
                    Web root = userContext.Site.RootWeb;
                    userContext.Load(root);
                    userContext.ExecuteQuery();

                    // Looking for only the Security Groups present on the site
                    var groups = userContext.LoadQuery(root.SiteUsers.Where(u => (u.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.SecurityGroup)));
                    userContext.ExecuteQuery();

                    Logger.LogInfoMessage(String.Format("There are [{0}] Security Groups to check...", groups.Count()), false);
                    foreach (User g in groups)
                    {
                        try
                        {
                            // The User.Title property is generally in "domain\alias" format for on-prem/SPO-D: (e.g., orion\joeUser)
                            // other values you might see: 
                            //  Everyone
                            //  All Users (windows)
                            //  NT AUTHORITY\authenticated users
                            Logger.LogInfoMessage(String.Format("Checking Security Group [{0}] for significance...", g.Title), false);

                            // We could also compare the User.LoginName property, but we would need the SID in order to build a string in the following format:
                            //  c:0+.w|s-1-5-21-1485757101-1923125180-2349192791-514
                            if (securityGroups.Any(x => x.Equals(g.Title, StringComparison.InvariantCultureIgnoreCase)))
                            {
                                GenerateSecurityGroupOutput csvObject = new GenerateSecurityGroupOutput();
                                csvObject.SecurityGroupName = g.Title;
                                csvObject.SiteCollectionUrl = siteUrl;

                                Logger.LogSuccessMessage(String.Format("Significant Security Group [{0}] found on site [{1}]", g.Title, siteUrl), true);
                                FileUtility.WriteCsVintoFile(csvOutputFileSpec, csvObject, ref csvOutputFileHasHeader);
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.LogErrorMessage(String.Format("ProcessSite() failed for Security Group [{0}]: Error={1}", g.Title, ex.Message), false);
                            ExceptionCsv.WriteException(
                                "N/A", siteUrl, "N/A",
                                "SecurityGroup",
                                ex.Message, ex.ToString(), "ProcessSite", ex.GetType().ToString(),
                                String.Format("ProcessSite() failed for Security Group [{0}]", g.Title)
                                );
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ProcessSite() failed for site [{0}]: Error={1}", siteUrl, ex.Message), false);
                ExceptionCsv.WriteException(
                    "N/A", siteUrl, "N/A",
                    "SecurityGroup",
                    ex.Message, ex.ToString(), "ProcessSite", ex.GetType().ToString(),
                    String.Format("ProcessSite() failed for site [{0}]", siteUrl)
                    );
            }
        }
    }
}
