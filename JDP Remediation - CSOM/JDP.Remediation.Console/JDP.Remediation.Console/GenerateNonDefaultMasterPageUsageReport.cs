using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;
using JDP.Remediation.Console.Common.Utilities;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Base;

namespace JDP.Remediation.Console
{
    public class GenerateNonDefaultMasterPageUsageReport
    {
        /// <summary>
        /// This method reports on the usage of custom master pages.
        /// 
        /// General Approach:
        /// - disable site feature: V4VisualUpgrade
        /// -  note: the feature deactivator processes MPs of all child webs
        /// - process webs
        /// -  delete custom hidden list: Lists/RedirectURL
        /// -  disable web feature: Redirect Url List
        /// -  disable web features: custom Master Pages (6)
        /// -  delete custom master page files
        /// </summary>

        public static string filePath = string.Empty;
        public static string outputPath = Environment.CurrentDirectory;
        public static bool headermasterPage = false;

        public static void DoWork()
        {
            string timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");
            string NonDefMasterFileName = outputPath + @"\" + Constants.NonDefaultMasterPageFileName + timeStamp + Constants.CSVExtension;
            Logger.OpenLog("GenerateNonDefaultMasterPageUsageReport", timeStamp);
            if (!ShowInformation())
                return;
            Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);

            if (!System.IO.File.Exists(NonDefMasterFileName))
            {
                headermasterPage = false;
            }
            else
                headermasterPage = true;

            string inputFileSpec = Environment.CurrentDirectory + "\\" + Constants.UsageReport_SitesInputFileName;
           
            if (System.IO.File.Exists(inputFileSpec))
            {
                string[]  siteUrls = Helper.ReadInputFile(inputFileSpec, false);
                Logger.LogInfoMessage(String.Format("Preparing to scan a total of {0} sites ...", siteUrls.Length), true);

                foreach (string siteUrl in siteUrls)
                {
                    ProcessSite(siteUrl, NonDefMasterFileName);
                }
                Logger.LogSuccessMessage(String.Format("[GenerateNonDefaultMasterPageUsageReport] Usage report is exported to the file {0}", NonDefMasterFileName), true);
            }
            else
                Logger.LogErrorMessage(String.Format("[GenerateNonDefaultMasterPageUsageReport] Input file {0} is not available", inputFileSpec), true);

            Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        /// <summary>
        /// Executes all site collection-level reporting. 
        /// Performs special processing for the site collection, then processes all child webs.
        /// </summary>
        /// <param name="siteUrl">URL of the site collection to process</param>
        private static void ProcessSite(string siteUrl, string NonDefMasterFileName)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Scanning Site: {0} ...", siteUrl), true);

                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, siteUrl))
                {
                    Site site = userContext.Site;
                    Web root = userContext.Site.RootWeb;
                    userContext.Load(site);
                    userContext.Load(root);
                    userContext.ExecuteQuery();

                    // Execute processing that is unique to the site and its root web

                    // Execute processing that is common to all webs (including the root web).
                    ProcessWeb(root.Url, true, siteUrl, NonDefMasterFileName);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[GenerateNonDefaultMasterPageUsageReport] ProcessSite() failed for {0}: Error={1}", siteUrl, ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, siteUrl, Constants.NotApplicable, "NonDefaultMastePageUsageReport", ex.Message, ex.ToString(), "ProcessSite()", ex.GetType().ToString(), String.Format("ProcessSite() failed for {0}: Error={1}", siteUrl, ex.Message));
            }
        }

        /// <summary>
        /// Executes all web-level transformations related to the custom solution. 
        /// Performs special processing for the root web, then recurses through all child webs.
        /// </summary>
        /// <param name="webUrl">Url of web to process</param>
        /// <param name="isRoot">True if this is the root web</param>
        private static void ProcessWeb(string webUrl, bool isRoot, string SiteURL, string NonDefMasterFileName)
        {
            try
            {
                string masterUrl = string.Empty;
                string customMasterurl = string.Empty;
                string WebUrl = string.Empty;
                bool IsMasterUrl = false;
                bool IsCustomMasterUrl = false;

                NonDefaultMasterpageOutput objMasterPageOutput = new NonDefaultMasterpageOutput();

                Logger.LogInfoMessage(String.Format("Scanning Web: {0} ...", webUrl), false);

                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    Web web = userContext.Web;
                    userContext.Load(web);
                    userContext.Load(web.AllProperties);
                    userContext.ExecuteQuery();

                    if (web.MasterUrl.ToLowerInvariant().Contains("seattle.master") == false)
                    {
                        Logger.LogSuccessMessage(String.Format("FOUND: System Master Page setting (Prop=MasterUrl) of web {0} is {1}", web.Url, web.MasterUrl), true);
                        IsMasterUrl = true;
                    }
                    else
                        IsMasterUrl = false;

                    if (web.CustomMasterUrl.ToLowerInvariant().Contains("seattle.master") == false)
                    {
                        Logger.LogSuccessMessage(String.Format("FOUND: Site Master Page setting (Prop=CustomMasterUrl) of web {0} is {1}", web.Url, web.CustomMasterUrl), true);
                        IsCustomMasterUrl = true;
                    }
                    else
                        IsCustomMasterUrl = false;

                    objMasterPageOutput.MasterUrl = web.MasterUrl;
                    objMasterPageOutput.CustomMasterUrl = web.CustomMasterUrl;
                    objMasterPageOutput.WebUrl = web.Url;
                    objMasterPageOutput.SiteCollection = SiteURL;

                    if (IsMasterUrl || IsCustomMasterUrl)
                        FileUtility.WriteCsVintoFile(NonDefMasterFileName, objMasterPageOutput, ref headermasterPage);

                    // Process all child webs
                    web.Context.Load(web.Webs);
                    web.Context.ExecuteQuery();
                    foreach (Web childWeb in web.Webs)
                    {
                        if (childWeb.Url.ToLower().Contains(SiteURL.ToLower()))
                            ProcessWeb(childWeb.Url, false, SiteURL, NonDefMasterFileName);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[GenerateNonDefaultMasterPageUsageReport] ProcessWeb() failed for {0}: Error={1}", webUrl, ex.Message), false);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "NonDefaultMastePageUsageReport", ex.Message, ex.ToString(), "ProcessWeb()", ex.GetType().ToString(), String.Format("ProcessWeb() failed for {0}: Error={1}", webUrl, ex.Message));
            }
        }

        private static bool ShowInformation()
        {
            bool doContinue = false;
            string option = string.Empty;
            System.Console.ForegroundColor = System.ConsoleColor.Yellow;
            System.Console.WriteLine(Constants.UsageReport_SitesInputFileName + " file needs to be present in current working directory (where JDP.Remediation.Console.exe is present) to generate usage report. ");
            System.Console.ResetColor();
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Press 'y' to proceed further. Press any key to go for Self Service Report Menu.");
            System.Console.ResetColor();
            option = System.Console.ReadLine().ToLower();
            if (option.Equals("y", StringComparison.OrdinalIgnoreCase))
                doContinue = true;
            return doContinue;
        }
    }
}
