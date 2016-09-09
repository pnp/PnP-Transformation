using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;

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
        public static void DoWork()
        {
            Logger.OpenLog("GenerateNonDefaultMasterPageUsageReport");
            Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);

            string inputFileSpec = Environment.CurrentDirectory + "\\" + Constants.UsageReport_SitesInputFileName;
            string[] siteUrls = Helper.ReadInputFile(inputFileSpec, false);
            Logger.LogInfoMessage(String.Format("Preparing to scan a total of {0} sites ...", siteUrls.Length), true);

            foreach (string siteUrl in siteUrls)
            {
                ProcessSite(siteUrl);
            }

            Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        /// <summary>
        /// Executes all site collection-level reporting. 
        /// Performs special processing for the site collection, then processes all child webs.
        /// </summary>
        /// <param name="siteUrl">URL of the site collection to process</param>
        private static void ProcessSite(string siteUrl)
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
                    ProcessWeb(root.Url, true);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ProcessSite() failed for {0}: Error={1}", siteUrl, ex.Message), false);
            }
        }

        /// <summary>
        /// Executes all web-level transformations related to the custom solution. 
        /// Performs special processing for the root web, then recurses through all child webs.
        /// </summary>
        /// <param name="webUrl">Url of web to process</param>
        /// <param name="isRoot">True if this is the root web</param>
        private static void ProcessWeb(string webUrl, bool isRoot)
        {
            try
            {
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
                    }
                    if (web.CustomMasterUrl.ToLowerInvariant().Contains("seattle.master") == false)
                    {
                        Logger.LogSuccessMessage(String.Format("FOUND: Site Master Page setting (Prop=CustomMasterUrl) of web {0} is {1}", web.Url, web.CustomMasterUrl), true);
                    }

                    // Process all child webs
                    web.Context.Load(web.Webs);
                    web.Context.ExecuteQuery();
                    foreach (Web childWeb in web.Webs)
                    {
                        ProcessWeb(childWeb.Url, false);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ProcessWeb() failed for {0}: Error={1}", webUrl, ex.Message), false);
            }
        }
    }
}
