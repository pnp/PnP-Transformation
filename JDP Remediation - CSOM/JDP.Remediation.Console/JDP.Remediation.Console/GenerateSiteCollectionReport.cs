using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;

using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using JDP.Remediation.Console.Common.Utilities;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Base;

namespace JDP.Remediation.Console
{
    public class GenerateSiteCollectionReport
    {
        public static string filePath = string.Empty;
        public static string outputPath = Environment.CurrentDirectory;
        public static bool headerSiteCollection = false;
        public static void DoWork()
        {
            string GenSiteColFileName = outputPath + @"\" + Constants.GenSiteCollectionFileName + DateTime.Now.ToString("yyyyMMdd_hhmmss") + Constants.CSVExtension;

            string timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");
            Logger.OpenLog("GenerateSiteCollectionReport", timeStamp);

            List<SiteEntity> sites = GetAllSites();

            GenerateReportFile(sites, GenSiteColFileName);

            Logger.LogInfoMessage(String.Format("Report completed at {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        private static void GenerateReportFile(List<SiteEntity> sites, string GenSiteColFileName)
        {
            if (!System.IO.File.Exists(GenSiteColFileName))
            {
                headerSiteCollection = false;
            }
            else
                headerSiteCollection = true;

            GenerateSiteCollectionnOutput objGenSiteColOutput = new GenerateSiteCollectionnOutput();

            if (sites == null)
            {
                Logger.LogInfoMessage(String.Format("No site collections were found"), true);
                return;
            }
           
            //Text File
            string outputFileSpecFormat = "{0}" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".txt";
            string outputFileSpec = String.Format(outputFileSpecFormat, Constants.GenSiteCollectionFileName);

            using (System.IO.StreamWriter outputFile = new System.IO.StreamWriter(outputFileSpec))
            {
                outputFile.AutoFlush = true;

                foreach (SiteEntity site in sites)
                {
                    //CSV File
                    outputFile.WriteLine(site.Url);
                    objGenSiteColOutput.SiteCollectionUrl = site.Url;
                    FileUtility.WriteCsVintoFile(GenSiteColFileName, objGenSiteColOutput, ref headerSiteCollection);
                }
                outputFile.Close();
                Logger.LogSuccessMessage(String.Format("[GenerateSiteCollectionReport] GenerateReportFile: Usage report is exported to the file {0}", GenSiteColFileName), true);
                Logger.LogMessage(String.Format("[GenerateSiteCollectionReport] GenerateReportFile: Report containing {0} sites has been saved to {1} (Text Format) and " + GenSiteColFileName + " (CSV Format) ", sites.Count, outputFileSpec), true);
            }
        }

        public static List<SiteEntity> GetAllSites()
        {
            string url = string.Empty;
            try
            {
                System.Console.ForegroundColor = System.ConsoleColor.Cyan;
                System.Console.WriteLine("Enter any Web/Site Url from the farm for which you want all site collections report:");
                System.Console.ResetColor();
                url = System.Console.ReadLine();
                if (string.IsNullOrEmpty(url))
                {
                    Logger.LogErrorMessage(String.Format("[GenerateSiteCollectionReport] GetAllSites() Url should not be null or empty"), true);
                    ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "SiteCollectionReport", "Url should not be null or empty", "Url should not be null or empty", "GetAllSites()", "NullReferenceException", String.Format("GetAllSites() failed: Error={0}", "Url should not be null or empty"));
                    return null;
                }
                Logger.LogInfoMessage(String.Format("Preparing to generate report ..."), true);
                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, url))
                {
                    // Lists all site collections across all web applications...
                    List<SiteEntity> sites = userContext.Web.SiteSearch();
                    return sites;
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[GenerateSiteCollectionReport] GetAllSites() failed: Error={0}", ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "SiteCollectionReport", ex.Message, ex.ToString(), "GetAllSites()", ex.GetType().ToString(), String.Format("GetAllSites() failed: Error={0}", ex.Message));
                return null;
            }
        }

        public static List<SiteEntity> GetAllSites(string webAppUrl)
        {
            try
            {
                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webAppUrl))
                {
                    // Lists all site collections across all web applications...
                    List<SiteEntity> sites = userContext.Web.SiteSearch();
                    return sites;
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[GenerateSiteCollectionReport] GetAllSites() failed: Error={0}", ex.Message), false);
                ExceptionCsv.WriteException(webAppUrl, Constants.NotApplicable, Constants.NotApplicable, "SiteCollectionReport", ex.Message, ex.ToString(), "GetAllSites()", ex.GetType().ToString(), String.Format("GetAllSites() failed: Error={0}", ex.Message));
                return null;
            }
        }

        /// <summary>
        /// Testing Method Only
        /// </summary>
        /// <returns></returns>
        //public static List<string> GetAllSites2()
        //{
        //    List<string> URL = new List<string>();
        //    try
        //    {
        //        using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, Constants.PortalRootSiteUrl))
        //        {
        //            // Lists all site collections across all web applications...
        //            var spoTenant = new Microsoft.Online.SharePoint.TenantAdministration.Tenant(userContext);
        //            var getSite = spoTenant.GetSiteProperties(0, true);
        //            userContext.Load(getSite);
        //            userContext.ExecuteQuery();
        //            foreach (var site in getSite)
        //            {
        //                URL.Add(site.Url);

        //            }
        //            return URL;

        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "SiteCollectionReport", ex.Message, ex.ToString(), "GetAllSites2()", ex.GetType().ToString());
        //        return URL;
        //    }

        //}
        ///// <summary>
        ///// Testing Method Only
        ///// </summary>
        ///// <returns></returns>
        //public static List<SiteEntity> GetAllSites3()
        //{
        //    try
        //    {
        //        using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, Constants.PortalRootSiteUrl))
        //        {
        //            // Lists all site collections across all web applications...
        //            //WebTemplateCollection wtc= userContext.Web.GetAvailableWebTemplates(1033, true);

        //            WebTemplateCollection wtc = userContext.Site.GetWebTemplates(1033, 0);


        //            userContext.Load(wtc);
        //            userContext.ExecuteQuery();
        //            List<string> template = new List<string>();

        //            foreach (WebTemplate wt in wtc)
        //            {
        //                template.Add(wt.Name);
        //            }

        //            List<SiteEntity> sites = new List<SiteEntity>();
        //            foreach (string templateKeyword in template)
        //            {
        //                try
        //                {
        //                    sites.AddRange(userContext.Web.SiteSearch(templateKeyword, true));
        //                }
        //                catch (Exception ex)
        //                {
        //                    Logger.LogErrorMessage(String.Format("[GenerateSiteCollectionReport] GetAllSites() failed: Error={0}", ex.Message), false);
        //                    ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "SiteCollectionReport", ex.Message, ex.ToString(), "GetAllSites3()", ex.GetType().ToString(), String.Format("GetAllSites3() failed: Error={0}", ex.Message));
        //                }


        //            }

        //            List<string> siteurl = new List<string>();
        //            foreach (SiteEntity site in sites)
        //            {
        //                if (!siteurl.Contains(site.Url))
        //                {
        //                    siteurl.Add(site.Url);
        //                }

        //            }
                    
        //            return sites;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Logger.LogErrorMessage(String.Format("[GenerateSiteCollectionReport] GetAllSites() failed: Error={0}", ex.Message), false);
        //        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "SiteCollectionReport", ex.Message, ex.ToString(), "GetAllSites3()", ex.GetType().ToString(), String.Format("GetAllSites3() failed: Error={0}", ex.Message));
        //        return null;
        //    }
        //}
    }
}
