using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;

using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;

namespace JDP.Remediation.Console
{
    public class GenerateSiteCollectionReport
    {
        public static void DoWork()
        {
            Logger.OpenLog("GenerateSiteCollectionReport");

            Logger.LogInfoMessage(String.Format("Preparing to generate report ..."), true);

            List<SiteEntity> sites = GetAllSites();

            GenerateReportFile(sites);

            Logger.LogInfoMessage(String.Format("Report completed at {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        private static void GenerateReportFile(List<SiteEntity> sites)
        {
            if (sites == null)
            {
                Logger.LogInfoMessage(String.Format("No site collections were found"), true);
                return;
            }

            string outputFileSpecFormat = "{0}-" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".txt";
            string outputFileSpec = String.Format(outputFileSpecFormat, "siteCollectionReport");

            using (System.IO.StreamWriter outputFile = new System.IO.StreamWriter(outputFileSpec))
            {
                outputFile.AutoFlush = true;

                foreach (SiteEntity site in sites)
                {
                    outputFile.WriteLine(site.Url);
                }
                outputFile.Close();

                Logger.LogMessage(String.Format("Report containing {0} sites has been saved to {1}", sites.Count, outputFileSpec), true);
            }
        }

        public static List<SiteEntity> GetAllSites()
        {
            try
            {
                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, Constants.PortalRootSiteUrl))
                {
                    // Lists all site collections across all web applications...
                    List<SiteEntity> sites = userContext.Web.SiteSearch();
                    return sites;
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("GetAllSites() failed: Error={0}", ex.Message), false);
                return null;
            }
        }
    }
}
