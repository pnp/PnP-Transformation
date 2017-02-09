using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Base;
using JDP.Remediation.Console.Common.Utilities;

namespace JDP.Remediation.Console
{
    public class ManageMaintenanceBanners
    {
        private enum BannerOperation
        {
            None = 0,
            Add = 1,
            Remove = 2
        }

        private static string csvOutputFileSpec = String.Empty;
        private static bool csvOutputFileHasHeader = false;

        public static void DoWork()
        {
            string timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");

            csvOutputFileSpec = Environment.CurrentDirectory + "\\ManageMaintenanceBanners-" + timeStamp + Constants.CSVExtension;
            csvOutputFileHasHeader = System.IO.File.Exists(csvOutputFileSpec);

            Logger.OpenLog("ManageMaintenanceBanners", timeStamp);
            Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);

            BannerOperation operationToPerform = BannerOperation.None;

            if (!ReadInputOptions(ref operationToPerform))
            {
                System.Console.ForegroundColor = System.ConsoleColor.Yellow;
                Logger.LogInfoMessage("Operation aborted by user.", true);
                Logger.LogInfoMessage(String.Format("Scan aborted {0}", DateTime.Now.ToString()), true);
                Logger.CloseLog();
                System.Console.ResetColor();
                return;
            }

            string cdnAbsoluteUrl = String.Empty;
            if (operationToPerform == BannerOperation.Add)
            {
                if (!ReadCdnUrl(ref cdnAbsoluteUrl))
                {
                    System.Console.ForegroundColor = System.ConsoleColor.Red;
                    Logger.LogErrorMessage("CDN Library Folder Url was not specified.");
                    Logger.LogInfoMessage(String.Format("Scan aborted {0}", DateTime.Now.ToString()), true);
                    Logger.CloseLog();
                    System.Console.ResetColor();
                    return;
                }
            }

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
            Logger.LogInfoMessage(String.Format("Preparing to process a total of {0} sites ...", siteUrls.Length), true);

            foreach (string siteUrl in siteUrls)
            {
                ProcessSite(siteUrl, operationToPerform, cdnAbsoluteUrl);
            }

            Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        private static void ProcessSite(string siteUrl, BannerOperation operationToPerform, string cdnAbsoluteUrl)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Processing Site: {0} ...", siteUrl), true);

                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, siteUrl))
                {
                    Site site = userContext.Site;
                    Web root = userContext.Site.RootWeb;
                    userContext.Load(site);
                    userContext.Load(root);
                    userContext.ExecuteQuery();

                    switch (operationToPerform)
                    {
                        case BannerOperation.Add:
                            // We remove any existing items before adding the new one
                            DeleteJsLinks(userContext, site);
                            AddJsLinks(userContext, site, cdnAbsoluteUrl);
                            break;

                        case BannerOperation.Remove:
                            DeleteJsLinks(userContext, site);
                            break;

                        case BannerOperation.None:
                        default:
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ProcessSite() failed for site [{0}]: Error={1}", siteUrl, ex.Message), false);
                ExceptionCsv.WriteException(
                    "N/A", siteUrl, "N/A",
                    "MaintenanceBanner",
                    ex.Message, ex.ToString(), "ProcessSite", ex.GetType().ToString(),
                    String.Format("ProcessSite() failed for site [{0}]", siteUrl)
                    );
            }
        }

        private static void AddJsLinks(ClientContext context, Site site, string cdnAbsoluteUrl)
        {
            ManageMaintenanceBannersOutput csvObject = new ManageMaintenanceBannersOutput();
            csvObject.BannerOperation = "Add Banner";
            csvObject.SiteCollectionUrl = site.Url;
            csvObject.ScriptLinkName = Constants.ScriptLinkDescription;
            csvObject.ScriptLinkFile = "N/A";

            try
            {
                string revision = Guid.NewGuid().ToString().Replace("-", "");

                string embedJsLink = string.Format("{0}/{1}?rev={2}", cdnAbsoluteUrl, Constants.EmbedJsFileName, revision);
                csvObject.ScriptLinkFile = embedJsLink;

                StringBuilder scripts = new StringBuilder(@"
                    var headID = document.getElementsByTagName('head')[0]; 
                    var");

                scripts.AppendFormat(@"
                    newScript = document.createElement('script');
                    newScript.type = 'text/javascript';
                    newScript.src = '{0}';
                    headID.appendChild(newScript);", embedJsLink);

                string scriptBlock = scripts.ToString();

                var existingActions = site.UserCustomActions;
                context.Load(existingActions);
                context.ExecuteQuery();
                var actions = existingActions.ToArray();

                Logger.LogInfoMessage(String.Format("Adding ScriptLink Action [{0}] to site [{1}]...", Constants.ScriptLinkDescription, site.Url), false);
                Logger.LogInfoMessage(String.Format("-embedJsLink = [{0}]", embedJsLink), false);

                var newAction = existingActions.Add();
                newAction.Description = Constants.ScriptLinkDescription;
                newAction.Location = Constants.ScriptLinkLocation;

                newAction.ScriptBlock = scriptBlock;
                newAction.Update();
                context.Load(site, s => s.UserCustomActions);
                context.ExecuteQuery();

                csvObject.Status = Constants.Success;
                FileUtility.WriteCsVintoFile(csvOutputFileSpec, csvObject, ref csvOutputFileHasHeader);

                Logger.LogSuccessMessage(String.Format("ScriptLink Action [{0}] added to site [{1}]", Constants.ScriptLinkDescription, site.Url), false);
            }
            catch (Exception ex)
            {
                csvObject.Status = Constants.Failure;
                FileUtility.WriteCsVintoFile(csvOutputFileSpec, csvObject, ref csvOutputFileHasHeader);

                Logger.LogErrorMessage(String.Format("AddJsLink() failed to add ScriptLink Action [{0}] to site [{1}]: Error={2}", Constants.ScriptLinkDescription, site.Url, ex.Message), false);
                ExceptionCsv.WriteException(
                    "N/A", site.Url, "N/A",
                    "MaintenanceBanner",
                    ex.Message, ex.ToString(), "AddJsLinks", ex.GetType().ToString(),
                    String.Format("AddJsLink() failed to add ScriptLink Action [{0}] to site [{1}]", Constants.ScriptLinkDescription, site.Url)
                    );
            }
        }

        private static void DeleteJsLinks(ClientContext context, Site site)
        {
            var existingActions = site.UserCustomActions;
            context.Load(existingActions);
            context.ExecuteQuery();

            var actions = existingActions.ToArray();
            if (actions.Count() == 0)
            {
                Logger.LogInfoMessage(String.Format("No ScriptLink Action [{0}] items are present on site [{1}]...", Constants.ScriptLinkDescription, site.Url), false);
                return;
            }

            Logger.LogInfoMessage(String.Format("Removing all existing ScriptLink Action [{0}] items from site [{1}]...", Constants.ScriptLinkDescription, site.Url), false);

            foreach (var action in actions)
            {
                if (action.Description.Equals(Constants.ScriptLinkDescription, StringComparison.InvariantCultureIgnoreCase) &&
                    action.Location.Equals(Constants.ScriptLinkLocation, StringComparison.InvariantCultureIgnoreCase)
                    )
                {
                    ManageMaintenanceBannersOutput csvObject = new ManageMaintenanceBannersOutput();
                    csvObject.BannerOperation = "Remove Banner";
                    csvObject.SiteCollectionUrl = site.Url;
                    csvObject.ScriptLinkName = Constants.ScriptLinkDescription;
                    csvObject.ScriptLinkFile = "N/A";

                    try
                    {
                        action.DeleteObject();
                        context.ExecuteQuery();

                        csvObject.Status = Constants.Success;
                        FileUtility.WriteCsVintoFile(csvOutputFileSpec, csvObject, ref csvOutputFileHasHeader);

                        Logger.LogInfoMessage(String.Format("ScriptLink Action [{0}] removed from site [{1}]", Constants.ScriptLinkDescription, site.Url), false);
                    }
                    catch (Exception ex)
                    {
                        csvObject.Status = Constants.Failure;
                        FileUtility.WriteCsVintoFile(csvOutputFileSpec, csvObject, ref csvOutputFileHasHeader);

                        Logger.LogErrorMessage(String.Format("DeleteJsLinks() failed to remove ScriptLink Action [{0}] from site [{1}]: Error={2}", Constants.ScriptLinkDescription, site.Url, ex.Message), false);
                        ExceptionCsv.WriteException(
                            "N/A", site.Url, "N/A",
                            "MaintenanceBanner",
                            ex.Message, ex.ToString(), "DeleteJsLinks", ex.GetType().ToString(),
                            String.Format("DeleteJsLinks() failed to remove ScriptLink Action [{0}] from site [{1}]", Constants.ScriptLinkDescription, site.Url)
                            );
                    }
                }
            }
        }

        private static bool ReadInputOptions(ref BannerOperation operationToPerform)
        {
            bool operationSelected = false;

            string processOption = string.Empty;
            System.Console.ForegroundColor = System.ConsoleColor.White;
            Logger.LogMessage("Please type an operation number and press [Enter] to execute the specified operation:");
            Logger.LogMessage("1. Add Maintenance Banner to Sites");
            Logger.LogMessage("2. Remove Maintenance Banner from Sites");
            Logger.LogMessage("3. Exit to Transformation Menu");
            System.Console.ResetColor();
            processOption = System.Console.ReadLine();

            switch (processOption)
            {
                case "1":
                    operationToPerform = BannerOperation.Add;
                    operationSelected = true;
                    Logger.LogInfoMessage(String.Format("Selected Operation = {0} Banners", operationToPerform.ToString()), false);
                    break;

                case "2":
                    operationToPerform = BannerOperation.Remove;
                    operationSelected = true;
                    Logger.LogInfoMessage(String.Format("Selected Operation = {0} Banners", operationToPerform.ToString()), false);
                    break;

                case "3":
                default:
                    operationToPerform = BannerOperation.None;
                    operationSelected = false;
                    break;
            }

            return operationSelected;
        }

        private static bool ReadCdnUrl(ref string cdnAbsoluteUrl)
        {
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            Logger.LogMessage(String.Format("Please enter the absolute Url of the CDN Library Folder that contains the [{0}] banner script file: ", Constants.EmbedJsFileName));
            Logger.LogMessage(String.Format("-Example: https://portal.contoso.com/style library/cdn/scripts"));
            System.Console.ForegroundColor = System.ConsoleColor.Yellow;
            Logger.LogMessage(String.Format("Note: if you have not already done so, please upload the file to your CDN Library Folder and ensure that it is published before continuing."));
            System.Console.ResetColor();
            cdnAbsoluteUrl = System.Console.ReadLine();

            Logger.LogInfoMessage(String.Format("Specified CDN Library Folder Url = {0}", cdnAbsoluteUrl), false);

            if (cdnAbsoluteUrl.EndsWith("/"))
            {
                cdnAbsoluteUrl = cdnAbsoluteUrl.TrimEnd(new char[] { '/' });
            }

            return !String.IsNullOrEmpty(cdnAbsoluteUrl);
        }

    }
}
