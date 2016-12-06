using JDP.Remediation.Console.Common.Base;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Utilities;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDP.Remediation.Console
{
    public class ReplaceMasterPage
    {
        public static string filePath = string.Empty;
        public static string outputPath = string.Empty;
        public static string timeStamp = string.Empty;

        public static void DoWork()
        {
            timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");
            bool processInputFile = false;
            bool processWebUrl = false;
            bool replaceMasterUrl = false;
            bool replaceCustomMasterUrl = false;
            bool replaceBothMaserUrls = false;
            string masterPageInputFile = string.Empty;
            string webUrl = string.Empty;
            try
            {
                outputPath = Environment.CurrentDirectory;
                Logger.OpenLog("ReplaceMasterPage", timeStamp);

                if (!ReadInputOptions(ref processInputFile, ref processWebUrl))
                {
                    System.Console.ForegroundColor = System.ConsoleColor.Red;
                    Logger.LogErrorMessage("Invalid option selected or Exit option is selected. Operation aborted!");
                    System.Console.ResetColor();
                    return;
                }

                if (processInputFile)
                {
                    if (!ReadInputFile(ref masterPageInputFile))
                    {
                        System.Console.ForegroundColor = System.ConsoleColor.Red;
                        Logger.LogErrorMessage("MasterPage input file is not valid or available. So, Operation aborted!");
                        Logger.LogErrorMessage("Please enter path like: E.g. C:\\<Working Directory>\\<InputFile>.csv");
                        System.Console.ResetColor();
                        return;
                    }
                    if (!ReadMasterUrlReplaceOptions(ref replaceMasterUrl, ref replaceCustomMasterUrl, ref replaceBothMaserUrls))
                    {
                        System.Console.ForegroundColor = System.ConsoleColor.Red;
                        Logger.LogErrorMessage("Invalid option selected. Operation aborted!");
                        System.Console.ResetColor();
                        return;
                    }
                    ProcessInputFile(masterPageInputFile, replaceMasterUrl, replaceCustomMasterUrl, replaceBothMaserUrls, outputPath);
                }

                if (processWebUrl)
                {
                    System.Console.ForegroundColor = System.ConsoleColor.Cyan;
                    Logger.LogMessage("Enter WebUrl to replace MasterPage: ");
                    System.Console.ResetColor();
                    webUrl = System.Console.ReadLine();
                    if (string.IsNullOrEmpty(webUrl))
                    {
                        Logger.LogErrorMessage("[ReplaceMasterPage: DoWork]WebUrl should not be empty or null. Operation aborted...", true);
                        return;
                    }
                    if (!ReadMasterUrlReplaceOptions(ref replaceMasterUrl, ref replaceCustomMasterUrl, ref replaceBothMaserUrls))
                    {
                        System.Console.ForegroundColor = System.ConsoleColor.Red;
                        Logger.LogErrorMessage("Invalid option selected. Operation aborted!");
                        System.Console.ResetColor();
                        return;
                    }
                    if (ProcessWebUrl(webUrl, null, replaceMasterUrl, replaceCustomMasterUrl, replaceBothMaserUrls))
                    {
                        System.Console.ForegroundColor = System.ConsoleColor.Green;
                        Logger.LogSuccessMessage("[ReplaceMasterPage: DoWork] Successfully processed given WebUrl and output file is present in the path: "
                            + outputPath, true);
                        System.Console.ResetColor();
                    }
                    else
                    {
                        Logger.LogErrorMessage("Replacing Custom Master Page with OOB Master Page is failed for the site " + webUrl);
                    }
                }

            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[ReplaceMasterPage: DoWork] failed: Error={0}", ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "ReplaceMasterPage", ex.Message, ex.ToString(), "ReplaceMasterPage: DoWork()", ex.GetType().ToString());
            }
            Logger.CloseLog();
        }

        private static bool ReadInputOptions(ref bool processInputFile, ref bool processWebUrl)
        {
            string processOption = string.Empty;
            System.Console.ForegroundColor = System.ConsoleColor.White;
            Logger.LogMessage("Type 1, 2 or 3 and press Enter to select the respective operation to execute:");
            Logger.LogMessage("1. Process with Input File");
            Logger.LogMessage("2. Process for WebUrl");
            Logger.LogMessage("3. Exit to Transformation Menu");
            System.Console.ResetColor();
            processOption = System.Console.ReadLine();

            if (processOption.Equals("1"))
                processInputFile = true;
            else if (processOption.Equals("2"))
                processWebUrl = true;
            else if (processOption.Equals("3"))
                return false;
            else
                return false;

            return true;
        }

        private static bool ReadMasterUrlReplaceOptions(ref bool replaceMasterUrl, ref bool replaceCustomMasterUrl, ref bool replaceBothMasterUrls)
        {
            string processOption = string.Empty;
            System.Console.ForegroundColor = System.ConsoleColor.White;
            Logger.LogMessage("Type 1, 2 or 3 and press Enter to select the respective operation to execute:");
            Logger.LogMessage("1. Replace MasterUrl");
            Logger.LogMessage("2. Replace Custom MasterUrl");
            Logger.LogMessage("3. Replace Both MasterUrls");
            System.Console.ResetColor();
            processOption = System.Console.ReadLine();

            if (processOption.Equals("1"))
                replaceMasterUrl = true;
            else if (processOption.Equals("2"))
                replaceCustomMasterUrl = true;
            else if (processOption.Equals("3"))
                replaceBothMasterUrls = true;
            else
                return false;

            return true;
        }

        private static bool ReplaceMasterUrl(ClientContext ctx, string oobMasterPageUrl, string serverRelativeUrl,
            bool replaceMasterUrl, bool replaceCustomMasterUrl, bool replaceBothMasterUrls)
        {
            bool replaceMasterUrlStatus = false;
            oobMasterPageUrl = @"/_catalogs/masterpage/" + oobMasterPageUrl;
            Web web;
            try
            {
                web = ctx.Web;
                ctx.Load(web);

                if (replaceMasterUrl)
                {
                    //Update the system master.
                    web.MasterUrl = serverRelativeUrl + oobMasterPageUrl;
                }
                else if (replaceCustomMasterUrl)
                {
                    web.CustomMasterUrl = serverRelativeUrl + oobMasterPageUrl;
                }
                else if (replaceBothMasterUrls)
                {
                    web.MasterUrl = serverRelativeUrl + oobMasterPageUrl;
                    web.CustomMasterUrl = serverRelativeUrl + oobMasterPageUrl;
                }
                web.Update();
                ctx.ExecuteQuery();
                replaceMasterUrlStatus = true;
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[ReplaceMasterpage: ReplaceMasterUrl]. Exception Message: " + ex.Message, true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "ReplaceMasterPage", ex.Message, ex.ToString(), "ReplaceMasterUrl()", ex.GetType().ToString());
            }
            return replaceMasterUrlStatus;
        }
        private static bool ReadInputFile(ref string masterPageInputFile)
        {
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            Logger.LogMessage("Enter Complete Input File Path of MasterPage Report Either Pre-Scan OR Discovery Report.");
            System.Console.ResetColor();
            masterPageInputFile = System.Console.ReadLine();
            Logger.LogMessage("Entered Input File of MasterPage Data " + masterPageInputFile, false);
            if (string.IsNullOrEmpty(masterPageInputFile) || !System.IO.File.Exists(masterPageInputFile))
                return false;
            return true;
        }

        private static void ReadCustomOOBMasterPages(ref string customMasterPage, ref string oobMasterPage)
        {
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            Logger.LogMessage("Enter Custom MasterPage to be replaced with OOB MasterPAge along with extension (E.g, contoso.master):");
            System.Console.ResetColor();
            customMasterPage = System.Console.ReadLine().ToLower();
            if (string.IsNullOrEmpty(customMasterPage))
            {
                Logger.LogErrorMessage("[ReplaceMasterPage: DoWork]customMasterPage should not be empty or null. Operation aborted...", true);
                return;
            }
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            Logger.LogMessage("Enter OOB MasterPage to replace Custom MasterPage along with extension (E.g, seattle.master):");
            System.Console.ResetColor();
            oobMasterPage = System.Console.ReadLine().ToLower();
            if (string.IsNullOrEmpty(oobMasterPage))
            {
                Logger.LogErrorMessage("[ReplaceMasterPage: DoWork]oobMasterPage should not be empty or null. Operation aborted...", true);
                return;
            }
        }

        private static void ProcessInputFile(string masterPageInputFile, bool replaceMasterUrl, bool replaceCustomMasterUrl, bool replaceBothMaserUrls, string outputDirectory)
        {
            DataTable dtMasterPagesInput = new DataTable();
            try
            {
                dtMasterPagesInput = ImportCSV.Read(masterPageInputFile, Constants.CsvDelimeter);

                List<string> lstWebUrls = dtMasterPagesInput.AsEnumerable()
                                                        .Select(r => r.Field<string>("WebUrl"))
                                                        .ToList();
                lstWebUrls = lstWebUrls.Distinct().ToList();
                foreach (string webUrl in lstWebUrls)
                {
                    string webApplicationUrl = string.Empty;
                    try
                    {
                        Logger.LogInfoMessage("[ReplaceMasterPage: ProcessInputFile] Processing the Site: " + webUrl, true);
                        if (ProcessWebUrl(webUrl, null, replaceMasterUrl, replaceCustomMasterUrl, replaceBothMaserUrls))
                        {
                            Logger.LogSuccessMessage("[ReplaceMasterPage:ProcessInputFile]successfully replaced master page for the site " + webUrl + "  and output file is present in the path: " + outputDirectory, true);
                        }
                        else
                            Logger.LogErrorMessage("[ReplaceMasterPage: ProcessInputFile] Replacing Custom Master Page with OOB Master Page is failed for the site " + webUrl);
                    }
                    catch (Exception ex)
                    {
                        Logger.LogErrorMessage("[ReplaceMasterPage: ProcessInputFile]. Exception Message: " + ex.Message, true);
                        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "ReplaceMasterPage", ex.Message, ex.ToString(), "ProcessInputFile()", ex.GetType().ToString());
                    }
                }
                System.Console.ForegroundColor = System.ConsoleColor.Green;
                Logger.LogSuccessMessage("[ReplaceMasterPage: DoWork] Successfully processed all sites and output file is present in the path: "
                    + outputPath, true);
                System.Console.ResetColor();
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[ReplaceMasterpage: ProcessInputFile]. Exception Message: " + ex.Message, true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "ReplaceMasterPage", ex.Message, ex.ToString(), "ProcessInputFile()", ex.GetType().ToString());
            }
            finally
            {
                dtMasterPagesInput.Dispose();
            }
        }

        public static bool ProcessWebUrl(string webUrl, DataRow drMasterPage, bool replaceMasterUrl, bool replaceCustomMasterUrl, bool replaceBothMaserUrls)
        {
            bool result = false;
            string customMasterPage = string.Empty;
            string oobMasterPage = string.Empty;
            string serverRelativeUrl = string.Empty;
            Site site;
            Web web;
            bool header = false;

            try
            {
                ReadCustomOOBMasterPages(ref customMasterPage, ref oobMasterPage);
                if (string.IsNullOrEmpty(customMasterPage) || string.IsNullOrEmpty(oobMasterPage))
                {
                    result = false;
                }
                else
                {
                    if (!customMasterPage.EndsWith(".master"))
                    {
                        customMasterPage = string.Empty;
                        Logger.LogErrorMessage("Invalid extension of Custom Master Page.");
                        result = false;
                    }
                    if (!oobMasterPage.EndsWith(".master"))
                    {
                        oobMasterPage = string.Empty;
                        Logger.LogErrorMessage("Invalid extension of OOB Master Page.");
                        result = false;
                    }
                    else
                    {
                        using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                        {
                            try
                            {
                                site = userContext.Site;
                                web = userContext.Web;
                                userContext.Load(site);
                                userContext.Load(web);
                                userContext.ExecuteQuery();
                                serverRelativeUrl = site.ServerRelativeUrl;

                                if (web.MasterUrl.ToLower().Contains(customMasterPage) || web.CustomMasterUrl.ToLower().Contains(customMasterPage))
                                {
                                    result = ReplaceMasterUrl(userContext, oobMasterPage, serverRelativeUrl, replaceMasterUrl, replaceCustomMasterUrl, replaceBothMaserUrls);
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.LogErrorMessage("[ReplaceMasterpage: ProcessWebUrl]. Exception Message: " + ex.Message, true);
                                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "ReplaceMasterPage", ex.Message, ex.ToString(), "ProcessWebUrl()", ex.GetType().ToString());
                            }
                        }
                    }
                }
                ReplaceMasterPageStatusBase objOutputBase = new ReplaceMasterPageStatusBase();
                objOutputBase.WebApplication = Constants.NotApplicable;
                objOutputBase.SiteCollection = Constants.NotApplicable;
                objOutputBase.WebUrl = webUrl;
                objOutputBase.ExecutionDateTime = DateTime.Now.ToString();

                if (!string.IsNullOrEmpty(customMasterPage))
                    objOutputBase.CustomMasterPageUrl = customMasterPage;
                else
                    objOutputBase.CustomMasterPageUrl = Constants.NotApplicable;

                if (!string.IsNullOrEmpty(oobMasterPage))
                    objOutputBase.OOTBMasterPageUrl = oobMasterPage;
                else
                    objOutputBase.OOTBMasterPageUrl = Constants.NotApplicable;

                if (result)
                {
                    objOutputBase.Status = Constants.Success;
                }
                else
                    objOutputBase.Status = Constants.Failure;

                if (!System.IO.File.Exists(outputPath + @"\" + Constants.ReplaceMasterPageFileName + timeStamp + Constants.CSVExtension))
                {
                    header = false;
                }
                else
                    header = true;
                FileUtility.WriteCsVintoFile(outputPath + @"\" + Constants.ReplaceMasterPageFileName + timeStamp + Constants.CSVExtension, objOutputBase, ref header);
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[ReplaceMasterpage: ProcessWebUrl]. Exception Message: " + ex.Message, true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "ReplaceMasterPage", ex.Message, ex.ToString(), "ProcessWebUrl()", ex.GetType().ToString());
            }
            finally
            {
                site = null;
                web = null;
            }
            return result;
        }

    }
}