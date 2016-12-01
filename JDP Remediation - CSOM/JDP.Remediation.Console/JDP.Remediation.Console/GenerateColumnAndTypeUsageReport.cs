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
    public class GenerateColumnAndTypeUsageReport
    {
        public static string filePath = string.Empty;
        public static string outputPath = Environment.CurrentDirectory;
        public static string ComponentName = string.Empty;
        public static string ListId = string.Empty;
        public static string ListTitle = string.Empty;
        public static string ContentTypeOrCustomFieldId = string.Empty;
        public static string ContentTypeOrCustomFieldName = string.Empty;
        public static string WebUrl = string.Empty;

        public static bool headerContentType = false;

        private class ContentTypeSpec
        {
            public string Id;
            public string Name;
            public ContentTypeSpec(string id, string name)
            {
                this.Id = id.Trim();
                this.Name = name.Trim();
            }
        }

        private class SiteColumnSpec
        {
            public Guid Id;
            public string Name;

            public SiteColumnSpec(string id, string name)
            {
                this.Id = new Guid(id.Trim());
                this.Name = name.Trim();
            }
        }

        public static void DoWork()
        {
            string timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");
            string CTCFFileName = outputPath + @"\" + Constants.ContentTypeAndCustomFieldFileName + timeStamp + Constants.CSVExtension;
            Logger.OpenLog("GenerateColumnORFieldAndTypeUsageReport", timeStamp);
            if (!ShowInformation())
                return;
            Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);

            string contentTypesInputFileSpec = Environment.CurrentDirectory + "\\" + Constants.ContentTypeInput;
            List<ContentTypeSpec> contentTypes = new List<ContentTypeSpec>();
            if (System.IO.File.Exists(contentTypesInputFileSpec))
            {
                IEnumerable<InputContentTypeBase> objInputContentType = ImportCSV.ReadMatchingColumns<InputContentTypeBase>(contentTypesInputFileSpec, Constants.CsvDelimeter);
                Logger.LogInfoMessage(String.Format("Loaded {0} content type definitions ...", objInputContentType.Count()), true);
                foreach (InputContentTypeBase s in objInputContentType)
                {
                    contentTypes.Add(new ContentTypeSpec(s.ContentTypeID, s.ContentTypeName));
                }
            }
            else
                Logger.LogErrorMessage(String.Format("[GenerateColumnORFieldAndTypeUsageReport] Input file {0} is not available", contentTypesInputFileSpec), true);

            string customFieldsInputFileSpec = Environment.CurrentDirectory + "\\" + Constants.CustomFieldsInput;

            List<SiteColumnSpec> customFields = new List<SiteColumnSpec>();
            if (System.IO.File.Exists(customFieldsInputFileSpec))
            {
                IEnumerable<InputCustomFieldBase> objInputCustomField = ImportCSV.ReadMatchingColumns<InputCustomFieldBase>(customFieldsInputFileSpec, Constants.CsvDelimeter);
                Logger.LogInfoMessage(String.Format("Loaded {0} site column/custom field definitions ...", objInputCustomField.Count()), true);
                foreach (InputCustomFieldBase s in objInputCustomField)
                {
                    customFields.Add(new SiteColumnSpec(s.ID, s.Name));
                }
            }
            else
                Logger.LogErrorMessage(String.Format("[GenerateColumnORFieldAndTypeUsageReport] Input file {0} is not available", customFieldsInputFileSpec), true);

            if (!System.IO.File.Exists(CTCFFileName))
            {
                headerContentType = false;
            }
            else
                headerContentType = true;

            string inputFileSpec = Environment.CurrentDirectory + "\\" + Constants.UsageReport_SitesInputFileName;

            if (System.IO.File.Exists(inputFileSpec))
            {
                string[] siteUrls = Helper.ReadInputFile(inputFileSpec, false);
                Logger.LogInfoMessage(String.Format("Preparing to scan a total of {0} sites ...", siteUrls.Length), true);

                foreach (string siteUrl in siteUrls)
                {
                    ProcessSite(siteUrl, contentTypes, customFields, CTCFFileName);
                }
                Logger.LogSuccessMessage(String.Format("[GenerateColumnORFieldAndTypeUsageReport] Usage report is exported to the file {0} ", CTCFFileName), true);
            }
            else
                Logger.LogErrorMessage(String.Format("[GenerateColumnORFieldAndTypeUsageReport] Input file {0} is not available", inputFileSpec), true);

            Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        /// <summary>
        /// Executes all site collection-level reporting. 
        /// Performs special processing for the site collection, then processes all child webs.
        /// </summary>
        /// <param name="siteUrl">URL of the site collection to process</param>
        private static void ProcessSite(string siteUrl, List<ContentTypeSpec> contentTypes, List<SiteColumnSpec> siteColumns, string CTCFFileName)
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

                    // Execute processing that is common to all webs (including the root web).
                    ProcessWeb(root.Url, true, contentTypes, siteColumns, siteUrl, CTCFFileName);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[GenerateColumnORFieldAndTypeUsageReport] ProcessSite() failed for {0}: Error={1}", siteUrl, ex.Message), false);
                ExceptionCsv.WriteException(Constants.NotApplicable, siteUrl, Constants.NotApplicable, "ColumnORFieldAndTypeUsageReport", ex.Message, ex.ToString(), "ProcessSite()", ex.GetType().ToString(), String.Format("ProcessSite() failed for {0}: Error={1}", siteUrl, ex.Message));
            }
        }

        /// <summary>
        /// Executes all web-level transformations related to the custom solution. 
        /// Performs special processing for the root web, then recurses through all child webs.
        /// </summary>
        /// <param name="webUrl">Url of web to process</param>
        /// <param name="isRoot">True if this is the root web</param>
        private static void ProcessWeb(string webUrl, bool isRoot, List<ContentTypeSpec> contentTypes, List<SiteColumnSpec> siteColumns, string siteURL, string CTCFFileName)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Scanning Web: {0} ...", webUrl), false);

                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    Web web = userContext.Web;
                    userContext.Load(web);
                    userContext.ExecuteQuery();

                    ScanWeb(web, contentTypes, siteColumns, siteURL, CTCFFileName);

                    ScanLists(web, contentTypes, siteColumns, siteURL, CTCFFileName);

                    // Process all child webs
                    web.Context.Load(web.Webs);
                    web.Context.ExecuteQuery();
                    foreach (Web childWeb in web.Webs)
                    {
                        if (childWeb.Url.ToLower().Contains(siteURL.ToLower()))
                            ProcessWeb(childWeb.Url, false, contentTypes, siteColumns, siteURL, CTCFFileName);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[GenerateColumnORFieldAndTypeUsageReport] ProcessWeb() failed for {0}: Error={1}", webUrl, ex.Message), false);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "ColumnORFieldAndTypeUsageReport", ex.Message, ex.ToString(), "ProcessWeb()", ex.GetType().ToString(), String.Format("ProcessWeb() failed for {0}: Error={1}", webUrl, ex.Message));
            }
        }

        private static void ScanWeb(Web web, List<ContentTypeSpec> contentTypes, List<SiteColumnSpec> siteColumns, string siteURL, string CTCFFileName)
        {
            try
            {
                // scan the web for our custom site columns
                foreach (SiteColumnSpec sc in siteColumns)
                {
                    ReportSiteColumnUsage(web, sc.Id, sc.Name, siteURL, CTCFFileName);
                }

                // scan the web for our custom content types; either as is or as parents
                foreach (ContentTypeSpec ct in contentTypes)
                {
                    ReportContentTypeUsage(web, ct.Id, ct.Name, siteURL, CTCFFileName);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[GenerateColumnORFieldAndTypeUsageReport] ScanWeb() failed on {0}: Error={1}", web.Url, ex.Message), false);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "ColumnORFieldAndTypeUsageReport", ex.Message, ex.ToString(), "ScanWeb()", ex.GetType().ToString(), String.Format("ScanWeb() failed on {0}: Error={1}", web.Url, ex.Message));
            }
        }

        private static void ScanLists(Web web, List<ContentTypeSpec> contentTypes, List<SiteColumnSpec> siteColumns, string siteURL, string CTCFFileName)
        {
            try
            {
                ListCollection lists = web.Lists;
                web.Context.Load(lists);
                web.Context.ExecuteQuery();

                foreach (List list in lists)
                {
                    // We dont want to abort the entire collection if we get an exception on a given list (e.g., a list definition is missing)
                    try
                    {
                        // scan the list for our custom site columns
                        foreach (SiteColumnSpec sc in siteColumns)
                        {
                            ReportSiteColumnUsage(web, list, sc.Id, sc.Name, siteURL, CTCFFileName);
                        }

                        // scan the web for our custom content types; either as is or as parents
                        foreach (ContentTypeSpec ct in contentTypes)
                        {
                            ReportContentTypeUsage(web, list, ct.Id, ct.Name, siteURL, CTCFFileName);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogErrorMessage(String.Format("[GenerateColumnORFieldAndTypeUsageReport] ScanLists() failed on a list of {0}: Error={1}", web.Url, ex.Message), false);
                        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "ColumnORFieldAndTypeUsageReport", ex.Message, ex.ToString(), "ScanLists()", ex.GetType().ToString(), String.Format("ScanLists() failed on a list of {0}: Error={1}", web.Url, ex.Message));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ScanLists() failed for {0}: Error={1}", web.Url, ex.Message), false);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "ColumnORFieldAndTypeUsageReport", ex.Message, ex.ToString(), "ScanLists()", ex.GetType().ToString(), String.Format("ScanLists() failed for {0}: Error={1}", web.Url, ex.Message));
            }
        }

        private static void ReportSiteColumnUsage(Web web, Guid siteColumnId, string siteColumnName, string siteColUrl, string CTCFFileName)
        {
            ContentTypeCustomFieldOutput objCTCFOutput = new ContentTypeCustomFieldOutput();
            objCTCFOutput.ComponentName = Constants.CustomFields;
            objCTCFOutput.ListId = "N/A";
            objCTCFOutput.ListTitle = "N/A";
            objCTCFOutput.ContentTypeOrCustomFieldId = siteColumnId.ToString();
            objCTCFOutput.ContentTypeOrCustomFieldName = siteColumnName;
            objCTCFOutput.WebUrl = web.Url;
            objCTCFOutput.SiteCollection = siteColUrl;

            if (web.FieldExistsById(siteColumnId) == true)
            {
                Logger.LogSuccessMessage(String.Format("FOUND: Site Column/Custom Field [{1}] on WEB: {0}", web.Url, siteColumnName), true);
                FileUtility.WriteCsVintoFile(CTCFFileName, objCTCFOutput, ref headerContentType);
            }
        }
        private static void ReportSiteColumnUsage(Web web, List list, Guid siteColumnId, string siteColumnName, string siteColUrl, string CTCFFileName)
        {
            ContentTypeCustomFieldOutput objCTCFOutput1 = new ContentTypeCustomFieldOutput();
            objCTCFOutput1.ComponentName = Constants.CustomFields;
            objCTCFOutput1.ListId = list.Id.ToString();
            objCTCFOutput1.ListTitle = list.Title;
            objCTCFOutput1.ContentTypeOrCustomFieldId = siteColumnId.ToString();
            objCTCFOutput1.ContentTypeOrCustomFieldName = siteColumnName;
            objCTCFOutput1.WebUrl = web.Url;
            objCTCFOutput1.SiteCollection = siteColUrl;

            if (list.FieldExistsById(siteColumnId) == true)
            {
                Logger.LogSuccessMessage(String.Format("FOUND: Site Column/Custom Field [{2}] on LIST [{0}] of WEB: {1}", list.Title, web.Url, siteColumnName), true);
                FileUtility.WriteCsVintoFile(CTCFFileName, objCTCFOutput1, ref headerContentType);
            }
        }

        private static void ReportContentTypeUsage(Web web, string targetContentTypeId, string targetContentTypeName, string siteColUrl, string CTCFFileName)
        {
            try
            {
                ContentTypeCustomFieldOutput objCTCFOutput2 = new ContentTypeCustomFieldOutput();
                objCTCFOutput2.ComponentName = Constants.ContentTypes;
                objCTCFOutput2.ListId = "N/A";
                objCTCFOutput2.ListTitle = "N/A";
                objCTCFOutput2.ContentTypeOrCustomFieldId = targetContentTypeId.ToString();
                objCTCFOutput2.ContentTypeOrCustomFieldName = targetContentTypeName;
                objCTCFOutput2.WebUrl = web.Url;
                objCTCFOutput2.SiteCollection = siteColUrl;

                ContentTypeCollection ctCol = web.ContentTypes;
                web.Context.Load(ctCol);
                web.Context.ExecuteQuery();

                foreach (ContentType ct in ctCol)
                {
                    try
                    {
                        string contentTypeId = ct.Id.StringValue;
                        if (contentTypeId.StartsWith(targetContentTypeId, StringComparison.InvariantCultureIgnoreCase))
                        {
                            if (ct.Name.Equals(targetContentTypeName, StringComparison.InvariantCultureIgnoreCase))
                            {
                                Logger.LogSuccessMessage(String.Format("FOUND: Content Type [{1}] on WEB: {0}", web.Url, targetContentTypeName), true);
                                FileUtility.WriteCsVintoFile(CTCFFileName, objCTCFOutput2, ref headerContentType);
                            }
                            else
                            {
                                Logger.LogSuccessMessage(String.Format("FOUND: Child Content Type [{2}] of [{1}] on WEB: {0}", web.Url, targetContentTypeName, ct.Name), true);
                                FileUtility.WriteCsVintoFile(CTCFFileName, objCTCFOutput2, ref headerContentType);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogErrorMessage(String.Format("ReportContentTypeUsage() failed on a Content Type of WEB {0}: Error={1}", web.Url, ex.Message), false);
                        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "ColumnORFieldAndTypeUsageReport", ex.Message, ex.ToString(), "ReportContentTypeUsage()", ex.GetType().ToString(), String.Format("ReportContentTypeUsage() failed on a Content Type of WEB {0}: Error={1}", web.Url, ex.Message));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ReportContentTypeUsage() failed for WEB {0}: Error={1}", web.Url, ex.Message), false);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "ColumnORFieldAndTypeUsageReport", ex.Message, ex.ToString(), "ReportContentTypeUsage()", ex.GetType().ToString(), String.Format("ReportContentTypeUsage() failed for WEB {0}: Error={1}", web.Url, ex.Message));
            }
        }
        private static void ReportContentTypeUsage(Web web, List list, string targetContentTypeId, string targetContentTypeName, string siteColUrl, string CTCFFileName)
        {
            try
            {
                ContentTypeCustomFieldOutput objCTCFOutput3 = new ContentTypeCustomFieldOutput();
                objCTCFOutput3.ComponentName = Constants.ContentTypes;
                objCTCFOutput3.ListId = list.Id.ToString();
                objCTCFOutput3.ListTitle = list.Title;
                objCTCFOutput3.ContentTypeOrCustomFieldId = targetContentTypeId.ToString();
                objCTCFOutput3.ContentTypeOrCustomFieldName = targetContentTypeName;
                objCTCFOutput3.WebUrl = web.Url;
                objCTCFOutput3.SiteCollection = siteColUrl;

                list.EnsureProperty(l => l.ContentTypesEnabled);

                if (!list.ContentTypesEnabled)
                {
                    return;
                }

                ContentTypeCollection ctCol = list.ContentTypes;
                list.Context.Load(ctCol);
                list.Context.ExecuteQuery();

                foreach (ContentType ct in ctCol)
                {
                    try
                    {

                        string contentTypeId = ct.Id.StringValue;
                        if (contentTypeId.StartsWith(targetContentTypeId, StringComparison.InvariantCultureIgnoreCase))
                        {
                            if (ct.Name.Equals(targetContentTypeName, StringComparison.InvariantCultureIgnoreCase))
                            {
                                Logger.LogSuccessMessage(String.Format("FOUND: Content Type [{2}] on LIST [{0}] of WEB: {1}", list.Title, web.Url, targetContentTypeName), true);
                                FileUtility.WriteCsVintoFile(CTCFFileName, objCTCFOutput3, ref headerContentType);
                            }
                            else
                            {
                                Logger.LogSuccessMessage(String.Format("FOUND: Child Content Type [{3}] of [{2}] on LIST [{0}] of WEB: {1}", list.Title, web.Url, targetContentTypeName, ct.Name), true);
                                FileUtility.WriteCsVintoFile(CTCFFileName, objCTCFOutput3, ref headerContentType);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogErrorMessage(String.Format("ReportContentTypeUsage() failed on a Content Type of LIST [{0}] of WEB {1}: Error={2}", list.Title, web.Url, ex.Message), false);
                        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "ColumnORFieldAndTypeUsageReport", ex.Message, ex.ToString(), "ReportContentTypeUsage()", ex.GetType().ToString(), String.Format("ReportContentTypeUsage() failed on a Content Type of LIST [{0}] of WEB {1}: Error={2}", list.Title, web.Url, ex.Message));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ReportContentTypeUsage() failed for LIST [{0}] of WEB {1}: Error={2}", list.Title, web.Url, ex.Message), false);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "ColumnORFieldAndTypeUsageReport", ex.Message, ex.ToString(), "ReportContentTypeUsage()", ex.GetType().ToString(), String.Format("ReportContentTypeUsage() failed for LIST [{0}] of WEB {1}: Error={2}", list.Title, web.Url, ex.Message));
            }
        }

        private static bool ShowInformation()
        {
            bool doContinue = false;
            string option = string.Empty;
            System.Console.ForegroundColor = System.ConsoleColor.Yellow;
            System.Console.WriteLine("Input files <" + Constants.ContentTypeInput + ", " + Constants.CustomFieldsInput + ", " + Constants.UsageReport_SitesInputFileName + "> needs to be present in current working directory (where JDP.Remediation.Console.exe is present) to generate usage report. ");
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
