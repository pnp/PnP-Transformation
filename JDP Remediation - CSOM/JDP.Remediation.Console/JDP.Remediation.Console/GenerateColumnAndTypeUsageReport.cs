using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;

namespace JDP.Remediation.Console
{
    public class GenerateColumnAndTypeUsageReport
    {
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
            Logger.OpenLog("GenerateColumnAndTypeUsageReport");
            Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);

            string contentTypesInputFileSpec = Environment.CurrentDirectory + "\\" + Constants.UsageReport_ContentTypesInputFileName;
            string[] contentTypeRows = Helper.ReadInputFile(contentTypesInputFileSpec, true);
            Logger.LogInfoMessage(String.Format("Loaded {0} content type definitions ...", contentTypeRows.Length), true);
            List<ContentTypeSpec> contentTypes = new List<ContentTypeSpec>();
            foreach (string s in contentTypeRows)
            {
                string[] arr = s.Split(',');
                contentTypes.Add(new ContentTypeSpec(arr[0], arr[1]));
            }            

            string siteColumnsInputFileSpec = Environment.CurrentDirectory + "\\" + Constants.UsageReport_SiteColumnsInputFileName;
            string[] siteColumnRows = Helper.ReadInputFile(siteColumnsInputFileSpec, true);
            Logger.LogInfoMessage(String.Format("Loaded {0} site column definitions ...", siteColumnRows.Length), true);
            List<SiteColumnSpec> siteColumns = new List<SiteColumnSpec>();
            foreach (string s in siteColumnRows)
            {
                string[] arr = s.Split(',');
                siteColumns.Add(new SiteColumnSpec(arr[0], arr[1]));
            }

            string inputFileSpec = Environment.CurrentDirectory + "\\" + Constants.UsageReport_SitesInputFileName;
            string[] siteUrls = Helper.ReadInputFile(inputFileSpec, false);
            Logger.LogInfoMessage(String.Format("Preparing to scan a total of {0} sites ...", siteUrls.Length), true);

            foreach (string siteUrl in siteUrls)
            {
                ProcessSite(siteUrl, contentTypes, siteColumns);
            }

            Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        /// <summary>
        /// Executes all site collection-level reporting. 
        /// Performs special processing for the site collection, then processes all child webs.
        /// </summary>
        /// <param name="siteUrl">URL of the site collection to process</param>
        private static void ProcessSite(string siteUrl, List<ContentTypeSpec> contentTypes, List<SiteColumnSpec> siteColumns)
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
                    ProcessWeb(root.Url, true, contentTypes, siteColumns);
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
        private static void ProcessWeb(string webUrl, bool isRoot, List<ContentTypeSpec> contentTypes, List<SiteColumnSpec> siteColumns)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Scanning Web: {0} ...", webUrl), false);

                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    Web web = userContext.Web;
                    userContext.Load(web);
                    userContext.ExecuteQuery();

                    ScanWeb(web, contentTypes, siteColumns);

                    ScanLists(web, contentTypes, siteColumns);

                    // Process all child webs
                    web.Context.Load(web.Webs);
                    web.Context.ExecuteQuery();
                    foreach (Web childWeb in web.Webs)
                    {
                        ProcessWeb(childWeb.Url, false, contentTypes, siteColumns);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ProcessWeb() failed for {0}: Error={1}", webUrl, ex.Message), false);
            }
        }

        private static void ScanWeb(Web web, List<ContentTypeSpec> contentTypes, List<SiteColumnSpec> siteColumns)
        {
            try
            { 
                // scan the web for our custom site columns
                foreach (SiteColumnSpec sc in siteColumns)
                {
                    ReportSiteColumnUsage(web, sc.Id, sc.Name);
                }

                // scan the web for our custom content types; either as is or as parents
                foreach (ContentTypeSpec ct in contentTypes)
                {
                    ReportContentTypeUsage(web, ct.Id, ct.Name);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ScanWeb() failed on {0}: Error={1}", web.Url, ex.Message), false);
            }
        }

        private static void ScanLists(Web web, List<ContentTypeSpec> contentTypes, List<SiteColumnSpec> siteColumns)
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
                            ReportSiteColumnUsage(web, list, sc.Id, sc.Name);
                        }

                        // scan the web for our custom content types; either as is or as parents
                        foreach (ContentTypeSpec ct in contentTypes)
                        {
                            ReportContentTypeUsage(web, list, ct.Id, ct.Name);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogErrorMessage(String.Format("ScanLists() failed on a list of {0}: Error={1}", web.Url, ex.Message), false);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ScanLists() failed for {0}: Error={1}", web.Url, ex.Message), false);
            }
        }

        private static void ReportSiteColumnUsage(Web web, Guid siteColumnId, string siteColumnName)
        {
            if (web.FieldExistsById(siteColumnId) == true)
            {
                Logger.LogSuccessMessage(String.Format("FOUND: Site Column [{1}] on WEB: {0}", web.Url, siteColumnName), true);
            }
        }
        private static void ReportSiteColumnUsage(Web web, List list, Guid siteColumnId, string siteColumnName)
        {
            if (list.FieldExistsById(siteColumnId) == true)
            {
                Logger.LogSuccessMessage(String.Format("FOUND: Site Column [{2}] on LIST [{0}] of WEB: {1}", list.Title, web.Url, siteColumnName), true);
            }
        }

        private static void ReportContentTypeUsage(Web web, string targetContentTypeId, string targetContentTypeName)
        {
            try
            {
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
                            }
                            else
                            {
                                Logger.LogSuccessMessage(String.Format("FOUND: Child Content Type [{2}] of [{1}] on WEB: {0}", web.Url, targetContentTypeName, ct.Name), true);
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogErrorMessage(String.Format("ReportContentTypeUsage() failed on a Content Type of WEB {0}: Error={1}", web.Url, ex.Message), false);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ReportContentTypeUsage() failed for WEB {0}: Error={1}", web.Url, ex.Message), false);
            }
        }
        private static void ReportContentTypeUsage(Web web, List list, string targetContentTypeId, string targetContentTypeName)
        {
            try
            {
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
                            }
                            else
                            {
                                Logger.LogSuccessMessage(String.Format("FOUND: Child Content Type [{3}] of [{2}] on LIST [{0}] of WEB: {1}", list.Title, web.Url, targetContentTypeName, ct.Name), true);
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogErrorMessage(String.Format("ReportContentTypeUsage() failed on a Content Type of LIST [{0}] of WEB {1}: Error={2}", list.Title, web.Url, ex.Message), false);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ReportContentTypeUsage() failed for LIST [{0}] of WEB {1}: Error={2}", list.Title, web.Url, ex.Message), false);
            }
        }
    }
}
