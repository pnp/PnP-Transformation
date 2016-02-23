using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Microsoft.SharePoint.Client;
using Transformation.PowerShell.Base;
using Transformation.PowerShell.Common;
using Transformation.PowerShell.Common.CSV;
using Transformation.PowerShell.PageLayout;
using Microsoft.SharePoint.Client.Publishing;
using Transformation.PowerShell.Common.Utilities;

namespace Transformation.PowerShell.PageLayouts
{
    public class PageLayoutHelper
    {
        /// <summary>
        /// This function will return true if Web = Publishing Web
        /// Ref. Links - 
        /// http://sharepoint.stackexchange.com/questions/133489/publishingweb-getpublishingweb-is-not-working-in-sharepoint-2013-csom
        /// http://sharepoint.stackexchange.com/questions/84914/how-to-check-if-publishing-is-on-using-csom/
        /// </summary>
       
        public bool IsPublishingWeb(string webUrl)
        {
            using (var clientContext = new ClientContext(webUrl))
            {
                Web web = clientContext.Web;

                var propName = "__PublishingFeatureActivated";
                //Ensure web properties are loaded
                if (!web.IsObjectPropertyInstantiated("AllProperties"))
                {
                    clientContext.Load(web, w => w.AllProperties);
                    clientContext.ExecuteQuery();
                }
                //Verify whether publishing feature is activated 
                if (web.AllProperties.FieldValues.ContainsKey(propName))
                {
                    bool propVal;
                    Boolean.TryParse((string)web.AllProperties[propName], out propVal);
                    return propVal;
                }
            }
            return false;
        }
        public bool IsPublishingWeb(ClientContext clientContext, Web web)
        {
            var _IsPublished = false;
            var propName = "__PublishingFeatureActivated";
            //Ensure web properties are loaded
            if (!web.IsObjectPropertyInstantiated("AllProperties"))
            {
                clientContext.Load(web, w => w.AllProperties);
                clientContext.ExecuteQuery();
            }
            //Verify whether publishing feature is activated 
            if (web.AllProperties.FieldValues.ContainsKey(propName))
            {
                bool propVal;
                Boolean.TryParse((string)web.AllProperties[propName], out propVal);
                _IsPublished = propVal;
                return propVal;
            }

            return _IsPublished;
        }
        
        /// <summary>
        /// Initialized of Exception and Logger Class. Deleted the Page/PageLayout Replace Usage File
        /// </summary>
        /// <param name="DiscoveryUsage_OutPutFolder"></param>
        private void PageAndPageLayout_Initialization(string DiscoveryUsage_OutPutFolder)
        {
            //Excception CSV Creation Command
            ExceptionCsv objException = ExceptionCsv.CurrentInstance;
            objException.CreateLogFile(DiscoveryUsage_OutPutFolder);

            ExceptionCsv.WebApplication = Constants.NotApplicable;
            ExceptionCsv.SiteCollection = Constants.NotApplicable;
            ExceptionCsv.WebUrl = Constants.NotApplicable;

            //Trace Log TXT File Creation Command
            Logger objTraceLogs = Logger.CurrentInstance;
            objTraceLogs.CreateLogFile(DiscoveryUsage_OutPutFolder);
            //Trace Log TXT File Creation Command

            //Delete Page and Page Layout Replace OUTPUT File
            Delete_PageAndPageLayout_ReplaceOutPutFiles(DiscoveryUsage_OutPutFolder);

        }
        /// <summary>
        /// This function delete all the existing files from <outPutFolder> folder
        /// </summary>
        /// <param name="outPutFolder"></param>
        private void Delete_PageAndPageLayout_ReplaceOutPutFiles(string outPutFolder)
        {
            FileUtility.DeleteFiles(outPutFolder + @"\" + Constants.PageLayoutUsage);
        }
        
        public bool Check_PageLayoutExistsINGallery(ClientContext clientContext, string PageLayoutURL)
        {
            //Checking The File in Root Web Gallery
            Microsoft.SharePoint.Client.File file = clientContext.Site.RootWeb.GetFileByServerRelativeUrl(PageLayoutURL);

            bool bExists = false;
            try
            {
                clientContext.Load(file);
                clientContext.ExecuteQuery(); //Raises exception if the file doesn't exist
                bExists = file.Exists;  
            }
            catch { }

            return bExists;
        }

        public void ChangePageLayoutForDiscoveryOutPut(string DiscoveryUsage_OutPutFolder, string PageLayoutUsageFilePath, string oldPageLayoutUrl = "N/A", string newPageLayoutUrl = "N/A", string newPageLayoutDescription = "N/A", string SharePointOnline_OR_OnPremise = "OP", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;
            try
            {
                PageAndPageLayout_Initialization(DiscoveryUsage_OutPutFolder);
                
                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## PageLayouts and Page Trasnformation Utility Execution Started : InputCSV ##############");
                Console.WriteLine("############## PageLayouts and Page Trasnformation Utility Execution Started : InputCSV ##############");
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ::: ChangePageLayoutForDiscoveryOutPut");
                Console.WriteLine("[START] ENTERING IN FUNCTION ::: ChangePageLayoutForDiscoveryOutPut");
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangePageLayoutForDiscoveryOutPut] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + DiscoveryUsage_OutPutFolder);
                Console.WriteLine("[ChangePageLayoutForDiscoveryOutPut] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + DiscoveryUsage_OutPutFolder);
               
                //Reading Page Layout Input File
                IEnumerable<PageLayoutInput> objPageInput;
                ReadPagesFromDiscoveryUsageFiles(DiscoveryUsage_OutPutFolder, PageLayoutUsageFilePath, out objPageInput);

                //Changing Pages/Layout for INPUT CSV
                ChangePageLayoutUsingPagesOutPut(DiscoveryUsage_OutPutFolder, objPageInput, oldPageLayoutUrl, newPageLayoutUrl, newPageLayoutDescription, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                Console.WriteLine("[END] EXIT FROM FUNCTION ::: ChangePageLayoutForDiscoveryOutPut");
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] EXIT FROM FUNCTION ::: ChangePageLayoutForDiscoveryOutPut");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## PageLayouts and Page Trasnformation Utility Execution Completed : InputCSV ##############");
                Console.WriteLine("############## PageLayouts and Page Trasnformation Utility Execution Completed : InputCSV ##############");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception] FUNCTION ChangePageLayoutForDiscoveryOutPut. Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] FUNCTION ChangePageLayoutForDiscoveryOutPut. Exception Message: " + ex.Message);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "PageLayout", ex.Message, ex.ToString(), "ChangePageLayoutForDiscoveryOutPut", ex.GetType().ToString());
            }
        }

        private void ReadPagesFromDiscoveryUsageFiles(string outPutFolder, string PageLayoutUsageFilePath, out IEnumerable<PageLayoutInput> objPLInput)
        {
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ::: ReadPagesFromDiscoveryUsageFiles");
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReadPagesFromDiscoveryUsageFiles] [START] Calling function ImportCsv.Read<PageLayoutInput>. Page Input CSV file is available at " + outPutFolder + " and Page Input file name is " + Constants.PageLayoutInput);

            objPLInput = null;
            //objPLInput = ImportCsv.Read<PageLayoutInput>(outPutFolder + @"\" + Transformation.PowerShell.Common.Constants.PageLayoutInput, Transformation.PowerShell.Common.Constants.CsvDelimeter);
            objPLInput = ImportCsv.ReadMatchingColumns<PageLayoutInput>(PageLayoutUsageFilePath, Transformation.PowerShell.Common.Constants.CsvDelimeter);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReadPagesFromDiscoveryUsageFiles] [END] Read all the INPUT from Page Input and saved in List - out IEnumerable<PageInput> objPLInput, for processing.");
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] EXIT FROM FUNCTION ::: ReadPagesFromDiscoveryUsageFiles");
        }
        
        private void ChangePageLayoutUsingPagesOutPut(string DiscoveryUsage_OutPutFolder, IEnumerable<PageLayoutInput> objPLInput, string oldPageLayoutUrl, string newPageLayoutUrl, string newPageLayoutDescription, string SharePointOnline_OR_OnPremise = "OP", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;
            
            try
            {
                if (objPLInput.Any())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ChangePageLayoutUsingPagesOutPut");

                    bool headerPageLayout = false;

                    foreach (PageLayoutInput objInput in objPLInput)
                    {
                        //This is for Exception Comments:
                        exceptionCommentsInfo1 = "PageLayoutUrl: " + objInput.PageLayout_ServerRelativeUrl + ", WebUrl: " + objInput.WebUrl + ", Page Layout Name: " + objInput.PageLayout_Name;
                        //This is for Exception Comments:
                        
                        //!= "all" && "" => This will update the page layout only in those pages in sites/webs, which have Page Layout URL == <Input>oldPageLayoutUrl
                        //if (oldPageLayoutUrl.ToLower().Trim() != Constants.Input_All && oldPageLayoutUrl.ToLower().Trim() != Constants.Input_Blank)
                        if (oldPageLayoutUrl.ToLower().Trim() != Constants.Input_All)
                        {
                            List<PageLayoutBase> objbase = ChangePageLayoutForPagesUsingOldPageLayoutInWeb(DiscoveryUsage_OutPutFolder, objInput.WebUrl, newPageLayoutUrl, oldPageLayoutUrl.ToLower().Trim(), newPageLayoutDescription, Constants.ActionType_CSV.ToLower(), SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                            if (objbase.Count > 0)
                            {
                                FileUtility.WriteCsVintoFile(DiscoveryUsage_OutPutFolder + @"\" + Constants.PageLayoutUsage, ref objbase, ref headerPageLayout);
                            }
                        }
                        //==all || =="" => This will update the page layout from all input web/site Using newPageLayoutUrl, in all the pages 
                        else
                        {
                            List<PageLayoutBase> objbase = ChangePageLayoutForPagesUsingOldPageLayoutInWeb(DiscoveryUsage_OutPutFolder, objInput.WebUrl, newPageLayoutUrl, objInput.PageLayout_ServerRelativeUrl, newPageLayoutDescription, Constants.ActionType_CSV.ToLower(), SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                            if (objbase.Count > 0)
                            {
                                FileUtility.WriteCsVintoFile(DiscoveryUsage_OutPutFolder + @"\" + Constants.PageLayoutUsage, ref objbase, ref headerPageLayout);
                            }
                        }
                    }

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] EXIT FROM FUNCTION ChangePageLayoutUsingPagesOutPut");
                    Console.WriteLine("[END] EXIT FROM FUNCTION ChangePageLayoutUsingPagesOutPut");
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "PageLayout", ex.Message, ex.ToString(), "ChangePageLayoutUsingPagesOutPut", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ChangePageLayoutUsingPagesOutPut] ChangePageLayoutUsingPagesOutPut. Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [ChangePageLayoutUsingPagesOutPut] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
       
        public void ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection(string outPutFolder, string SiteCollectionUrl = "N/A", string oldPageLayoutUrl = "N/A", string newPageLayoutUrl = "N/A", string newPageLayoutDescription = "N/A", string SharePointOnline_OR_OnPremise = "OP", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;
            
            PageAndPageLayout_Initialization(outPutFolder);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Pages/Page Layout Trasnformation Utility Execution Started - For Site Collection ##############");
            Console.WriteLine("############## Pages/Page Layout Trasnformation Utility Execution Started - For Site Collection ##############");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
            Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ::: ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection");
            Console.WriteLine("[START] ENTERING IN FUNCTION ::: ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] Initiated Logger and Exception Class. Logger and Exception file will be available at path " + outPutFolder);
            Console.WriteLine("[ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] Initiated Logger and Exception Class. Logger and Exception file will be available at path " + outPutFolder);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] SiteCollectionUrl is " + SiteCollectionUrl);
            Console.WriteLine("[ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] SiteCollectionUrl is " + SiteCollectionUrl);

            try
            {
                AuthenticationHelper ObjAuth = new AuthenticationHelper();
                ClientContext clientContext = null;
                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.OnPremise)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] ENTERING IN FUNCTION GetNetworkCredentialAuthenticatedContext for SiteCollectionUrl: " + SiteCollectionUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(SiteCollectionUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] EXIT FROM FUNCTION GetNetworkCredentialAuthenticatedContext for SiteCollectionUrl: " + SiteCollectionUrl);
                
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.Online)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] ENTERING IN FUNCTION GetSharePointOnlineAuthenticatedContextTenant for SiteCollectionUrl: " + SiteCollectionUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(SiteCollectionUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] EXIT FROM FUNCTION GetSharePointOnlineAuthenticatedContextTenant for SiteCollectionUrl: " + SiteCollectionUrl);
                }

                if (clientContext != null)
                {
                    List<PageLayoutBase> objPLBase = new List<PageLayoutBase>();
                    bool headerMasterPage = false;
                    ExceptionCsv.SiteCollection = SiteCollectionUrl;

                    Web rootWeb = clientContext.Web;
                    clientContext.Load(rootWeb);
                    clientContext.ExecuteQuery();

                    WebCollection webCollection = rootWeb.Webs;
                    clientContext.Load(webCollection);
                    clientContext.ExecuteQuery();

                    //Root Web
                    objPLBase = ChangePageLayoutForPagesUsingOldPageLayoutInWeb(outPutFolder, rootWeb.Url.ToString(), newPageLayoutUrl, oldPageLayoutUrl, newPageLayoutDescription, Constants.ActionType_SiteCollection, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
                    if (objPLBase != null)
                    {
                        FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.PageLayoutUsage, ref objPLBase,
                            ref headerMasterPage);
                    } 
                    
                    foreach (Web webSite in webCollection)
                    {
                        try
                        {
                            //Web
                            objPLBase = ChangePageLayoutForPagesUsingOldPageLayoutInWeb(outPutFolder, webSite.Url,newPageLayoutUrl, oldPageLayoutUrl, newPageLayoutDescription,Constants.ActionType_SiteCollection, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
                            
                            if (objPLBase != null)
                            {
                                FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.PageLayoutUsage, ref objPLBase,
                                    ref headerMasterPage);
                            }
                        }
                        catch (Exception ex)
                        {
                            ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "PageLayout", ex.Message, ex.ToString(), "ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection", ex.GetType().ToString(), exceptionCommentsInfo1);
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection. Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                            
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("[EXCEPTION] [ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] Exception Message: " + ex.Message);
                            Console.ForegroundColor = ConsoleColor.Gray;
                        }
                    }                    
                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] [ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] EXIT FROM FUNCTION ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection for SiteCollectionUrl: " + SiteCollectionUrl);
                Console.WriteLine("[END] [ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] EXIT FROM FUNCTION ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection for SiteCollectionUrl: " + SiteCollectionUrl);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Pages/Page Layout Trasnformation Utility Execution Completed - For Site Collection ##############");
                Console.WriteLine("############## Pages/Page Layout  Trasnformation Utility Execution Completed - For Site Collection ##############");
                
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "PageLayout", ex.Message, ex.ToString(), "ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        
        public List<PageLayoutBase> ChangePageLayoutForPagesUsingOldPageLayoutInWeb(string outPutFolder, string WebUrl, string newPageLayoutUrl, string oldPageLayoutUrl = "N/A", string newPageLayoutDescription = "N/A", string ActionType = "", string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            bool headerMasterPage = false;
            string exceptionCommentsInfo1 = string.Empty;
            List<PageLayoutBase> objPLBase = new List<PageLayoutBase>();

            ExceptionCsv.WebUrl = WebUrl;
            
            ///<ActionType==""> That means this function running only for a web. We have to write the output in this function only
            ///<Action Type=="SiteCollection"> The function will return object PageLayoutBase, and consolidated output will be written in SiteCollection function - ChangePageLayoutForPagesUsingOldPageLayoutInWeb

            //If ==> This is for WEB
            if (ActionType.ToString().ToLower() == Constants.ActionType_Web.ToLower())
            {
                PageAndPageLayout_Initialization(outPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Pages/Page Layout Trasnformation Utility Execution Started - For Web ##############");
                Console.WriteLine("############## Pages/Page Layout  Trasnformation Utility Execution Started - For Web ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ::: ChangePageLayoutForPagesUsingOldPageLayoutInWeb");
                Console.WriteLine("[START] ENTERING IN FUNCTION ::: ChangePageLayoutForPagesUsingOldPageLayoutInWeb");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Initiated Logger and Exception Class. Logger and Exception file will be available in path " + outPutFolder);
                Console.WriteLine("[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] WebUrl is " + WebUrl);
                Console.WriteLine("[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] WebUrl is " + WebUrl);

            }

            try
            {
                AuthenticationHelper ObjAuth = new AuthenticationHelper();
                ClientContext clientContext = null;

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.OnPremise)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ChangePageLayoutForPagesUsingOldPageLayoutInWeb] ENTERING IN FUNCTION GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(WebUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ChangePageLayoutForPagesUsingOldPageLayoutInWeb] EXIT FROM FUNCTION GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.Online)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ChangePageLayoutForPagesUsingOldPageLayoutInWeb] ENTERING IN FUNCTION GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(WebUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ChangePageLayoutForPagesUsingOldPageLayoutInWeb] EXIT FROM FUNCTION GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                }

                if (clientContext != null)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ChangePageLayoutForPagesUsingOldPageLayoutInWeb] ENTERING IN FUNCTION ChangePageLayoutForPagesUsingOldPageLayoutInWeb for WebUrl: " + WebUrl + ", after successful authentication");
                    Console.WriteLine("[START][ChangePageLayoutForPagesUsingOldPageLayoutInWeb] ENTERING IN FUNCTION ChangePageLayoutForPagesUsingOldPageLayoutInWeb for WebUrl: " + WebUrl + ", after successful authentication");
                    Web web = clientContext.Web;

                    //Load Web to get old PageLayout details
                    clientContext.Load(web);
                    clientContext.ExecuteQuery();
                    //Load Web to get old PageLayout details

                    //Create New Page Layout Relative URL
                    string _newPageLayoutUrl = string.Empty;
                    _newPageLayoutUrl = GetPageLayoutRelativeURL(clientContext, newPageLayoutUrl);

                    //Create oldPageLayoutUrl Relative URL
                    string _strOLDPageLayoutUrl = string.Empty;
                    if (oldPageLayoutUrl.Trim() != Constants.Input_Blank && oldPageLayoutUrl.Trim().ToLower() != Constants.Input_All)
                    {
                        _strOLDPageLayoutUrl = GetPageLayoutRelativeURL(clientContext, oldPageLayoutUrl);
                    }
                    
                    bool _IsPublishingWeb = IsPublishingWeb(clientContext, web);

                    exceptionCommentsInfo1 = "newPageLayoutUrl: " + newPageLayoutUrl + ", oldPageLayoutUrl: " + oldPageLayoutUrl + ", newPageLayoutDescription: " + newPageLayoutDescription;
                    
                    if (_IsPublishingWeb)
                    {
                        //The publishing feature is activated on this web
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Web: " + WebUrl + " is a publishing site");
                        Console.WriteLine("[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Web : " + WebUrl + " is a publishing site");

                         //Check if new page alyout is available in Gallery
                        if (Check_PageLayoutExistsINGallery(clientContext, _newPageLayoutUrl))
                        {
                            //Get the List of all pages availble in this web
                            List pagesList = web.Lists.GetByTitle("Pages");
                            var allItemsQuery = CamlQuery.CreateAllItemsQuery();

                            ListItemCollection items = pagesList.GetItems(allItemsQuery);
                            clientContext.Load(items);
                            clientContext.ExecuteQuery();

                            foreach (ListItem item in items)
                            {
                                var pageLayout = item["PublishingPageLayout"] as FieldUrlValue;
                                // This will return the Full URL of Page Layout - https://.....aspx

                                try
                                {
                                    if (oldPageLayoutUrl.Trim().ToLower() != Constants.Input_Blank && oldPageLayoutUrl.Trim().ToLower() != Constants.Input_All)
                                    {
                                        if (pageLayout.Url.ToString().Trim().ToLower().EndsWith(_strOLDPageLayoutUrl.Trim().ToLower()))
                                        {
                                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Web : " + WebUrl + " contain page layout with Url : " + _strOLDPageLayoutUrl);
                                            Console.WriteLine("[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Web : " + WebUrl + " contain page layout with Url : " + _strOLDPageLayoutUrl);

                                            item.File.CheckOut();
                                            item["PublishingPageLayout"] = new FieldUrlValue() { Url = _newPageLayoutUrl, Description = newPageLayoutDescription };
                                            item.Update();

                                            item.File.CheckIn("comment", CheckinType.MajorCheckIn);
                                            item.File.Publish("comments");
                                            clientContext.ExecuteQuery();

                                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangePageLayoutForPagesUsingOldPageLayoutInWeb][OLD PageLayout is !=<Blank>]Changed Page Layout from  " + _strOLDPageLayoutUrl + " to  " + _newPageLayoutUrl + "for Page ID: " + item["ID"].ToString() + ", Title: " + item["Title"].ToString());
                                            Console.WriteLine("[ChangePageLayoutForPagesUsingOldPageLayoutInWeb][OLD PageLayout is !=<Blank>]Changed Page Layout from  " + _strOLDPageLayoutUrl + " to  " + _newPageLayoutUrl + "for Page ID: " + item["ID"].ToString() + ", Title: " + item["Title"].ToString());
                                           
                                            PageLayoutBase objPLOut = new PageLayoutBase();
                                            objPLOut.PageId = item["ID"].ToString();
                                            objPLOut.PageTitle = item["Title"].ToString();
                                            objPLOut.OldPageLayoutUrl = pageLayout.Url;
                                            objPLOut.OldPageLayoutDescription = pageLayout.Description;
                                            objPLOut.NewPageLayoutUrl = _newPageLayoutUrl;
                                            objPLOut.NewPageLayoutDescription = newPageLayoutDescription;
                                            objPLOut.WebUrl = WebUrl;
                                            objPLOut.SiteCollection = Constants.NotApplicable;
                                            objPLOut.WebApplication = Constants.NotApplicable;

                                            objPLBase.Add(objPLOut);
                                        }
                                        else
                                        {
                                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangePageLayoutForPagesUsingOldPageLayoutInWeb][OLD PageLayout !=\"\"]: [NO Update in PageLayout] <INPUT> OLD PageLayout " + _strOLDPageLayoutUrl.Trim().ToLower() + ", <WEB> OLD PageLayout URL " + pageLayout.Url.ToString().Trim().ToLower());
                                            Console.WriteLine("[ChangePageLayoutForPagesUsingOldPageLayoutInWeb]:[OLD PageLayout !=\"\"]: [NO Update in PageLayout] <INPUT> OLD PageLayout  " + _strOLDPageLayoutUrl.Trim().ToLower() + ", <WEB> OLD PageLayout URL " + pageLayout.Url.ToString().Trim().ToLower());
                                        }
                                    }
                                    else
                                    {
                                        item.File.CheckOut();

                                        item["PublishingPageLayout"] = new FieldUrlValue() { Url = _newPageLayoutUrl, Description = newPageLayoutDescription };
                                        item.Update();

                                        item.File.CheckIn("comment", CheckinType.MajorCheckIn);
                                        item.File.Publish("comments");
                                        clientContext.ExecuteQuery();

                                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] [OLD PageLayout is ==<Blank>]Changed Page Layout from  " + pageLayout.Url + " to  " + _newPageLayoutUrl + "for Page ID: " + item["ID"].ToString() + ", Title: " + item["Title"].ToString());
                                        Console.WriteLine("[ChangePageLayoutForPagesUsingOldPageLayoutInWeb][OLD PageLayout is ==<Blank>]Changed Page Layout from  " + pageLayout.Url + " to  " + _newPageLayoutUrl + "for Page ID: " + item["ID"].ToString() + ", Title: " + item["Title"].ToString());
                                       
                                        PageLayoutBase objPLOut = new PageLayoutBase();
                                        objPLOut.PageId = item["ID"].ToString();
                                        objPLOut.PageTitle = item["Title"].ToString();
                                        objPLOut.OldPageLayoutUrl = pageLayout.Url;
                                        objPLOut.OldPageLayoutDescription = pageLayout.Description;
                                        objPLOut.NewPageLayoutUrl = _newPageLayoutUrl;
                                        objPLOut.NewPageLayoutDescription = newPageLayoutDescription;
                                        objPLOut.WebUrl = WebUrl;
                                        objPLOut.SiteCollection = Constants.NotApplicable;
                                        objPLOut.WebApplication = Constants.NotApplicable;

                                        objPLBase.Add(objPLOut);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "PageLayout", ex.Message, ex.ToString(), "ChangePageLayoutForPagesUsingOldPageLayoutInWeb", ex.GetType().ToString(), exceptionCommentsInfo1);
                                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                                    
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("[EXCEPTION] [ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Exception Message: " + ex.Message);
                                    Console.ForegroundColor = ConsoleColor.Gray;
                                }
                            }
                        }
                        else
                        {
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] We have not changed the any page because this new  Page Layout " + _newPageLayoutUrl + " is not present in Gallary, for Web " + WebUrl);
                            Console.WriteLine("[ChangePageLayoutForPagesUsingOldPageLayoutInWeb]  We have not changed the any page because this new  Page Layout " + _newPageLayoutUrl + " is not present in Gallary, for Web " + WebUrl);
                        }
                    }
                    else
                    {
                        Console.WriteLine("[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Web : " + WebUrl + " does not contain Pages List or it's not a publishing web");
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Web : " + WebUrl + " does not contain Pages List or it's not a publishing web");
                    }
                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] EXIT FROM FUNCTION ChangePageLayoutForPagesUsingOldPageLayoutInWeb for WebUrl: " + WebUrl);
                Console.WriteLine("[END] EXIT FROM FUNCTION ChangePageLayoutForPagesUsingOldPageLayoutInWeb for WebUrl: " + WebUrl);

            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "PageLayout", ex.Message, ex.ToString(), "ChangePageLayoutForPagesUsingOldPageLayoutInWeb", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            //If ==> This is for WEB
            if (ActionType.ToString().ToLower() == Constants.ActionType_Web.ToLower())
            {
                if (objPLBase != null)
                {
                    FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.PageLayoutUsage, ref objPLBase,
                        ref headerMasterPage);

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Writing the Replace Output CSV file after replacing the page layout - FileUtility.WriteCsVintoFile");
                    Console.WriteLine("[ChangePageLayoutForPagesUsingOldPageLayoutInWeb] Writing the Replace Output CSV file after replacing the page layout - FileUtility.WriteCsVintoFile");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ChangePageLayoutForPagesUsingOldPageLayoutInWeb] EXIT FROM FUNCTION ChangePageLayoutForPagesUsingOldPageLayoutInWeb for WebUrl: " + WebUrl);
                    Console.WriteLine("[END][ChangePageLayoutForPagesUsingOldPageLayoutInWeb] EXIT FROM FUNCTION ChangePageLayoutForPagesUsingOldPageLayoutInWeb for WebUrl: " + WebUrl);

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Pages/Page Layout Trasnformation Utility Execution Completed for Web ##############");
                    Console.WriteLine("############## Pages/Page Layout Trasnformation Utility Execution Completed  for Web ##############");
                }
            }

            return objPLBase;
        }
       
        public string GetPageLayoutRelativeURL(ClientContext clientContext, string PageLayoutUrl)
        {
            string _relativePageLayoutUrl = string.Empty;

            //User has Input Page Layout URL from Root Gallery
            if (PageLayoutUrl.ToLower().StartsWith("/_catalogs/masterpage/"))
            {
                Web rootWeb = clientContext.Site.RootWeb;
                clientContext.Load(rootWeb);
                clientContext.ExecuteQuery();

                _relativePageLayoutUrl = rootWeb.ServerRelativeUrl.ToString() + PageLayoutUrl;
            }
            else if (PageLayoutUrl.ToLower().StartsWith("/sites/"))
            {
                _relativePageLayoutUrl = PageLayoutUrl;
            }
            else if (PageLayoutUrl.ToLower().Contains("/_catalogs/masterpage/"))
            {
                _relativePageLayoutUrl = PageLayoutUrl;
            }
            else
            {
                Web rootWeb = clientContext.Site.RootWeb;

                clientContext.Load(rootWeb);
                clientContext.ExecuteQuery();

                _relativePageLayoutUrl = rootWeb.ServerRelativeUrl.ToString() + "/_catalogs/masterpage/" + PageLayoutUrl;
            }

            return _relativePageLayoutUrl;
        }

        #region Not Using these functions
        /// <summary>
        /// 
        /// </summary>
        /// <param name="webUrl"></param>
        /// <param name="serverRelativeUrl"></param>
        /// <returns></returns>
        private string GetPageLayoutUrl(string webUrl, string serverRelativeUrl)
        {
            string pageLayoutUrl = null;
            try
            {
                string[] webUrlSplit = webUrl.Split('/');

                pageLayoutUrl = webUrlSplit[0] + "/" + webUrlSplit[1] + "/" + webUrlSplit[2] + serverRelativeUrl;
            }
            catch
            { }
            return pageLayoutUrl;
        }
        private void ChangePageLayoutForPagesInWeb(string webUrl, string pageTitle, string pagaLayoutUrl, string description = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;
            try
            {
                using (var clientContext = new ClientContext(webUrl))
                {
                    ExceptionCsv.WebUrl = webUrl;

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ChangePageLayoutForPagesInWeb for WebUrl: " + webUrl);

                    Web web = clientContext.Web;
                    ListCollection listCollection = web.Lists;
                    clientContext.Load(listCollection);
                    clientContext.ExecuteQuery();
                    ///
                    /// Following code is to check whether the web is as publishing or not. 
                    ///

                    /*var pWeb = PublishingWeb.GetPublishingWeb(clientContext, web);
                    clientContext.Load(pWeb);
                    if (pWeb != null)
                    {
                        Console.WriteLine("Web : " + webUrl + " is a publishing web");
                    }
                    else
                    {
                        Console.WriteLine("Web : " + webUrl + "is not a publishing web");
                    }*/

                    bool pagesListAvailability = false;
                    foreach (List oList in listCollection)
                    {
                        if (oList.Title.Equals("Pages"))
                        {
                            pagesListAvailability = true;
                        }
                    }
                    exceptionCommentsInfo1 = "NewPageLayoutUrl: " + pagaLayoutUrl + " PageTitle: " + pageTitle + ", PageLayoutDescription: " + description;

                    if (pagesListAvailability)
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "Web : " + webUrl + " contain Pages List");

                        List pagesList = web.Lists.GetByTitle("Pages");
                        var allItemsQuery = CamlQuery.CreateAllItemsQuery();
                        ListItemCollection items = pagesList.GetItems(allItemsQuery);
                        clientContext.Load(items);
                        clientContext.ExecuteQuery();
                        foreach (ListItem item in items)
                        {
                            var title = item["Title"];
                            try
                            {
                                if (title.Equals(pageTitle))
                                {
                                    Logger.AddMessageToTraceLogFile(Constants.Logging, "Web : " + webUrl + " contain page with Title : " + pageTitle);
                                    item.File.CheckOut();
                                    item["PublishingPageLayout"] = new FieldUrlValue() { Url = pagaLayoutUrl, Description = description };
                                    item.Update();
                                    item.File.CheckIn("comment", CheckinType.MajorCheckIn);
                                    item.File.Publish("comments");
                                    clientContext.ExecuteQuery();
                                    Logger.AddMessageToTraceLogFile(Constants.Logging, "Changed Page Layout for page with Title - " + pageTitle + ", New page layout  is " + pagaLayoutUrl);
                                }
                            }
                            catch (Exception ex)
                            {
                                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "PageLayout", ex.Message, ex.ToString(), "ChangePageLayoutForPagesInWeb", ex.GetType().ToString(), exceptionCommentsInfo1);
                            }
                        }
                    }
                    else
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "Web : " + webUrl + " does not contain Pages List");
                    }
                }
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] EXIT FROM FUNCTION ChangePageLayoutForPagesInWeb for WebUrl: " + webUrl);
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "PageLayout", ex.Message, ex.ToString(), "ChangePageLayoutForPagesInWeb", ex.GetType().ToString(), exceptionCommentsInfo1);
            }

        }

        private void ChangePageLayoutForPagesInWebUsingId(string webUrl, string id, string pagaLayoutUrl, string description = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;
            try
            {
                using (var clientContext = new ClientContext(webUrl))
                {
                    ExceptionCsv.WebUrl = webUrl;
                    //Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ChangePageLayoutForPagesInWebUsingId for WebUrl: " + webUrl);

                    Web web = clientContext.Web;
                    ListCollection listCollection = web.Lists;
                    clientContext.Load(listCollection);
                    clientContext.ExecuteQuery();
                    ///
                    /// Following code is to check whether the web is as publishing or not. 
                    ///

                    /*var pWeb = PublishingWeb.GetPublishingWeb(clientContext, web);
                    clientContext.Load(pWeb);
                    if (pWeb != null)
                    {
                        Console.WriteLine("Web : " + webUrl + " is a publishing web");
                    }
                    else
                    {
                        Console.WriteLine("Web : " + webUrl + "is not a publishing web");
                    }*/

                    bool pagesListAvailability = false;
                    foreach (List oList in listCollection)
                    {
                        if (oList.Title.Equals("Pages"))
                        {
                            pagesListAvailability = true;
                        }
                    }
                    exceptionCommentsInfo1 = "NewPageLayoutUrl: " + pagaLayoutUrl + " PageId: " + id + ", PageLayoutDescription: " + description;

                    if (pagesListAvailability)
                    {
                        //Logger.AddMessageToTraceLogFile(Constants.Logging, "Web : " + webUrl + " contain Pages List"); 

                        List pagesList = web.Lists.GetByTitle("Pages");
                        var allItemsQuery = CamlQuery.CreateAllItemsQuery();
                        ListItemCollection items = pagesList.GetItems(allItemsQuery);
                        clientContext.Load(items);
                        clientContext.ExecuteQuery();
                        foreach (ListItem item in items)
                        {
                            var itemId = item["ID"];
                            try
                            {
                                if (itemId.ToString().Equals(id))
                                {
                                    //Logger.AddMessageToTraceLogFile(Constants.Logging, "Web : " + webUrl + " contain page with Id : " + id); 
                                    item.File.CheckOut();
                                    item["PublishingPageLayout"] = new FieldUrlValue() { Url = pagaLayoutUrl, Description = description };
                                    item.Update();
                                    item.File.CheckIn("comment", CheckinType.MajorCheckIn);
                                    item.File.Publish("comments");
                                    clientContext.ExecuteQuery();
                                    //Logger.AddMessageToTraceLogFile(Constants.Logging, "Changed Page Layout for page with Id - " + id + ", New page layout  is " + pagaLayoutUrl);

                                }
                            }
                            catch (Exception ex)
                            {
                                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "PageLayout", ex.Message, ex.ToString(), "ChangePageLayoutForPagesInWebUsingId", ex.GetType().ToString(), exceptionCommentsInfo1);
                            }
                        }
                    }
                    else
                    {
                        //Logger.AddMessageToTraceLogFile(Constants.Logging, "Web : " + webUrl + " does not contain Pages List"); 
                        Console.WriteLine("Web : " + webUrl + " does not contain Pages List");
                    }
                }
                //Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] EXIT FROM FUNCTION ChangePageLayoutForPagesInWebUsingId for WebUrl: " + webUrl);
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "PageLayout", ex.Message, ex.ToString(), "ChangePageLayoutForPagesInWebUsingId", ex.GetType().ToString(), exceptionCommentsInfo1);
            }
        }
        #endregion

    }   
}
