using JDP.Remediation.Console.Common.Base;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Utilities;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDP.Remediation.Console
{
    public class DeleteWebparts
    {
        public static string filePath = string.Empty;
        public static string outputPath = string.Empty;
        public static void DoWork()
        {
            outputPath = Environment.CurrentDirectory;
            string webPartsInputFile = string.Empty;
            string webpartType = string.Empty;
            IEnumerable<WebpartInput> objWPDInput;
            string timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");
            //Trace Log TXT File Creation Command
            Logger.OpenLog("DeleteWebparts", timeStamp);

            if (!ReadInputFile(ref webPartsInputFile))
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("Webparts input file is not valid or available. So, Operation aborted!");
                Logger.LogErrorMessage("Please enter path like: E.g. C:\\<Working Directory>\\<InputFile>.csv");
                System.Console.ResetColor();
                return;
            }

            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            Logger.LogMessage("Please enter Webpart Type (enter 'all' to delete all webparts):");
            System.Console.ResetColor();
            webpartType = System.Console.ReadLine().ToLower();

            try
            {
                string csvFile = outputPath + @"/" + Constants.DeleteWebpartStatus + timeStamp + Constants.CSVExtension;
                if (System.IO.File.Exists(csvFile))
                    System.IO.File.Delete(csvFile);

                if (System.IO.File.Exists(webPartsInputFile))
                {
                    if (String.Equals(Constants.WebpartType_All, webpartType, StringComparison.CurrentCultureIgnoreCase))
                    {
                        //Reading Input File
                        objWPDInput = ImportCSV.ReadMatchingColumns<WebpartInput>(webPartsInputFile, Constants.CsvDelimeter);

                        if (objWPDInput.Any())
                        {
                            IEnumerable<string> webPartTypes = objWPDInput.Select(x => x.WebPartType);

                            webPartTypes = webPartTypes.Distinct();
                            Logger.LogInfoMessage(String.Format("Preparing to delete a total of {0} webparts ...", webPartTypes.Count()), true);

                            foreach (string webPartType in webPartTypes)
                            {
                                try
                                {
                                    DeleteWebPart_UsingCSV(webPartType, webPartsInputFile, csvFile);
                                }
                                catch (Exception ex)
                                {
                                    Logger.LogErrorMessage("[DeleteMissingWebparts: DoWork]. Exception Message: " + ex.Message, true);
                                    ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Webpart", ex.Message,
                                        ex.ToString(), "DoWork", ex.GetType().ToString(), Constants.NotApplicable);
                                }
                            }
                            webPartTypes = null;
                        }
                        else
                            Logger.LogInfoMessage("There is nothing to delete from the '" + webPartsInputFile + "' File ", true);
                    }
                    else
                    {
                        DeleteWebPart_UsingCSV(webpartType, webPartsInputFile, csvFile);
                    }
                    Logger.LogInfoMessage("Processing input file has been comepleted...");
                }
                else
                {
                    Logger.LogErrorMessage("[DeleteMissingWebparts: DoWork]The input file " + webPartsInputFile + " is not present", true);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DeleteMissingWebparts: DoWork]. Exception Message: " + ex.Message, true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Webpart", ex.Message,
                    ex.ToString(), "DoWork", ex.GetType().ToString(), Constants.NotApplicable);
            }
            finally
            {
                objWPDInput = null;
            }

            Logger.CloseLog();
        }

        public static void DeleteWebPart_UsingCSV(string webPartType, string webPartsInputFile, string csvFile)
        {
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                //Reading Input File
                IEnumerable<WebpartInput> objWPDInput;
                ReadWebPartUsageCSV(webPartType, webPartsInputFile, out objWPDInput);

                bool headerTransformWebPart = false;

                if (objWPDInput.Any())
                {
                    Logger.LogInfoMessage(String.Format("Preparing to delete a total of {0} webparts with webpart type {1} ...", objWPDInput.Count(), webPartType), true);
                    for (int i = 0; i < objWPDInput.Count(); i++)
                    {
                        WebpartInput objInput = objWPDInput.ElementAt(i);
                        WebpartDeleteOutputBase objWPOutputBase = new WebpartDeleteOutputBase();

                        try
                        {
                            bool status = DeleteWebPart(objInput.WebUrl, objInput.PageUrl.ToString(), objInput.StorageKey);

                            if (status)
                            {
                                objWPOutputBase.Status = Constants.Success;
                                System.Console.ForegroundColor = System.ConsoleColor.Green;
                                Logger.LogSuccessMessage("Successfully Deleted WebPart with Webpart Type " + objInput.WebPartType + " and with StorageKey " + objInput.StorageKey + "and output file is present in the path: " + Environment.CurrentDirectory, true);
                                System.Console.ResetColor();
                            }
                            else
                            {
                                objWPOutputBase.Status = Constants.Failure;
                                System.Console.ForegroundColor = System.ConsoleColor.Gray;
                                Logger.LogErrorMessage("Failed to Delete WebPart with Webpart Type " + objInput.WebPartType + " and with StorageKey " + objInput.StorageKey, true);
                                System.Console.ResetColor();
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.LogInfoMessage("Failed to Deleted WebPart with Webpart Type " + objInput.WebPartType + " and with StorageKey " + objInput.StorageKey);
                            Logger.LogErrorMessage("[DeleteMissingWebparts: DeleteWebPart_UsingCSV]. Exception Message: " + ex.Message, true);
                            ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Webpart", ex.Message,
                                ex.ToString(), "DeleteWebPart_UsingCSV", ex.GetType().ToString(), "Failed to Deleted WebPart with Webpart Type");
                        }

                        objWPOutputBase.WebPartType = objInput.WebPartType;
                        objWPOutputBase.PageUrl = objInput.PageUrl;
                        objWPOutputBase.WebUrl = objInput.WebUrl;
                        objWPOutputBase.StorageKey = objInput.StorageKey;
                        objWPOutputBase.ExecutionDateTime = DateTime.Now.ToString();

                        if (System.IO.File.Exists(csvFile))
                        {
                            headerTransformWebPart = true;
                        }
                        FileUtility.WriteCsVintoFile(csvFile, objWPOutputBase, ref headerTransformWebPart);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DeleteMissingWebparts: DeleteWebPart_UsingCSV]. Exception Message: " + ex.Message, true);
            }
        }

        private static void ReadWebPartUsageCSV(string sourceWebPartType, string usageFilePath, out IEnumerable<WebpartInput> objWPDInput)
        {
            objWPDInput = null;
            objWPDInput = ImportCSV.ReadMatchingColumns<WebpartInput>(usageFilePath, Constants.CsvDelimeter);

            try
            {
                if (objWPDInput.Any())
                {
                    objWPDInput = from p in objWPDInput
                                  where p.WebPartType.Equals(sourceWebPartType, StringComparison.OrdinalIgnoreCase)
                                  select p;

                    if (objWPDInput.Any())
                        Logger.LogInfoMessage("Number of Webparts found with WebpartType '" + sourceWebPartType + "' are " + objWPDInput.Count());
                    else
                        Logger.LogInfoMessage("No Webparts found with WebpartType '" + sourceWebPartType + "'");
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DeleteMissingWebparts: ReadWebPartUsageCSV]. Exception Message: " + ex.Message
                    + ", Exception Comments: Exception occured while rading input file ", true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Webpart", ex.Message,
                                ex.ToString(), "ReadWebPartUsageCSV", ex.GetType().ToString(), "Exception occured while rading input file");
            }
        }

        public static bool DeleteWebPart(string webUrl, string pageUrl, string _storageKey)
        {
            bool isWebPartDeleted = false;
            string webPartXml = string.Empty;
            string exceptionCommentsInfo1 = string.Empty;
            Web web = null;
            List list = null;

            try
            {
                //This function is Get Relative URL of the page
                string ServerRelativePageUrl = string.Empty;
                ServerRelativePageUrl = GetPageRelativeURL(webUrl, pageUrl);

                Guid storageKey = new Guid(GetWebPartID(_storageKey));

                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    web = userContext.Web;
                    userContext.Load(web);
                    userContext.ExecuteQuery();

                    Logger.LogInfoMessage("Successful authentication", false);

                    Logger.LogInfoMessage("Checking Out File ...", false);

                    list = GetPageList(userContext);

                    //Boolean to check if a call to Update method is required
                    bool needsUpdate = false;
                    bool forceCheckOut = false;
                    bool enableVersioning = false;
                    bool enableMinorVersions = false;
                    bool enableModeration = false;
                    DraftVisibilityType dVisibility = DraftVisibilityType.Author;

                    if (list != null)
                    {
                        try
                        {
                            userContext.Load(list, l => l.ForceCheckout,
                                       l => l.EnableVersioning,
                                       l => l.EnableMinorVersions,
                                       l => l.EnableModeration,
                                       l => l.Title,
                                       l => l.DraftVersionVisibility,
                                       l => l.DefaultViewUrl);

                            userContext.ExecuteQueryRetry();

                            #region Remove Versioning in List
                            forceCheckOut = list.ForceCheckout;
                            enableVersioning = list.EnableVersioning;
                            enableMinorVersions = list.EnableMinorVersions;
                            enableModeration = list.EnableModeration;
                            dVisibility = list.DraftVersionVisibility;

                            Logger.LogInfoMessage("Removing Versioning", false);
                            //Boolean to check if a call to Update method is required
                            needsUpdate = false;

                            if (enableVersioning)
                            {
                                list.EnableVersioning = false;
                                needsUpdate = true;
                            }
                            if (forceCheckOut)
                            {
                                list.ForceCheckout = false;
                                needsUpdate = true;
                            }
                            if (enableModeration)
                            {
                                list.EnableModeration = false;
                                needsUpdate = true;
                            }

                            if (needsUpdate)
                            {
                                list.Update();
                                userContext.ExecuteQuery();
                            }
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            Logger.LogErrorMessage("[DeleteMissingWebparts: DeleteWebPart]. Exception Message: " + ex.Message
                                + ", Exception Comments: Exception while removing Version to the list", true);
                            ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "Webpart", ex.Message,
                                ex.ToString(), "DeleteWebPart", ex.GetType().ToString(), "Exception while removing Version to the list");
                        }
                    }

                    try
                    {
                        if (DeleteWebPart(userContext.Web, ServerRelativePageUrl, storageKey))
                        {
                            isWebPartDeleted = true;
                            Logger.LogSuccessMessage("Successfully Deleted the WebPart", true);
                        }
                        else
                        {
                            Logger.LogErrorMessage("WebPart with StorageKey: " + storageKey + " does not exist in the Page: " + ServerRelativePageUrl, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogErrorMessage("[DeleteMissingWebparts: DeleteWebPart]. Exception Message: " + ex.Message, true);
                        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "Webpart", ex.Message,
                            ex.ToString(), "DeleteWebPart", ex.GetType().ToString(), Constants.NotApplicable);
                    }
                    finally
                    {
                        if (list != null)
                        {
                            #region Enable Versioning in List
                            //Reset the boolean so that it can used to test if we need to call Update method
                            needsUpdate = false;
                            if (enableVersioning)
                            {
                                list.EnableVersioning = true;
                                if (enableMinorVersions)
                                {
                                    list.EnableMinorVersions = true;
                                }
                                if (enableMinorVersions)
                                {
                                    list.EnableMinorVersions = true;
                                }

                                list.DraftVersionVisibility = dVisibility;
                                needsUpdate = true;
                            }
                            if (enableModeration)
                            {
                                list.EnableModeration = enableModeration;
                                needsUpdate = true;
                            }
                            if (forceCheckOut)
                            {
                                list.ForceCheckout = true;
                                needsUpdate = true;
                            }
                            if (needsUpdate)
                            {
                                list.Update();
                                userContext.ExecuteQuery();
                            }
                            #endregion
                        }
                        web = null;
                        list = null;
                    }
                    Logger.LogInfoMessage("File Checked in after successfully deleting the webpart.", false);
                }

            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DeleteMissingWebparts: DeleteWebPart]. Exception Message: " + ex.Message, true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "Webpart", ex.Message,
                    ex.ToString(), "DeleteWebPart", ex.GetType().ToString(), Constants.NotApplicable);
            }
            return isWebPartDeleted;
        }

        private static bool DeleteWebPart(Web web, string serverRelativePageUrl, Guid storageKey)
        {
            bool isWebPartDeleted = false;
            LimitedWebPartManager limitedWebPartManager = null;
            try
            {
                var webPartPage = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

                web.Context.Load(webPartPage);
                web.Context.ExecuteQueryRetry();

                limitedWebPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
                web.Context.Load(limitedWebPartManager.WebParts, wps => wps.Include(wp => wp.Id));
                web.Context.ExecuteQueryRetry();

                if (limitedWebPartManager.WebParts.Count >= 0)
                {
                    foreach (WebPartDefinition webpartDef in limitedWebPartManager.WebParts)
                    {
                        Microsoft.SharePoint.Client.WebParts.WebPart oWebPart = null;
                        try
                        {
                            oWebPart = webpartDef.WebPart;
                            if (webpartDef.Id.Equals(storageKey))
                            {
                                webpartDef.DeleteWebPart();
                                web.Context.ExecuteQueryRetry();
                                isWebPartDeleted = true;
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.LogErrorMessage("[DeleteMissingWebparts: DeleteWebPart]. Exception Message: " + ex.Message
                                + ", Exception Comments: Exception occured while deleting the webpart", true);
                            ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "Webpart", ex.Message,
                                ex.ToString(), "DeleteWebPart", ex.GetType().ToString(), Constants.NotApplicable);
                        }
                        finally
                        {
                            oWebPart = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DeleteMissingWebparts: DeleteWebPart]. Exception Message: "
                    + ex.Message + ", Exception Comments: Exception occure while fetching webparts using LimitedWebPartManager", true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "Webpart", ex.Message,
                    ex.ToString(), "DeleteWebPart", ex.GetType().ToString(), Constants.NotApplicable);
            }
            finally
            {
                limitedWebPartManager = null;
            }
            return isWebPartDeleted;
        }

        private static string GetPageRelativeURL(string WebUrl, string PageUrl, string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string _relativePageUrl = string.Empty;
            try
            {
                if (WebUrl != "" || PageUrl != "")
                {
                    using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, WebUrl))
                    {
                        Web _Web = userContext.Web;
                        userContext.Load(_Web);
                        userContext.ExecuteQuery();

                        Logger.LogInfoMessage("Web.ServerRelativeUrl: " + _Web.ServerRelativeUrl + " And PageUrl: " + PageUrl, true);

                        //Issue: Found in MARS Retraction Process, the root web ServerRelativeUrl would result "/" only
                        //Hence appending "/" would throw exception for ServerRelativeUrl parameter
                        if (_Web.ServerRelativeUrl.ToString().Equals("/"))
                        {
                            _relativePageUrl = _Web.ServerRelativeUrl.ToString() + PageUrl;
                        }
                        else if (!PageUrl.Contains(_Web.ServerRelativeUrl))
                        {
                            _relativePageUrl = _Web.ServerRelativeUrl.ToString() + "/" + PageUrl;
                        }
                        else
                        {
                            _relativePageUrl = PageUrl;
                        }
                        Logger.LogInfoMessage("RelativePageUrl Framed: " + _relativePageUrl, false);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DeleteMissingWebparts: GetPageRelativeURL]. Exception Message: " + ex.Message
                    + ", Exception Comments: Exception occured while reading page relive url", true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, WebUrl, "Webpart", ex.Message,
                    ex.ToString(), "GetPageRelativeURL", ex.GetType().ToString(), Constants.NotApplicable);
            }

            return _relativePageUrl;
        }

        private static string GetWebPartID(string webPartID)
        {
            string _webPartID = string.Empty;

            try
            {
                string[] tempStr = webPartID.Split('_');

                if (tempStr.Length > 5)
                {
                    _webPartID = webPartID.Remove(0, tempStr[0].Length + 1).Replace('_', '-');
                }
                else
                {
                    _webPartID = webPartID.Replace('_', '-');
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DeleteMissingWebparts: GetWebPartID]. Exception Message: " + ex.Message, true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Webpart", ex.Message,
                    ex.ToString(), "GetWebPartID", ex.GetType().ToString(), Constants.NotApplicable);
            }
            return _webPartID;
        }

        private static List GetPageList(ClientContext clientContext)
        {
            List list = null;
            Web web = null;
            try
            {
                web = clientContext.Web;

                // Get a few properties from the web
                clientContext.Load(web,
                                    w => w.Url,
                                    w => w.ServerRelativeUrl,
                                    w => w.AllProperties,
                                    w => w.WebTemplate);

                clientContext.ExecuteQueryRetry();

                string pagesListID = string.Empty;
                bool _IsPublishingWeb = IsPublishingWeb(clientContext, web);

                if (_IsPublishingWeb)
                {
                    Logger.LogInfoMessage("Web: " + web.Url + "is a publishing web", false);
                    pagesListID = web.AllProperties["__PagesListId"] as string;
                    list = web.Lists.GetById(new Guid(pagesListID));


                    clientContext.Load(list, l => l.ForceCheckout,
                                       l => l.EnableVersioning,
                                       l => l.EnableMinorVersions,
                                       l => l.EnableModeration,
                                       l => l.Title,
                                       l => l.DraftVersionVisibility,
                                       l => l.DefaultViewUrl);

                    clientContext.ExecuteQueryRetry();

                }
                else
                {
                    Logger.LogInfoMessage("Web: " + web.Url + "is not a publishing web", false);
                    clientContext.Load(web.Lists);

                    clientContext.ExecuteQueryRetry();

                    try
                    {
                        //list = web.Lists.GetByTitle(Constants.TEAMSITE_PAGES_LIBRARY);
                        //WebPageLibrary, Wiki Page Library. Value = 119.
                        IEnumerable<List> libraries = clientContext.LoadQuery(web.Lists.Where(l => l.BaseTemplate == 119));
                        clientContext.ExecuteQuery();

                        if (libraries.Any() && libraries.Count() > 0)
                        {
                            list = libraries.First();
                        }

                        clientContext.Load(list);
                        clientContext.ExecuteQueryRetry();
                    }
                    catch
                    {
                        list = null;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DeleteMissingWebparts: GetPageList]. Exception Message: " + ex.Message
                    + ", Exception Comments: Exception occured while finding page list", true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Webpart", ex.Message,
                    ex.ToString(), "GetPageList", ex.GetType().ToString(), "Exception occured while finding page list");
            }
            finally
            {
                clientContext.Dispose();
                web = null;
            }
            return list;
        }

        private static bool IsPublishingWeb(ClientContext clientContext, Web web)
        {
            Logger.LogInfoMessage("Checking if the current web is a publishing web", false);

            var _IsPublished = false;
            var propName = "__PublishingFeatureActivated";

            try
            {

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
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DeleteMissingWebparts: IsPublishingWeb]. Exception Message: "
                    + ex.Message + ", Exception Comments: Exception occured while finding publishing page", true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "Webpart", ex.Message,
                    ex.ToString(), "IsPublishingWeb", ex.GetType().ToString(), "Exception occured while finding publishing page");
            }
            finally
            {
                clientContext.Dispose();
                web = null;
            }
            return _IsPublished;
        }

        private static bool ReadInputFile(ref string webPartsInputFile)
        {
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            Logger.LogMessage("Enter Complete Input File Path of Webparts Report Either Pre-Scan OR Discovery Report:");
            System.Console.ForegroundColor = System.ConsoleColor.Yellow;
            System.Console.WriteLine("Please make sure you verify the data before executing Clean-up option as cleaned Webparts can't be rollback.");
            System.Console.ResetColor();
            webPartsInputFile = System.Console.ReadLine();
            Logger.LogMessage("[DownloadAndModifyListTemplate: ReadInputFile] Entered Input File of List Template Data " + webPartsInputFile, false);
            if (string.IsNullOrEmpty(webPartsInputFile) || !System.IO.File.Exists(webPartsInputFile))
                return false;
            return true;
        }
    }
}
