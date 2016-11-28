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
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Entities;
using System.IO;
using System.Xml.Serialization;
using System.Xml;
using WebPartTransformation;


namespace JDP.Remediation.Console
{
    public class AddWebPart
    {
        public static string filePath = string.Empty;
        public static string outputPath = string.Empty;

        public static void DoWork()
        {
            string timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");
            outputPath = Environment.CurrentDirectory;
            string webUrl = string.Empty;
            string serverRelativePageUrl = string.Empty;
            string webPartZoneIndex = string.Empty;
            string webPartZoneID = string.Empty;
            string webPartFileName = string.Empty;
            string webPartXmlFilePath = string.Empty;
            bool headerAddWebPart = false;

            //Trace Log TXT File Creation Command
            Logger.OpenLog("AddWebPart", timeStamp);

            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            //System.Console.ResetColor();
            System.Console.WriteLine("Please enter Web Url : ");
            System.Console.ResetColor();
            webUrl = System.Console.ReadLine().ToLower();
            if (string.IsNullOrEmpty(webUrl))
            {
                Logger.LogErrorMessage("[AddWebpart: DoWork]WebUrl should not be empty or null. Operation aborted...", true);
                return;
            }

            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Please enter Server Relative PageUrl (E:g- /sites/DTTesting/SitePages/WebPartPage.aspx): ");
            System.Console.ResetColor();
            serverRelativePageUrl = System.Console.ReadLine().ToLower();
            if (string.IsNullOrEmpty(serverRelativePageUrl))
            {
                Logger.LogErrorMessage("[AddWebpart: DoWork]ServerRelative PageUrl should not be empty or null. Operation aborted...", true);
                return;
            }

            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Please enter WebPart ZoneIndex : ");
            System.Console.ResetColor();
            webPartZoneIndex = System.Console.ReadLine().ToLower();
            if (string.IsNullOrEmpty(webPartZoneIndex))
            {
                Logger.LogErrorMessage("[AddWebpart: DoWork]WebPart ZoneIndex should not be empty or null. Operation aborted...", true);
                return;
            }

            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Please enter WebPart ZoneID : ");
            System.Console.ResetColor();
            webPartZoneID = System.Console.ReadLine().ToLower();
            if (string.IsNullOrEmpty(webPartZoneID))
            {
                Logger.LogErrorMessage("[AddWebpart: DoWork]WebPart ZoneId should not be empty or null. Operation aborted...", true);
                return;
            }

            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Please enter WebPart File Name (WebPart must be present in the WebPart gallery) : ");
            System.Console.ResetColor();
            webPartFileName = System.Console.ReadLine().ToLower();
            if (string.IsNullOrEmpty(webPartFileName))
            {
                Logger.LogErrorMessage("[AddWebpart: DoWork]WebPart File Name should not be empty or null. Operation aborted...", true);
                return;
            }

            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Please enter WebPart XmlFile Path : ");
            System.Console.ResetColor();
            webPartXmlFilePath = System.Console.ReadLine().ToLower();
            if (string.IsNullOrEmpty(webPartXmlFilePath) || !System.IO.File.Exists(webPartXmlFilePath))
            {
                Logger.LogErrorMessage("[AddWebpart: DoWork]WebPart XmlFile Path is not valid or available. Operation aborted...", true);
                return;
            }
            System.Console.ResetColor();
            Logger.LogInfoMessage(String.Format("Process started {0}", DateTime.Now.ToString()), true);
            try
            {

                AddWebPartStatusBase objWPOutputBase = new AddWebPartStatusBase();
                objWPOutputBase.WebApplication = Constants.NotApplicable;
                objWPOutputBase.SiteCollection = Constants.NotApplicable;
                objWPOutputBase.WebUrl = webUrl;
                objWPOutputBase.WebPartFileName = webPartFileName;
                objWPOutputBase.ZoneID = webPartZoneID;
                objWPOutputBase.ZoneIndex = webPartZoneIndex;
                objWPOutputBase.PageUrl = serverRelativePageUrl;
                objWPOutputBase.ExecutionDateTime = DateTime.Now.ToString();

                if (AddWebPartToPage(webUrl, webPartFileName, webPartXmlFilePath, webPartZoneIndex, webPartZoneID, serverRelativePageUrl, outputPath))
                {
                    System.Console.ForegroundColor = System.ConsoleColor.Green;
                    Logger.LogSuccessMessage("[AddWebPart: DoWork] Successfully Added WebPart and output file is present in the path: " + outputPath, true);
                    System.Console.ResetColor();
                    objWPOutputBase.Status = Constants.Success;
                }
                else
                {
                    Logger.LogInfoMessage("Adding WebPart to the page is failed for the site " + webUrl);
                    objWPOutputBase.Status = Constants.Failure;
                }

                if (!System.IO.File.Exists(outputPath + @"\" + Constants.AddWebPartStatusFileName + timeStamp + Constants.CSVExtension))
                {
                    headerAddWebPart = false;
                }
                else
                    headerAddWebPart = true;

                FileUtility.WriteCsVintoFile(outputPath + @"\" + Constants.AddWebPartStatusFileName + timeStamp + Constants.CSVExtension, objWPOutputBase, ref headerAddWebPart);

            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[AddWebPart: DoWork]. Exception Message: " + ex.Message, true);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "AddWebPart", ex.Message, ex.ToString(), "AddWebpart: DoWork()", ex.GetType().ToString());
            }
            Logger.LogInfoMessage(String.Format("Process completed {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        public static bool AddWebPartToPage(string webUrl, string configuredWebPartFileName, string configuredWebPartXmlFilePath, string webPartZoneIndex, string webPartZoneID, string serverRelativePageUrl, string outPutDirectory, string sourceWebPartId = "")
        {
            bool isWebPartAdded = false;
            Web web = null;
            WebPartEntity webPart = new WebPartEntity();
            ClientContext clientContext = null;
            string webPartXml = string.Empty;

            ExceptionCsv.WebUrl = webUrl;
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    web = userContext.Web;
                    userContext.Load(web);
                    userContext.ExecuteQuery();
                    clientContext = userContext;

                    webPart.WebPartIndex = Convert.ToInt32(webPartZoneIndex);

                    Logger.LogInfoMessage("[AddWebPartToPage] Successful authentication", false);

                    Logger.LogInfoMessage("[AddWebPartToPage] Checking Out File ...", false);

                    //Prepare Exception Comments
                    exceptionCommentsInfo1 = "Web Url: " + webUrl + ", Configured Web Part File Name: " + configuredWebPartFileName + " , Page Url: " + serverRelativePageUrl;

                    Logger.LogInfoMessage("[AddWebPartToPage] Successful authentication", Constants.Logging);

                    Logger.LogInfoMessage("[AddWebPartToPage] Checking for web part in the Web Part Gallery", Constants.Logging);

                    //check for the target web part in the gallery
                    bool isWebPartInGallery = CheckWebPartOrAppPartPresenceInSite(clientContext, configuredWebPartFileName, configuredWebPartXmlFilePath);

                    if (isWebPartInGallery)
                    {
                        using (System.IO.FileStream fs = new System.IO.FileStream(configuredWebPartXmlFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                        {
                            StreamReader reader = new StreamReader(fs);
                            webPart.WebPartXml = reader.ReadToEnd();
                        }

                        webPart.WebPartZone = webPartZoneID;

                        List list = GetPageList(ref clientContext);

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
                                #region Remove Versioning in List

                                Logger.LogInfoMessage(
                                   "[AddWebPart] List Details " + serverRelativePageUrl + ". " +
                                   "Force Check Out: " + forceCheckOut +
                                   "Enable Versioning: " + enableVersioning +
                                   "Enable Minor Versions: " + enableMinorVersions +
                                   "Enable Moderation: " + enableModeration +
                                   "Draft Version Visibility: " + dVisibility);

                                Logger.LogInfoMessage(
                                   "[AddWebPart] Removing Versioning");

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
                                    clientContext.ExecuteQuery();
                                }
                                #endregion
                            }
                            catch (Exception ex)
                            {
                                System.Console.ForegroundColor = System.ConsoleColor.Red;
                                Logger.LogErrorMessage("[AddWebPartToPage] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                                System.Console.ResetColor();
                                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, ExceptionCsv.WebUrl, "AddWebPart", ex.Message, ex.ToString(), "AddWebPartToPage()", ex.GetType().ToString(), exceptionCommentsInfo1);
                            }
                        }


                        try
                        {
                            Logger.LogInfoMessage("[AddWebPartToPage] Adding web part to the page at " + serverRelativePageUrl);
                            isWebPartAdded = AddWebPartt(clientContext.Web, serverRelativePageUrl, webPart, sourceWebPartId);
                        }
                        catch
                        {
                            throw;
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
                                    clientContext.ExecuteQuery();
                                }
                                #endregion
                            }
                        }

                    }

                    else
                    {
                        Logger.LogInfoMessage("[AddWebPartToPage]. Target Webpart should be present in the site for the webpart to be added", Constants.Logging);
                        throw new Exception("Target Webpart should be present in the site for the webpart to be added");
                    }

                }

            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[AddWebPartToPage] Exception Message: " + ex.Message);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, ExceptionCsv.WebUrl, "AddWebPart", ex.Message, ex.ToString(), "AddWebPartToPage()", ex.GetType().ToString());
                return isWebPartAdded;
            }
            return isWebPartAdded;
        }

        public static List GetPageList(ref ClientContext clientContext)
        {
            List list = null;

            try
            {
                Web web = clientContext.Web;

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
                    Logger.LogInfoMessage("[GetPageList] Web:  + web.Url  is a publishing web");
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
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[GetPageList] Exception Message: " + ex.Message);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "GetPageList", ex.Message, ex.ToString(), "GetPageList()", ex.GetType().ToString());
            }
            return list;
        }
        public static bool IsPublishingWeb(ClientContext clientContext, Web web)
        {
            Logger.LogInfoMessage("Checking if the current web is a publishing web");

            Logger.LogInfoMessage("Checking for PublishingFeatureActivated ...");

            var _IsPublished = false;
            var propName = "__PublishingFeatureActivated";

            try
            {

                //Ensure web properties are loaded
                if (!web.IsObjectPropertyInstantiated("AllProperties"))
                {
                    clientContext.Load(web, w => w.AllProperties);
                    clientContext.ExecuteQuery();
                    Logger.LogInfoMessage("Ensure web properties are loaded...");

                }
                //Verify whether publishing feature is activated 
                if (web.AllProperties.FieldValues.ContainsKey(propName))
                {
                    bool propVal;
                    Boolean.TryParse((string)web.AllProperties[propName], out propVal);

                    Logger.LogInfoMessage("Verify whether publishing feature is activated...");

                    _IsPublished = propVal;
                    return propVal;
                }
            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[IsPublishingWeb] Exception Message: " + ex.Message);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "IsPublishingWeb", ex.Message, ex.ToString(), "IsPublishingWeb()", ex.GetType().ToString());
            }
            return _IsPublished;
        }
        //sourceWebPartId - Used to update the content of the wikipage with new web part id 
        public static bool AddWebPartt(Web web, string serverRelativePageUrl, WebPartEntity webPartEntity, string sourceWebPartId = "")
        {
            bool isWebPartAdded = false;
            try
            {
                Microsoft.SharePoint.Client.File webPartPage = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

                web.Context.Load(webPartPage);
                web.Context.ExecuteQueryRetry();

                LimitedWebPartManager webPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);

                WebPartDefinition importedWebPart = webPartManager.ImportWebPart(webPartEntity.WebPartXml);
                WebPartDefinition webPart = webPartManager.AddWebPart(importedWebPart.WebPart, webPartEntity.WebPartZone, webPartEntity.WebPartIndex);
                web.Context.Load(webPart);
                web.Context.ExecuteQuery();

                CheckForWikiFieldOrPublishingPageContentAndUpdate(webPart, web, webPartPage, sourceWebPartId);

                isWebPartAdded = true;

            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[AddWebPartt] Exception Message: " + ex.Message);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "AddWebPart", ex.Message, ex.ToString(), "AddWebPartt()", ex.GetType().ToString());
            }

            return isWebPartAdded;
        }

        //Used to update the content of the wikipage/PublishingPageContent with new web part id 
        public static void CheckForWikiFieldOrPublishingPageContentAndUpdate(WebPartDefinition webPart, Web web, Microsoft.SharePoint.Client.File webPartPage, string sourceWebPartId = "")
        {
            string marker = String.Format(System.Globalization.CultureInfo.InvariantCulture, "<div class=\"ms-rtestate-read ms-rte-wpbox\" contentEditable=\"false\"><div class=\"ms-rtestate-read {0}\" id=\"div_{0}\"></div><div id=\"vid_{0}\"></div></div>", webPart.Id);
            ListItem item = webPartPage.ListItemAllFields;
            web.Context.Load(item);
            web.Context.ExecuteQuery();
            FieldUserValue modifiedby = (FieldUserValue)item["Editor"];
            FieldUserValue createdby = (FieldUserValue)item["Author"];
            DateTime modifiedDate = DateTime.SpecifyKind(
                                        DateTime.Parse(item["Modified"].ToString()),
                                        DateTimeKind.Utc);

            DateTime createdDate = DateTime.SpecifyKind(
                                        DateTime.Parse(item["Created"].ToString()),
                                        DateTimeKind.Utc);

            item["Editor"] = modifiedby.LookupId;
            item["Author"] = createdby.LookupId;
            item["Modified"] = modifiedDate;
            item["Created"] = createdDate;

            try
            {
                if (item["WikiField"] != null)
                {
                    string markerToBeUpdated = item["WikiField"].ToString();
                    if (!string.IsNullOrEmpty(sourceWebPartId))
                    {
                        string updatedWikiField = markerToBeUpdated.Replace(sourceWebPartId, webPart.Id.ToString());
                        if (!markerToBeUpdated.Equals(updatedWikiField))
                        {
                            item["WikiField"] = updatedWikiField;
                        }
                    }
                    else
                    {
                        string parentTag = markerToBeUpdated.Substring(markerToBeUpdated.LastIndexOf("<div class=\"ExternalClass"), (markerToBeUpdated.IndexOf("<div class=\"ExternalClass") + markerToBeUpdated.IndexOf("\">")) + 2);
                        int strPos = markerToBeUpdated.IndexOf(parentTag) + parentTag.Length;

                        StringBuilder markerOfAllWebparts = new StringBuilder();
                        markerOfAllWebparts.AppendLine(markerToBeUpdated.Substring(0, strPos));
                        markerOfAllWebparts.AppendLine(marker);
                        markerOfAllWebparts.AppendLine(markerToBeUpdated.Substring(strPos, (markerToBeUpdated.Length - strPos)));

                        item["WikiField"] = markerOfAllWebparts.ToString();
                    }
                }


            }
            catch
            {
                try
                {
                    if (item["PublishingPageContent"] != null)
                    {
                        string markerToBeUpdated = item["PublishingPageContent"].ToString();
                        if (!string.IsNullOrEmpty(sourceWebPartId))
                        {
                            string updatedWikiField = markerToBeUpdated.Replace(sourceWebPartId, webPart.Id.ToString());
                            if (!markerToBeUpdated.Equals(updatedWikiField))
                            {
                                item["PublishingPageContent"] = updatedWikiField;
                            }
                        }
                        else
                        {
                            item["PublishingPageContent"] = markerToBeUpdated + marker;
                        }
                    }
                }
                catch
                {
                    //do nothing
                }
            }

            item.Update();
            web.Context.ExecuteQuery();
        }
        private static bool CheckWebPartOrAppPartPresenceInSite(ClientContext clientContext, string targetWebPartXmlFileName, string targetWebPartXmlFilePath)
        {
            bool isWebPartInSite = false;

            webParts targetWebPart = null;

            string webPartPropertiesXml = string.Empty;

            string webPartType = string.Empty;

            clientContext.Load(clientContext.Web);
            clientContext.ExecuteQuery();

            ExceptionCsv.WebUrl = clientContext.Web.Url;
            string exceptionCommentsInfo1 = string.Empty;

            try
            {

                //Prepare Exception Comments
                exceptionCommentsInfo1 = "Web Url: " + clientContext.Web.Url + ", Target Web Part File Name: " + targetWebPartXmlFileName + " , Target WebPart Xml File Path: " + targetWebPartXmlFilePath;

                using (System.IO.FileStream fs = new System.IO.FileStream(targetWebPartXmlFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                {
                    System.IO.StreamReader reader = new System.IO.StreamReader(fs);
                    webPartPropertiesXml = reader.ReadToEnd();
                }

                Logger.LogInfoMessage("[CheckWebPartOrAppPartPresenceInSite] Checking for web part schema version");


                if (webPartPropertiesXml.Contains("WebPart/v2"))
                {
                    Logger.LogInfoMessage("[CheckWebPartOrAppPartPresenceInSite] Web part schema version is V2");


                    XmlDataDocument xmldoc = new XmlDataDocument();
                    xmldoc.LoadXml(webPartPropertiesXml);
                    webPartType = GetWebPartShortTypeName(xmldoc.DocumentElement.GetElementsByTagName("TypeName").Item(0).FirstChild.Value);

                    Logger.LogInfoMessage("[CheckWebPartOrAppPartPresenceInSite] Web part Type: " + webPartType);

                    xmldoc = null;
                }
                else
                {
                    Logger.LogInfoMessage("[CheckWebPartOrAppPartPresenceInSite] Web part schema version is V3");

                    using (System.IO.FileStream fs = new System.IO.FileStream(targetWebPartXmlFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                    {
                        XmlReader reader = new XmlTextReader(fs);
                        XmlSerializer serializer = new XmlSerializer(typeof(webParts));
                        targetWebPart = (webParts)serializer.Deserialize(reader);
                        if (targetWebPart != null)
                        {
                            webPartType = GetWebPartShortTypeName(targetWebPart.webPart.metaData.type.name);

                            Logger.LogInfoMessage("[CheckWebPartOrAppPartPresenceInSite] Web part Type: " + webPartType);
                        }
                    }
                }

                if (webPartType.Equals("ClientWebPart", StringComparison.CurrentCultureIgnoreCase))
                {
                    foreach (var item in targetWebPart.webPart.data.properties)
                    {
                        if (item.name.Equals("FeatureId", StringComparison.CurrentCultureIgnoreCase))
                        {
                            Guid featureID = new Guid(item.Value);
                            isWebPartInSite = IsFeatureOnWeb(featureID, clientContext);
                            break;
                        }
                    }

                }
                else
                {
                    Web web = clientContext.Site.RootWeb;
                    clientContext.Load(web, w => w.Url);
                    clientContext.ExecuteQuery();

                    //List list = web.Lists.GetByTitle("Web Part Gallery");
                    //WebPartCatalog, Web Part gallery. Value = 113.

                    List list = null;
                    IEnumerable<List> libraries = clientContext.LoadQuery(web.Lists.Where(l => l.BaseTemplate == 113));
                    clientContext.ExecuteQuery();

                    if (libraries.Any() && libraries.Count() > 0)
                    {
                        list = libraries.First();
                    }

                    clientContext.Load(list);
                    clientContext.ExecuteQueryRetry();

                    CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery(1000);
                    Microsoft.SharePoint.Client.ListItemCollection items = list.GetItems(camlQuery);
                    clientContext.Load(items);
                    clientContext.ExecuteQuery();

                    foreach (ListItem item in items)
                    {
                        if (item["FileLeafRef"].ToString().Equals(targetWebPartXmlFileName, StringComparison.CurrentCultureIgnoreCase))
                        {
                            isWebPartInSite = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[CheckWebPartOrAppPartPresenceInSite] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, ExceptionCsv.WebUrl, "AddWebPart", ex.Message, ex.ToString(), "CheckWebPartOrAppPartPresenceInSite()", ex.GetType().ToString(), exceptionCommentsInfo1);
                return isWebPartInSite;
            }

            return isWebPartInSite;
        }

        public static string GetWebPartShortTypeName(string webPartType)
        {
            string _webPartType = string.Empty;
            try
            {
                string[] tempWebPartTypeName = webPartType.Split(',');

                string[] tempWebPartType = tempWebPartTypeName[0].Split('.');
                if (tempWebPartType.Length == 1)
                {
                    _webPartType = tempWebPartType[0];
                }
                else
                {
                    _webPartType = tempWebPartType[tempWebPartType.Length - 1];
                }
            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[GetWebPartShortTypeName] Exception Message: " + ex.Message);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "GetWebPartShortTypeName", ex.Message, ex.ToString(), "GetWebPartShortTypeName()", ex.GetType().ToString());
            }
            return _webPartType;
        }
        private static bool IsFeatureOnWeb(Guid FeatureID, ClientContext clientContext)
        {
            bool isFeatureAvailable = false;
            try
            {
                FeatureCollection features = clientContext.Web.Features;
                clientContext.Load(features);
                clientContext.ExecuteQuery();

                Feature feature = features.GetById(FeatureID);
                if (feature != null)
                {
                    clientContext.Load(feature);
                    clientContext.ExecuteQuery();
                    if (feature.DefinitionId != null)
                    {
                        isFeatureAvailable = true;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[IsFeatureOnWeb] Exception Message: " + ex.Message);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "AddWebPart", ex.Message, ex.ToString(), "IsFeatureOnWeb()", ex.GetType().ToString());
            }


            return isFeatureAvailable;
        }

    }
}
