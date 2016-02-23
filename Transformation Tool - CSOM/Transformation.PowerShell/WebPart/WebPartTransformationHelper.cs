using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client.WebParts;
using Microsoft.SharePoint.Client;
using Transformation.PowerShell.Base;
using Transformation.PowerShell.Common;
using System.IO;
using System.Xml;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Entities;
using Transformation.PowerShell.WebPartPagesService;
using WebPartTransformation;
using System.Xml.Serialization;
using Transformation.PowerShell.Common.CSV;
using Transformation.PowerShell.Common.Utilities;
using OfficeDevPnP.Core.Framework.TimerJobs;
using System.Reflection;

namespace Transformation.PowerShell.WebPart
{
    public class WebPartTransformationHelper : TimerJob
    {
        public string WebPartType { get; set; }

        public string OutPutDirectory { get; set; }

        public bool headerWebPart = false;

        public WebPartTransformationHelper()
            : base("")
        {

        }

        public WebPartTransformationHelper(string name)
            : base(name)
        {
            if (name.Equals("GetWebPartUsage"))
            {
               
                TimerJobRun += GetWebPartUsage_TimerJobRun;
            }
        }
        
        public void WebPart_Initialization(string DiscoveryUsage_OutPutFolder)
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

        }
        
        public void DeleteUsageFiles_WebPartHelper(string outPutFolder, string fileName)
        {
            //Delete Usage File
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
            Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());
            FileUtility.DeleteFiles(outPutFolder + @"\" + fileName);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[DeleteUsageFiles_WebPartHelper] Deleted the Usage file : " + fileName + " from : " + outPutFolder);
            Console.WriteLine("[DeleteUsageFiles_WebPartHelper] Deleted the Usage file : " + fileName + " from : " + outPutFolder);

        }
        
        void GetWebPartUsage_TimerJobRun(object sender, TimerJobRunEventArgs e)
        {
            e.WebClientContext.Load(e.WebClientContext.Web, p => p.Url);
            e.WebClientContext.ExecuteQueryRetry();
            GetWebPartUsage(WebPartType, e.WebClientContext, OutPutDirectory);
        }

        public void GetWebPartUsage(string webPartType, ClientContext clientContext,string outPutDirectory)
        {
            ExceptionCsv.WebUrl = clientContext.Web.Url;
            string exceptionCommentsInfo1 = string.Empty;
            string webPartUsageFileName = outPutDirectory + "\\" + Constants.WEBPART_USAGE_ENTITY_FILENAME;
            try
            {                
                //Initialized Exception and Logger. 
                WebPart_Initialization(outPutDirectory);

                string webUrl = clientContext.Web.Url;
                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started ##############");
                Console.WriteLine("############## Web Part  Trasnformation Utility Execution Started ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());
                
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][GetWebPartUsage]");
                Console.WriteLine("[START][GetWebPartUsage]");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartUsage] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
                Console.WriteLine("[GetWebPartUsage] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartUsage] Finding WebPartUsage details for Web Part: " + webPartType + " in Web: " + webUrl);
                Console.WriteLine("[GetWebPartUsage] Finding WebPartUsage details for Web Part: " + webPartType + " in Web: " + webUrl);

                WebPartUsageEntity webPartUsageEntity = null;
                List<WebPartUsageEntity> webPartUsage = new List<WebPartUsageEntity>();
                //Prepare Exception Comments
                exceptionCommentsInfo1 = "Web Url: " + clientContext.Web.Url + ", Web Part Type: " + webPartType;

                if (clientContext != null)
                {
                    List list = GetPageList(ref clientContext);
                    if (list != null)
                    {
                        var items = list.GetItems(CamlQuery.CreateAllItemsQuery());

                        //make sure to include the File on each Item fetched
                        clientContext.Load(items,
                                            i => i.Include(
                                                    item => item.File,
                                                     item => item["EncodedAbsUrl"]));
                        clientContext.ExecuteQuery();

                        // Iterate through all available pages in the pages list
                        foreach (var item in items)
                        {
                            Microsoft.SharePoint.Client.File page = item.File;

                            String pageUrl = page.ServerRelativeUrl;// item.FieldValues["EncodedAbsUrl"].ToString();

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartUsage] Checking for the Web Part on the Page: " + page.Name);
                            Console.WriteLine("[GetWebPartUsage] Checking for the Web Part on the Page:" + page.Name);

                            // Requires Full Control permissions on the Web
                            LimitedWebPartManager webPartManager = page.GetLimitedWebPartManager(PersonalizationScope.Shared);
                            clientContext.Load(webPartManager,
                                                wpm => wpm.WebParts,
                                                wpm => wpm.WebParts.Include(
                                                                    wp => wp.WebPart.Hidden,
                                                                    wp => wp.WebPart.IsClosed,
                                                                    wp => wp.WebPart.Properties,
                                                                    wp => wp.WebPart.Subtitle,
                                                                    wp => wp.WebPart.Title,
                                                                    wp => wp.WebPart.TitleUrl,
                                                                    wp => wp.WebPart.ZoneIndex));
                            clientContext.ExecuteQuery();

                            foreach (WebPartDefinition webPartDefinition in webPartManager.WebParts)
                            {
                                Microsoft.SharePoint.Client.WebParts.WebPart webPart = webPartDefinition.WebPart;

                                string webPartPropertiesXml = GetWebPartPropertiesServiceCall(clientContext, webPartDefinition.Id.ToString(), pageUrl);

                                string WebPartTypeName = string.Empty;

                                if (webPartPropertiesXml.Contains("WebPart/v2"))
                                {
                                    XmlDataDocument xmldoc = new XmlDataDocument();
                                    xmldoc.LoadXml(webPartPropertiesXml);
                                    WebPartTypeName = xmldoc.DocumentElement.GetElementsByTagName("TypeName").Item(0).FirstChild.Value;
                                    xmldoc = null;
                                }
                                else
                                {
                                    webParts webPartProp = null;
                                    byte[] byteArray = Encoding.UTF8.GetBytes(webPartPropertiesXml);
                                    using (MemoryStream stream = new MemoryStream(byteArray))
                                    {
                                        StreamReader streamReader = new StreamReader(stream);
                                        System.Xml.XmlReader reader = new XmlTextReader(streamReader);
                                        XmlSerializer serializer = new XmlSerializer(typeof(webParts));
                                        webPartProp = (webParts)serializer.Deserialize(reader);
                                        WebPartTypeName = webPartProp.webPart.metaData.type.name;
                                        stream.Flush();
                                    }
                                    byteArray = null;
                                }

                                string actualWebPartType = GetWebPartShortTypeName(WebPartTypeName);

                                // only modify if we find the old web part
                                if (actualWebPartType.Equals(webPartType))
                                {
                                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartUsage] Found WebPart: " + webPartType + " in Page: " + page.Name);
                                    Console.WriteLine("[GetWebPartUsage] Found WebPart: " + webPartType + " in Page: " + page.Name);

                                    webPartUsageEntity = new WebPartUsageEntity();
                                    webPartUsageEntity.PageUrl = pageUrl;
                                    webPartUsageEntity.WebPartID = webPartDefinition.Id.ToString();
                                    webPartUsageEntity.WebURL = webUrl;
                                    webPartUsageEntity.WebPartTitle = webPart.Title;
                                    webPartUsageEntity.ZoneIndex = webPart.ZoneIndex.ToString();
                                    webPartUsageEntity.WebPartType = actualWebPartType;

                                    FileUtility.WriteCsVintoFile(webPartUsageFileName, webPartUsageEntity, ref headerWebPart);
                                }
                            }
                        }
                    }
                }
                //Default Pages
                GetWebPartUsage_DefaultPages(webPartType, clientContext, outPutDirectory);
                //Default Pages

                Console.WriteLine("[GetWebPartUsage] WebPart Usage is exported to the file " + webPartUsageFileName);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartUsage] WebPart Usage is exported to the file " + webPartUsageFileName);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][GetWebPartUsage]");
                Console.WriteLine("[END][GetWebPartUsage]");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part Trasnformation Utility Execution Completed for Web ##############");
                Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  for Web ##############");
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "GetWebPartUsage", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][GetWebPartUsage] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][GetWebPartUsage] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        public void GetWebPartUsage_DefaultPages(string webPartType, ClientContext clientContext, string outPutDirectory)
        {
            ExceptionCsv.WebUrl = clientContext.Web.Url;
            string exceptionCommentsInfo1 = string.Empty;
            string webPartUsageFileName = outPutDirectory + "\\" + Constants.WEBPART_USAGE_ENTITY_FILENAME;

            try
            {
                string webUrl = clientContext.Web.Url;

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartUsage_DefaultPages] Finding WebPartUsage details for Web Part: " + webPartType + " in Web: " + webUrl);
                Console.WriteLine("[GetWebPartUsage_DefaultPages] Finding WebPartUsage details for Web Part: " + webPartType + " in Web: " + webUrl);

                WebPartUsageEntity webPartUsageEntity = null;
                List<WebPartUsageEntity> webPartUsage = new List<WebPartUsageEntity>();

                if (clientContext != null)
                {
                    clientContext.Load(clientContext.Web.RootFolder.Files);
                    clientContext.ExecuteQuery();

                    foreach (Microsoft.SharePoint.Client.File page in clientContext.Web.RootFolder.Files)
                    {
                        exceptionCommentsInfo1 = "Web Url: " + clientContext.Web.Url + ", Web Part Type: " + webPartType + ", PageTitle: " + page.ServerRelativeUrl;

                        if (Path.GetExtension(page.ServerRelativeUrl).Equals(".aspx", StringComparison.CurrentCultureIgnoreCase))
                        {
                            String pageUrl = page.ServerRelativeUrl;

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartUsage_DefaultPages] Checking for the Web Part on the Page: " + page.Name);
                            Console.WriteLine("[GetWebPartUsage_DefaultPages] Checking for the Web Part on the Page:" + page.Name);

                            // Requires Full Control permissions on the Web
                            LimitedWebPartManager webPartManager = page.GetLimitedWebPartManager(PersonalizationScope.Shared);
                            clientContext.Load(webPartManager,
                                                wpm => wpm.WebParts,
                                                wpm => wpm.WebParts.Include(
                                                                    wp => wp.WebPart.Hidden,
                                                                    wp => wp.WebPart.IsClosed,
                                                                    wp => wp.WebPart.Properties,
                                                                    wp => wp.WebPart.Subtitle,
                                                                    wp => wp.WebPart.Title,
                                                                    wp => wp.WebPart.TitleUrl,
                                                                    wp => wp.WebPart.ZoneIndex));
                            clientContext.ExecuteQuery();

                            foreach (WebPartDefinition webPartDefinition in webPartManager.WebParts)
                            {
                                Microsoft.SharePoint.Client.WebParts.WebPart webPart = webPartDefinition.WebPart;

                                string webPartPropertiesXml = GetWebPartPropertiesServiceCall(clientContext, webPartDefinition.Id.ToString(), pageUrl);

                                string WebPartTypeName = string.Empty;

                                if (webPartPropertiesXml.Contains("WebPart/v2"))
                                {
                                    XmlDataDocument xmldoc = new XmlDataDocument();
                                    xmldoc.LoadXml(webPartPropertiesXml);
                                    WebPartTypeName = xmldoc.DocumentElement.GetElementsByTagName("TypeName").Item(0).FirstChild.Value;
                                    xmldoc = null;
                                }
                                else
                                {
                                    webParts webPartProp = null;
                                    byte[] byteArray = Encoding.UTF8.GetBytes(webPartPropertiesXml);
                                    using (MemoryStream stream = new MemoryStream(byteArray))
                                    {
                                        StreamReader streamReader = new StreamReader(stream);
                                        System.Xml.XmlReader reader = new XmlTextReader(streamReader);
                                        XmlSerializer serializer = new XmlSerializer(typeof(webParts));
                                        webPartProp = (webParts)serializer.Deserialize(reader);
                                        WebPartTypeName = webPartProp.webPart.metaData.type.name;
                                        stream.Flush();
                                    }
                                    byteArray = null;
                                }

                                string actualWebPartType = GetWebPartShortTypeName(WebPartTypeName);

                                // only modify if we find the old web part
                                if (actualWebPartType.Equals(webPartType))
                                {
                                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartUsage_DefaultPages] Found WebPart: " + webPartType + " in Page: " + page.Name + ", " + page.ServerRelativeUrl);
                                    Console.WriteLine("[GetWebPartUsage_DefaultPages] Found WebPart: " + webPartType + " in Page: " + page.Name + ", " + page.ServerRelativeUrl);

                                    webPartUsageEntity = new WebPartUsageEntity();
                                    webPartUsageEntity.PageUrl = pageUrl;
                                    webPartUsageEntity.WebPartID = webPartDefinition.Id.ToString();
                                    webPartUsageEntity.WebURL = webUrl;
                                    webPartUsageEntity.WebPartTitle = webPart.Title;
                                    webPartUsageEntity.ZoneIndex = webPart.ZoneIndex.ToString();
                                    webPartUsageEntity.WebPartType = actualWebPartType;

                                    FileUtility.WriteCsVintoFile(webPartUsageFileName, webPartUsageEntity, ref headerWebPart);
                                }
                            }
                        }
                    }
                }

                Console.WriteLine("[GetWebPartUsage_DefaultPages] Default Pages WebPart Usage is exported to the file " + webPartUsageFileName);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartUsage_DefaultPages] Default Pages WebPart Usage is exported to the file " + webPartUsageFileName);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][GetWebPartUsage_DefaultPages]");
                Console.WriteLine("[END][GetWebPartUsage_DefaultPages]");
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "GetWebPartUsage_DefaultPages", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][GetWebPartUsage_DefaultPages] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][GetWebPartUsage_DefaultPages] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        
        public void GetWebPartsByWeb(string WebUrl, string outPutDirectory, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            ClientContext clientContext = new ClientContext(WebUrl);
            AuthenticationHelper ObjAuth = new AuthenticationHelper();

            //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
            if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.OnPremise.ToUpper()))
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][GetWebPartsByWeb] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(WebUrl, UserName, Password, Domain);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][GetWebPartsByWeb] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
            }
            //SharePointOnline  => OL (Online)
            else if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.Online.ToUpper()))
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][GetWebPartsByWeb] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(WebUrl, UserName, Password);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][GetWebPartsByWeb] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
            }

            if (clientContext != null)
            {
                clientContext.Load(clientContext.Web);
                clientContext.ExecuteQuery();

                ExceptionCsv.WebUrl = clientContext.Web.Url;
                string exceptionCommentsInfo1 = string.Empty;
                string webPartInfoFileName = outPutDirectory + "\\" + "WebParts.csv";

                try
                {
                    //Initialized Exception and Logger. 
                    WebPart_Initialization(outPutDirectory);

                    //Deleted the Web Part Usage File
                    DeleteUsageFiles_WebPartHelper(outPutDirectory, "WebParts.csv");

                    string webUrl = clientContext.Web.Url;
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started ##############");
                    Console.WriteLine("############## Web Part  Trasnformation Utility Execution Started ##############");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                    Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][GetWebPartsByWeb]");
                    Console.WriteLine("[START][GetWebPartsByWeb]");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartsByWeb] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
                    Console.WriteLine("[GetWebPartsByWeb] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);

                    WebPartInfo webPartInfo = null;

                    List list = GetPageList(ref clientContext);

                    if (list != null)
                    {
                        var items = list.GetItems(CamlQuery.CreateAllItemsQuery());

                        //make sure to include the File on each Item fetched
                        clientContext.Load(items,
                                            i => i.Include(
                                                    item => item.File,
                                                     item => item["EncodedAbsUrl"]));
                        clientContext.ExecuteQuery();

                        bool headerWebPart = false;

                        // Iterate through all available pages in the pages list
                        foreach (var item in items)
                        {
                            Microsoft.SharePoint.Client.File page = item.File;

                            String pageUrl = page.ServerRelativeUrl;// item.FieldValues["EncodedAbsUrl"].ToString();

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartsByWeb] Checking for the Web Part on the Page: " + page.Name);
                            Console.WriteLine("[GetWebPartsByWeb] Checking for the Web Part on the Page:" + page.Name);

                            // Requires Full Control permissions on the Web
                            LimitedWebPartManager webPartManager = page.GetLimitedWebPartManager(PersonalizationScope.Shared);
                            clientContext.Load(webPartManager,
                                                wpm => wpm.WebParts,
                                                wpm => wpm.WebParts.Include(
                                                                    wp => wp.WebPart.Hidden,
                                                                    wp => wp.WebPart.IsClosed,
                                                                    wp => wp.WebPart.Properties,
                                                                    wp => wp.WebPart.Subtitle,
                                                                    wp => wp.WebPart.Title,
                                                                    wp => wp.WebPart.TitleUrl,
                                                                    wp => wp.WebPart.ZoneIndex));
                            clientContext.ExecuteQuery();

                            foreach (WebPartDefinition webPartDefinition in webPartManager.WebParts)
                            {
                                Microsoft.SharePoint.Client.WebParts.WebPart webPart = webPartDefinition.WebPart;

                                //Prepare Exception Comments
                                exceptionCommentsInfo1 = "Web Url: " + clientContext.Web.Url + ", Web Part Type: " + webPart.GetType().Name;

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartsByWeb] Found WebPart details for Web Part: " + webPart.Title + " in Web: " + webUrl);
                                Console.WriteLine("[GetWebPartsByWeb] Found WebPart details for Web Part: " + webPart.Title + " in Web: " + webUrl);

                                webPartInfo = new WebPartInfo();
                                webPartInfo.PageUrl = pageUrl;
                                webPartInfo.WebPartID = webPartDefinition.Id.ToString();
                                webPartInfo.WebURL = webUrl;
                                webPartInfo.WebPartTitle = webPart.Title;
                                webPartInfo.ZoneIndex = webPart.ZoneIndex.ToString();
                                webPartInfo.WebPartTitle = webPart.Title;
                                string webPartPropertiesXml = GetWebPartPropertiesServiceCall(clientContext, webPartDefinition.Id.ToString(), pageUrl);

                                string WebPartTypeName = string.Empty;

                                if (webPartPropertiesXml.Contains("WebPart/v2"))
                                {
                                    XmlDataDocument xmldoc = new XmlDataDocument();
                                    xmldoc.LoadXml(webPartPropertiesXml);
                                    WebPartTypeName = xmldoc.DocumentElement.GetElementsByTagName("TypeName").Item(0).FirstChild.Value;
                                    xmldoc = null;
                                }
                                else
                                {
                                    webParts webPartProp = null;
                                    byte[] byteArray = Encoding.UTF8.GetBytes(webPartPropertiesXml);
                                    using (MemoryStream stream = new MemoryStream(byteArray))
                                    {
                                        StreamReader streamReader = new StreamReader(stream);
                                        System.Xml.XmlReader reader = new XmlTextReader(streamReader);
                                        XmlSerializer serializer = new XmlSerializer(typeof(webParts));
                                        webPartProp = (webParts)serializer.Deserialize(reader);
                                        WebPartTypeName = webPartProp.webPart.metaData.type.name;
                                        stream.Flush();
                                    }
                                    byteArray = null;
                                }

                                string actualWebPartType = GetWebPartShortTypeName(WebPartTypeName);

                                webPartInfo.WebPartTypeFullName = actualWebPartType;
                                webPartInfo.IsClosed = webPart.IsClosed;
                                webPartInfo.IsVisible = webPart.GetType().IsVisible;

                                FileUtility.WriteCsVintoFile(webPartInfoFileName, webPartInfo, ref headerWebPart);
                            }
                        }
                    }

                    Console.WriteLine("[GetWebPartsByWeb] WebPart Usage is exported to the file " + webPartInfoFileName);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartsByWeb] WebPart Usage is exported to the file " + webPartInfoFileName);

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][GetWebPartsByWeb]");
                    Console.WriteLine("[END][GetWebPartsByWeb]");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part Trasnformation Utility Execution Completed for Web ##############");
                    Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  for Web ##############");

                }
                catch (Exception ex)
                {
                    ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "GetWebPartsByWeb", ex.GetType().ToString(), exceptionCommentsInfo1);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][GetWebPartsByWeb] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("[EXCEPTION][GetWebPartsByWeb] Exception Message: " + ex.Message);
                    Console.ForegroundColor = ConsoleColor.Gray;
                }
            }
        }

        public void GetWebPartProperties_UsingCSV(string sourceWebPartType, string usageFilePath, string outPutFolder, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A", string ActionType="")
        {
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                if (ActionType.ToLower() != Constants.ActionType_All.ToLower())
                {
                    WebPart_Initialization(outPutFolder);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started - GetWebPartProperties_UsingCSV - For InputCSV ##############");
                    Console.WriteLine("############## Web Part Trasnformation Utility Execution Started - GetWebPartProperties_UsingCSV - For InputCSV ##############");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                    Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartProperties_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);
                    Console.WriteLine("[GetWebPartProperties_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);
                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][GetWebPartProperties_UsingCSV]");
                Console.WriteLine("[START][GetWebPartProperties_UsingCSV]");
                //Reading Input File
                IEnumerable<WebPartDiscoveryInput> objWPDInput;

                ReadWebPartUsageCSV(sourceWebPartType, usageFilePath, outPutFolder, out objWPDInput, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                for (int i = 0; i < objWPDInput.Count(); i++)
                {
                    try
                    {
                        WebPartDiscoveryInput objInput = objWPDInput.ElementAt(i);
                        SetExceptionCSVWebAppSiteColWebUrls(objInput.WebUrl.ToString(), SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                        //This is for Exception Comments:
                        exceptionCommentsInfo1 = "WebUrl: " + objInput.WebUrl + ", PageUrl: " + objInput.PageUrl + ", WebPartID: " + objInput.WebPartId + ", StorageKey: " + objInput.StorageKey;
                        //This is for Exception Comments:

                        //This function is Get Relative URL of the page
                        string _relativePageUrl = string.Empty;
                        _relativePageUrl = GetPageRelativeURL(objInput.WebUrl.ToString(), objInput.PageUrl.ToString(), SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                        string _StorageKey = string.Empty;
                        _StorageKey = GetWebPartID(objInput.StorageKey);

                        GetWebPartProperties(_relativePageUrl, _StorageKey, objInput.WebUrl, outPutFolder, SharePointOnline_OR_OnPremise, UserName, Password, Domain, "CSVUpdates");
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Error in Processing Web:" + ExceptionCsv.WebUrl + " `\r\nError Details:" + ex.Message + " `\r\nExceptionComments:" + exceptionCommentsInfo1);
                        Console.ForegroundColor = ConsoleColor.Gray;
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception][GetWebPartProperties_UsingCSV] Exception Message: " + ex.Message);
                        ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "GetWebPartProperties_UsingCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                    }
                }

                Console.WriteLine("[END][GetWebPartProperties_UsingCSV]");
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][GetWebPartProperties_UsingCSV]");

                if (ActionType.ToLower() != Constants.ActionType_All.ToLower())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "##############  Web Part Trasnformation Utility Execution Completed  - GetWebPartProperties_UsingCSV - InputCSV ##############");
                    Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  - GetWebPartProperties_UsingCSV - InputCSV ##############");
                }
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception][GetWebPartProperties_UsingCSV] Exception Message: " + ex.Message + " Exception Comments: " + exceptionCommentsInfo1);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "GetWebPartProperties_UsingCSV", ex.GetType().ToString(), exceptionCommentsInfo1.ToString());

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception][GetWebPartProperties_UsingCSV] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        
        public string GetWebPartProperties(string pageUrl, string StorageKey, string webUrl, string outPutDirectory, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A", string ActionType = "web")
        {
            string webPartPropertiesFileName = string.Empty;

            AuthenticationHelper ObjAuth = new AuthenticationHelper();
            ClientContext clientContext = new ClientContext(webUrl);
            
            string webPartXml = string.Empty;
            ExceptionCsv.WebUrl = webUrl;
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                if (ActionType.ToLower().Trim() == Constants.ActionType_Web.ToLower())
                {
                    //Initialized Exception and Logger. 
                    WebPart_Initialization(outPutDirectory);

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started - GetWebPartProperties ##############");
                    Console.WriteLine("############## Web Part  Trasnformation Utility Execution Started - GetWebPartProperties ##############");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                    Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][GetWebPartProperties] ");
                    Console.WriteLine("[START][GetWebPartProperties] ");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartProperties] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
                    Console.WriteLine("[GetWebPartProperties] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
                }
               
                string sourceWebPartXmlFilesDir = outPutDirectory + @"\" + Constants.SOURCE_WEBPART_XML_DIR;

                if (!System.IO.Directory.Exists(sourceWebPartXmlFilesDir))
                {
                    System.IO.Directory.CreateDirectory(sourceWebPartXmlFilesDir);
                }

                //Deleted the Web Part Usage File
                DeleteUsageFiles_WebPartHelper(sourceWebPartXmlFilesDir, StorageKey + "_" + Constants.WEBPART_PROPERTIES_FILENAME);
                
                //Prepare Exception Comments
                exceptionCommentsInfo1 = "Web Url: " + webUrl + ", Page Url: " + pageUrl + ", StorageKey" + StorageKey;

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.OnPremise.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][GetWebPartProperties] GetNetworkCredentialAuthenticatedContext for WebUrl: " + webUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(webUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][GetWebPartProperties] GetNetworkCredentialAuthenticatedContext for WebUrl: " + webUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.Online.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][GetWebPartProperties] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + webUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(webUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][GetWebPartProperties] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + webUrl);
                }

                if (clientContext != null)
                {
                    Web web = clientContext.Web;
                    clientContext.Load(web,w => w.Url);

                    clientContext.ExecuteQueryRetry();

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartProperties] Retrieving WebPart Properties for StorageKey: " + StorageKey.ToString() + " in the Page" + pageUrl);
                    Console.WriteLine("[GetWebPartProperties] Retrieving WebPart Properties for StorageKey: " + StorageKey.ToString() + " in the Page" + pageUrl);

                    var service = new Transformation.PowerShell.WebPartPagesService.WebPartPagesWebService();
                    service.Url = clientContext.Web.Url + Constants.WEBPART_SERVICE;

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartProperties] Service Url used to retrieve WebPart Properties : " + service.Url);
                    Console.WriteLine("[GetWebPartProperties] Service Url used to retrieve WebPart Properties : " + service.Url);

                    service.PreAuthenticate = true;

                    service.Credentials = clientContext.Credentials;
                    
                    //For Publishing Pages, Pass - WebPartID
                    //For SitePage or Team Site, Pass - StorageKey.ToGuid()
                    webPartXml = service.GetWebPart2(pageUrl, StorageKey.ToGuid(), Storage.Shared, SPWebServiceBehavior.Version3);                    

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartProperties] Successfully retreived Web Part Properties");
                    Console.WriteLine("[GetWebPartProperties] Successfully retreived Web Part Properties");

                    webPartPropertiesFileName = sourceWebPartXmlFilesDir + "\\" + StorageKey + "_" + Constants.WEBPART_PROPERTIES_FILENAME;

                    using (StreamWriter fsWebPartProperties = new StreamWriter(webPartPropertiesFileName))
                    {
                        fsWebPartProperties.WriteLine(webPartXml);
                        fsWebPartProperties.Flush();
                    }

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[GetWebPartProperties] WebPart Properties in xml format is exported to the file " + webPartPropertiesFileName);
                    Console.WriteLine("[GetWebPartProperties] WebPart Properties in xml format is exported to the file " + webPartPropertiesFileName);
                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][GetWebPartProperties]");
                Console.WriteLine("[END][GetWebPartProperties]");
                
                if (ActionType.ToLower().Trim() == Constants.ActionType_Web.ToLower())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part Trasnformation Utility Execution Completed for Web - GetWebPartProperties ##############");
                    Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  for Web - GetWebPartProperties ##############");
                }

            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "GetWebPartProperties", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][GetWebPartProperties] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][GetWebPartProperties] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            return webPartPropertiesFileName;
        }

        private void UploadDisplayTemplateFile(string webUrl, string fileName, string localFilePath, bool overwriteIfExists, string displayTemplateCategory, string outPutDirectory, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            ClientContext clientContext = new ClientContext(webUrl);
            AuthenticationHelper ObjAuth = new AuthenticationHelper();
            string webPartXml = string.Empty;
            Web web = clientContext.Web;
            Folder folder = null;

            ExceptionCsv.WebUrl = webUrl;
            string exceptionCommentsInfo1 = string.Empty;
            try
            {
                //Initialized Exception and Logger. 
                WebPart_Initialization(outPutDirectory);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started - UploadDisplayTemplateFile ##############");
                Console.WriteLine("############## Web Part  Trasnformation Utility Execution Started - UploadDisplayTemplateFile ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UploadDisplayTemplateFile]");
                Console.WriteLine("[START][UploadDisplayTemplateFile]");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadDisplayTemplateFile] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
                Console.WriteLine("[UploadDisplayTemplateFile] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);

                //Prepare Exception Comments
                exceptionCommentsInfo1 = "Web Url: " + webUrl + ", Display Template File Name : " + fileName + ", Category: " + displayTemplateCategory;

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.OnPremise.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UploadDisplayTemplateFile] GetNetworkCredentialAuthenticatedContext for WebUrl: " + webUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(webUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UploadDisplayTemplateFile] GetNetworkCredentialAuthenticatedContext for WebUrl: " + webUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.Online.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UploadDisplayTemplateFile]GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + webUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(webUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UploadDisplayTemplateFile] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + webUrl);
                }

                if (clientContext != null)
                {
                    folder = clientContext.Web.Lists.GetByTitle("Master Page Gallery").RootFolder.ResolveSubFolder("Display Templates").ResolveSubFolder(displayTemplateCategory);

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadDisplayTemplateFile] Resolving the relative path to get the folder to be uploaded");
                    Console.WriteLine("[UploadDisplayTemplateFile] Resolving the relative path to get the folder to be uploaded");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadDisplayTemplateFile] Uploading the File " + fileName);
                    Console.WriteLine("[UploadDisplayTemplateFile] Uploading the File " + fileName);

                    FileFolderExtensions.UploadFile(folder, fileName, localFilePath, overwriteIfExists);

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadDisplayTemplateFile] Successfully Uploaded File to " + folder.ServerRelativeUrl);
                    Console.WriteLine("[UploadDisplayTemplateFile] Successfully Uploaded File to " + folder.ServerRelativeUrl);
                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UploadDisplayTemplateFile]");
                Console.WriteLine("[END][UploadDisplayTemplateFile]");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part Trasnformation Utility Execution Completed for Web - UploadDisplayTemplateFile ##############");
                Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  for Web - UploadDisplayTemplateFile ##############");
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "UploadDisplayTemplateFile", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][UploadDisplayTemplateFile] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][UploadDisplayTemplateFile] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        private void UploadWebPartFile(string webUrl, string fileName, string localFilePath, bool overwriteIfExists, string group, string outPutDirectory, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            ExceptionCsv.WebUrl = webUrl;
            string exceptionCommentsInfo1 = string.Empty;

            ClientContext clientContext = new ClientContext(webUrl);
            AuthenticationHelper ObjAuth = new AuthenticationHelper();
            Web web = clientContext.Web;
            Folder folder = null;
            string webPartXml = string.Empty;

            try
            {
                //Initialized Exception and Logger. 
                WebPart_Initialization(outPutDirectory);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started - UploadWebPartFile ##############");
                Console.WriteLine("############## Web Part  Trasnformation Utility Execution Started - UploadWebPartFile ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UploadWebPartFile] ");
                Console.WriteLine("[START][UploadWebPartFile] ");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadWebPartFile] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
                Console.WriteLine("[UploadWebPartFile] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);

                //Prepare Exception Comments
                exceptionCommentsInfo1 = "Web Url: " + webUrl + ", Web Part File Name : " + fileName + ", Group: " + group;

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.OnPremise.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UploadWebPartFile] GetNetworkCredentialAuthenticatedContext for WebUrl: " + webUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(webUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UploadWebPartFile] GetNetworkCredentialAuthenticatedContext for WebUrl: " + webUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.Online.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UploadWebPartFile] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + webUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(webUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UploadWebPartFile] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + webUrl);
                }
                if (clientContext != null)
                {
                    folder = clientContext.Web.Lists.GetByTitle("Web Part Gallery").RootFolder;
                    clientContext.Load(folder);
                    clientContext.ExecuteQuery();

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadWebPartFile] Uploading the File " + fileName);
                    Console.WriteLine("[UploadWebPartFile] Uploading the File " + fileName);

                    FileFolderExtensions.UploadFile(folder, fileName, localFilePath, overwriteIfExists);

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadWebPartFile] Successfully uploaded web part to the web part gallery " + folder);
                    Console.WriteLine("[UploadWebPartFile] Successfully uploaded web part to the web part gallery " + folder);

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadWebPartFile] Updating group of the web parts");
                    Console.WriteLine("[UploadWebPartFile] Updating group of the web parts");

                    List list = clientContext.Web.Lists.GetByTitle("Web Part Gallery");
                    CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery(100);
                    
                    Microsoft.SharePoint.Client.ListItemCollection items = list.GetItems(camlQuery);
                    clientContext.Load(items);
                    clientContext.ExecuteQuery();

                    foreach (ListItem item in items)
                    {
                        if (item["FileLeafRef"].ToString().ToLowerInvariant() == fileName.ToLowerInvariant())
                        {
                            item["Group"] = group;
                            item.Update();
                            clientContext.ExecuteQuery();
                        }
                    }

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadWebPartFile] Successfully updated the group of the web parts");
                    Console.WriteLine("[UploadWebPartFile] Successfully updated the group of the web parts");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadWebPartFile] Uploaded the Newly configured Web Part at " + folder.ServerRelativeUrl);
                    Console.WriteLine("[UploadWebPartFile] Uploaded the Newly configured Web Part at " + folder.ServerRelativeUrl);

                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UploadWebPartFile]");
                Console.WriteLine("[END][UploadWebPartFile]");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part Trasnformation Utility Execution Completed for Web - UploadWebPartFile ##############");
                Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  for Web - UploadWebPartFile ##############");
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "UploadWebPartFile", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][UploadWebPartFile] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][UploadWebPartFile] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        public void UploadDependencyFile(string webUrl, string folderServerRelativeUrl,string fileName, string localFilePath, bool overwriteIfExists, string outPutDirectory, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            ClientContext clientContext = new ClientContext(webUrl);
            AuthenticationHelper ObjAuth = new AuthenticationHelper();
            string webPartXml = string.Empty;
            Web web = clientContext.Web;
            clientContext.Load(web);
            clientContext.ExecuteQuery();

            ExceptionCsv.WebUrl = web.Url;
            string exceptionCommentsInfo1 = string.Empty;
            try
            {
                //Initialized Exception and Logger. 
                WebPart_Initialization(outPutDirectory);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started - UploadDependencyFile ##############");
                Console.WriteLine("############## Web Part  Trasnformation Utility Execution Started - UploadDependencyFile ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START]UploadDependencyFile");
                Console.WriteLine("[START]UploadDependencyFile");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadDependencyFile] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
                Console.WriteLine("[UploadDependencyFile] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);

                
                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.OnPremise.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UploadDependencyFile] GetNetworkCredentialAuthenticatedContext for WebUrl: " + web.Url);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(web.Url, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UploadDependencyFile] GetNetworkCredentialAuthenticatedContext for WebUrl: " + web.Url);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.Online.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UploadDependencyFile] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + web.Url);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(web.Url, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UploadDependencyFile] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + web.Url);
                }
                if (clientContext != null)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadDependencyFile] Successful authentication");
                    Console.WriteLine("[UploadDependencyFile] Successful authentication");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UploadDependencyFile] ::: UploadFile");
                    Console.WriteLine("[START][UploadDependencyFile]  ::: UploadFile");
                                        
                    Folder folder = clientContext.Web.RootFolder;

                    //Prepare Exception Comments
                    exceptionCommentsInfo1 = "Display Template File Name : " + fileName + " , Folder: " + folder;

                    clientContext.Load(folder, f => f.Folders, f => f.ServerRelativeUrl);
                    clientContext.ExecuteQuery();
                    while (folderServerRelativeUrl.Contains(folder.ServerRelativeUrl) && !folderServerRelativeUrl.Equals(folder.ServerRelativeUrl))
                    {
                        foreach (Folder _folder in folder.Folders)
                        {
                            if (folderServerRelativeUrl.Contains(_folder.ServerRelativeUrl))
                            {
                                folder = _folder;
                                break;
                            }
                        }
                        clientContext.Load(folder.Folders);
                        clientContext.ExecuteQuery();
                    }

                    FileFolderExtensions.UploadFile(folder, fileName, localFilePath, overwriteIfExists);

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UploadDependencyFile] UploadFile");
                    Console.WriteLine("[END][UploadDependencyFile] UploadFile ");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadDependencyFile] Successfully Upload File to " + folder.ServerRelativeUrl);
                    Console.WriteLine("[UploadDependencyFile] Successfully Upload File to " + folder.ServerRelativeUrl);

                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UploadDependencyFile]");
                Console.WriteLine("[END][UploadDependencyFile]");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part Trasnformation Utility Execution Completed for Web - UploadDependencyFile ##############");
                Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  for Web - UploadDependencyFile ##############");
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "UploadDependencyFile", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][UploadDependencyFile] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][UploadDependencyFile] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        public void ConfigureNewWebPartXml(string targetWebPartXmlFilePath, string sourceXmlFilesDirectory, string OutPutDirectory, string ActionType="")
        {
            string sourceWebPartXmlFilePath = string.Empty;
            if (ActionType.ToLower() != Constants.ActionType_All.ToLower())
            {
                //Initialized Exception and Logger. 
                WebPart_Initialization(OutPutDirectory);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started - ConfigureNewWebPartXml ##############");
                Console.WriteLine("############## Web Part  Trasnformation Utility Execution Started - ConfigureNewWebPartXml ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START]::: ConfigureNewWebPartXml");
                Console.WriteLine("[START]::: ConfigureNewWebPartXml");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ConfigureNewWebPartXml] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + OutPutDirectory);
                Console.WriteLine("[ConfigureNewWebPartXml] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + OutPutDirectory);
            }
            
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ConfigureNewWebPartXml] Configuring target webpart with source web part");
            Console.WriteLine("[ConfigureNewWebPartXml] Configuring target webpart with source web part");

            try
            {
                DirectoryInfo d = new DirectoryInfo(sourceXmlFilesDirectory);
                FileInfo[] Files = d.GetFiles("*.xml");

                foreach (FileInfo file in Files)
                {
                    sourceWebPartXmlFilePath = file.DirectoryName + @"\" + file.Name;
                    ConfigureNewWebPartXmlFile(targetWebPartXmlFilePath, sourceWebPartXmlFilePath, OutPutDirectory);
                }

                if (ActionType.ToLower() != Constants.ActionType_All.ToLower())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ConfigureNewWebPartXml]::: ConfigureNewWebPartXml");
                    Console.WriteLine("[END][ConfigureNewWebPartXml]::: ConfigureNewWebPartXml ");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part Trasnformation Utility Execution Completed for Web - ConfigureNewWebPartXml ##############");
                    Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  for Web - ConfigureNewWebPartXml ##############");
                }
                
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "ConfigureNewWebPartXml", ex.GetType().ToString());
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][ConfigureNewWebPartXml] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][ConfigureNewWebPartXml] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        public void ConfigureNewWebPartXmlFile(string targetWebPartXmlFilePath, string sourceXmlFile, string OutPutDirectory)
        {

            string targetXmlFilesDirectory = OutPutDirectory + @"\" + Constants.TARGET_WEBPART_XML_DIR;

            if (!System.IO.Directory.Exists(targetXmlFilesDirectory))
            {
                System.IO.Directory.CreateDirectory(targetXmlFilesDirectory);
            }

            string exceptionCommentsInfo1 = string.Empty;
            webParts sourceWebPart;
            webParts targetWebPart;
            bool isUpdatePoperty = false;
            string sourceWebPartXmlFilePath = string.Empty;
            StringBuilder notUpdatedPropertyInfo = new StringBuilder();

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ConfigureNewWebPartXmlFile] Configuring target webpart with source web part");
            Console.WriteLine("[ConfigureNewWebPartXmlFile] Configuring target webpart with source web part");

            try
            {

                FileInfo file = new FileInfo(sourceXmlFile);

                sourceWebPartXmlFilePath = file.DirectoryName + @"\" + file.Name;
                string[] webPartId = System.IO.Path.GetFileNameWithoutExtension(sourceWebPartXmlFilePath).Split('_');
                string newWebPartXmlFilePath = targetXmlFilesDirectory + @"\Configured_" + webPartId[0] + "_" + System.IO.Path.GetFileNameWithoutExtension(targetWebPartXmlFilePath) + ".xml";
                string notUpdatedPropertiesXmlFilePath = targetXmlFilesDirectory + @"\NotUpdated" + "_" + webPartId[0] + ".xml";

                //Prepare Exception Comments
                exceptionCommentsInfo1 = "Source Web Part File Path : " + sourceWebPartXmlFilePath + ", Target Web Part File Path: " + targetWebPartXmlFilePath;

                using (System.IO.FileStream fs = new System.IO.FileStream(sourceWebPartXmlFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                {
                    System.Xml.XmlReader reader = new XmlTextReader(fs);
                    XmlSerializer serializer = new XmlSerializer(typeof(webParts));
                    sourceWebPart = (webParts)serializer.Deserialize(reader);
                }
                using (System.IO.FileStream fs = new System.IO.FileStream(targetWebPartXmlFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                {
                    XmlReader reader = new XmlTextReader(fs);
                    XmlSerializer serializer = new XmlSerializer(typeof(webParts));
                    targetWebPart = (webParts)serializer.Deserialize(reader);
                }

                webPartsWebPart notUpdatedProperties = new webPartsWebPart();
                notUpdatedProperties.data = new webPartsWebPartData();
                notUpdatedProperties.data.properties = new webPartsWebPartDataProperty[sourceWebPart.webPart.data.properties.Length];
                for (int i = 0; i < sourceWebPart.webPart.data.properties.Length; i++)
                {
                    webPartsWebPartDataProperty customeWBProperty = sourceWebPart.webPart.data.properties[i];
                    isUpdatePoperty = false;
                    foreach (webPartsWebPartDataProperty oOTBWBProperty in targetWebPart.webPart.data.properties)
                    {
                        if (oOTBWBProperty.name.Equals(customeWBProperty.name))
                        {
                            oOTBWBProperty.Value = customeWBProperty.Value;
                            isUpdatePoperty = true;
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ConfigureNewWebPartXmlFile] Property:" + oOTBWBProperty.name + " matched in both " + System.IO.Path.GetFileNameWithoutExtension(sourceWebPartXmlFilePath) + " and " + System.IO.Path.GetFileNameWithoutExtension(targetWebPartXmlFilePath));
                            //Console.WriteLine("[ConfigureNewWebPartXmlFile] Property:" + oOTBWBProperty.name + " matched in both " + System.IO.Path.GetFileNameWithoutExtension(sourceWebPartXmlFilePath) + " and " + System.IO.Path.GetFileNameWithoutExtension(targetWebPartXmlFilePath));
                            break;
                        }
                    }
                    if (!isUpdatePoperty)
                    {
                        notUpdatedProperties.data.properties.SetValue(customeWBProperty, i);
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[ConfigureNewWebPartXmlFile] Property:" + customeWBProperty.name + "doesn't matched in both " + System.IO.Path.GetFileNameWithoutExtension(sourceWebPartXmlFilePath) + " and " + System.IO.Path.GetFileNameWithoutExtension(targetWebPartXmlFilePath));
                        //Console.WriteLine("[ConfigureNewWebPartXmlFile] Property:" + customeWBProperty.name + " doesn't matched in both " + System.IO.Path.GetFileNameWithoutExtension(sourceWebPartXmlFilePath) + " and " + System.IO.Path.GetFileNameWithoutExtension(targetWebPartXmlFilePath));

                    }
                }
                using (System.IO.StreamWriter writer = new System.IO.StreamWriter(newWebPartXmlFilePath, false))
                {
                    XmlSerializer sz = new XmlSerializer(typeof(webParts));
                    sz.Serialize(writer, targetWebPart);
                }

                StringBuilder newWebPartXmlFile = new StringBuilder();
                using (System.IO.StreamReader reader = new System.IO.StreamReader(newWebPartXmlFilePath, false))
                {

                    while (!reader.EndOfStream)
                    {
                        string currentLine = reader.ReadLine();
                        if (currentLine.Contains("www.w3.org"))
                        {
                            string[] currentLineArray = currentLine.Split(' ');
                            if (currentLineArray.Count() > 2)
                            {
                                currentLine = currentLineArray[0] + ">";
                            }
                        }
                        newWebPartXmlFile.AppendLine(currentLine);
                    }
                }
                using (System.IO.StreamWriter writer = new System.IO.StreamWriter(newWebPartXmlFilePath, false))
                {
                    writer.WriteLine(newWebPartXmlFile.ToString());
                }
                using (System.IO.StreamWriter writer = new System.IO.StreamWriter(notUpdatedPropertiesXmlFilePath, false))
                {
                    XmlSerializer sz = new XmlSerializer(typeof(webPartsWebPart));
                    sz.Serialize(writer, notUpdatedProperties);
                }
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ConfigureNewWebPartXmlFile] New Configured web part is created at " + newWebPartXmlFilePath);
                Console.WriteLine("[ConfigureNewWebPartXmlFile] New Configured web part is created at " + newWebPartXmlFilePath);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ConfigureNewWebPartXmlFile] The properties which are not configured are extracted into a new file at " + notUpdatedPropertiesXmlFilePath);
                Console.WriteLine("[ConfigureNewWebPartXmlFile] The properties which are not configured are extracted into a new file at " + notUpdatedPropertiesXmlFilePath);

                string result = newWebPartXmlFilePath + ";" + notUpdatedPropertiesXmlFilePath;
                string[] filePaths = result.Split(';');
                if (filePaths.Count() > 1)
                {
                    if (!String.IsNullOrEmpty(filePaths[0].Trim()))
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[ConfigureNewWebPartXmlFile] New Configured Web Part Xml File:" + filePaths[0]);
                        Console.WriteLine("[ConfigureNewWebPartXmlFile] New Configured Web Part Xml File:" + filePaths[0]);

                    }
                    if (!String.IsNullOrEmpty(filePaths[1].Trim()))
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[ConfigureNewWebPartXmlFile] Custom Properties not configured in the new web part xml file:" + filePaths[1]);
                        Console.WriteLine("[ConfigureNewWebPartXmlFile] Custom Properties not configured in the new web part xml file:" + filePaths[1]);

                    }
                }

                Console.ForegroundColor = ConsoleColor.Gray;
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ConfigureNewWebPartXmlFile] Successfully configured web part");
                Console.WriteLine("[ConfigureNewWebPartXmlFile] Successfully configured web part");


            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "ConfigureNewWebPartXmlFile", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][ConfigureNewWebPartXmlFile] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][ConfigureNewWebPartXmlFile] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        public void UploadAppInAppCatalog(string appCatalogUrl, string fileName, string appFilePath, string outPutDirectory, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            ClientContext clientContext = new ClientContext(appCatalogUrl);
            AuthenticationHelper ObjAuth = new AuthenticationHelper();
            string webPartXml = string.Empty;
            ExceptionCsv.WebUrl = appCatalogUrl;
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                //Initialized Exception and Logger. 
                WebPart_Initialization(outPutDirectory);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started - UploadAppInAppCatalog ##############");
                Console.WriteLine("############## Web Part  Trasnformation Utility Execution Started - UploadAppInAppCatalog ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: UploadAppInAppCatalog");
                Console.WriteLine("[START] ::: UploadAppInAppCatalog");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadAppInAppCatalog] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
                Console.WriteLine("[UploadAppInAppCatalog] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);

                //Prepare Exception Comments
                exceptionCommentsInfo1 = "App File Name : " + fileName + ", App Catalogue: " + appCatalogUrl;

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.OnPremise.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UploadAppInAppCatalog] GetNetworkCredentialAuthenticatedContext for WebUrl: " + appCatalogUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(appCatalogUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UploadAppInAppCatalog] GetNetworkCredentialAuthenticatedContext for WebUrl: " + appCatalogUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.Online.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UploadDisplayTemplateFile] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + appCatalogUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(appCatalogUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UploadDisplayTemplateFile] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + appCatalogUrl);
                }

                if (clientContext != null)
                {
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    clientContext.ExecuteQuery();

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadAppInAppCatalog] Successful authentication");
                    Console.WriteLine("[UploadAppInAppCatalog] Successful authentication");

                    ListCollection listCollection = web.Lists;
                    clientContext.Load(listCollection);
                    clientContext.ExecuteQuery();
                    if (IsLibraryExist("Apps for SharePoint", listCollection))
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadAppInAppCatalog] App Catalogue: " + appCatalogUrl + " contains list: Apps for SharePoint");
                        Console.WriteLine("[UploadAppInAppCatalog] App Catalogue: " + appCatalogUrl + " contains list: Apps for SharePoint");

                        using (var fs = new FileStream(appFilePath, FileMode.Open))
                        {
                            //Uploading the App to the App Catalogue
                            var fi = new FileInfo(appFilePath);
                            var list = clientContext.Web.Lists.GetByTitle("Apps for SharePoint");
                            clientContext.Load(list.RootFolder);
                            clientContext.ExecuteQuery();

                            Folder folder = list.RootFolder;
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UploadAppInAppCatalog] UploadFile");
                            Console.WriteLine("[START][UploadAppInAppCatalog] UploadFile");

                            Microsoft.SharePoint.Client.File file = FileFolderExtensions.UploadFile(folder, fileName, appFilePath, true);

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UploadAppInAppCatalog] UploadFile");
                            Console.WriteLine("[END][UploadAppInAppCatalog] UploadFile ");

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[UploadAppInAppCatalog] Successfully Upload File to " + appCatalogUrl);
                            Console.WriteLine("[UploadAppInAppCatalog] Successfully Upload File to " + appCatalogUrl);
                        }
                    }
                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part Trasnformation Utility Execution Completed for Web - UploadAppInAppCatalog ##############");
                Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  for Web - UploadAppInAppCatalog ##############");

            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "UploadAppInAppCatalog", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][UploadAppInAppCatalog] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][UploadAppInAppCatalog] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        private AppInstance DeployApp(ClientContext context, string appFullPath)
        {
            using (var packageStream = System.IO.File.OpenRead(appFullPath))
            {
                var appInstance = context.Web.LoadAndInstallApp(packageStream);
                context.Load(appInstance,
                                     a => a.AppWebFullUrl,
                                     a => a.Status);
                context.ExecuteQuery();

                Console.WriteLine(appInstance.AppWebFullUrl);
                return appInstance;
            }
        }

        public void DeleteWebPart_UsingCSV(string sourceWebPartType, string usageFileName, string outPutFolder, string csvFile, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                WebPart_Initialization(outPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started - DeleteWebPart_UsingCSV - For InputCSV ##############");
                Console.WriteLine("############## Web Part Trasnformation Utility Execution Started - DeleteWebPart_UsingCSV - For InputCSV ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DeleteWebPart_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);
                Console.WriteLine("[DeleteWebPart_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);

                //Reading Input File
                IEnumerable<WebPartDiscoveryInput> objWPDInput;
                ReadWebPartUsageCSV(sourceWebPartType, usageFileName, outPutFolder, out objWPDInput, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                bool headerTransformWebPart = false;

                if (objWPDInput.Any())
                {
                    for (int i = 0; i < objWPDInput.Count(); i++)
                    {
                        WebPartDiscoveryInput objInput = objWPDInput.ElementAt(i);
                        TranformWebPartStatusBase objWPOutputBase = new TranformWebPartStatusBase();

                        try
                        {
                            SetExceptionCSVWebAppSiteColWebUrls(objInput.WebUrl.ToString(), SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                            //This is for Exception Comments:
                            exceptionCommentsInfo1 = "WebUrl: " + objInput.WebUrl + ", ZoneID: " + objInput.ZoneID + ", StorageKey" + objInput.StorageKey;
                            //This is for Exception Comments:

                            bool status = DeleteWebPart(objInput.WebUrl, objInput.PageUrl.ToString(), objInput.StorageKey, outPutFolder, SharePointOnline_OR_OnPremise, UserName, Password, Domain, Constants.ActionType_CSV.ToLower());

                            if (status)
                            {
                                objWPOutputBase.Status = "Successfully Deleted WebPart";
                            }
                            else
                            {
                                objWPOutputBase.Status = "Failed to Deleted WebPart";
                            }
                        }
                        catch (Exception ex)
                        {
                            objWPOutputBase.Status = "Failed to Deleted WebPart";
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Error in Processing Web:" + ExceptionCsv.WebUrl + " `\r\nError Details:" + ex.Message + " `\r\nExceptionComments:" + exceptionCommentsInfo1);
                            Console.ForegroundColor = ConsoleColor.Gray;
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception][DeleteWebPart_UsingCSV] Exception Message: " + ex.Message);
                            ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "DeleteWebPart_UsingCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                        }

                        objWPOutputBase.WebPartType = objInput.WebPartType;
                        objWPOutputBase.ZoneID = objInput.ZoneID;
                        objWPOutputBase.ZoneIndex = objInput.ZoneIndex;
                        objWPOutputBase.WebPartId = objInput.WebPartId;
                        objWPOutputBase.PageUrl = objInput.PageUrl;
                        objWPOutputBase.WebUrl = ExceptionCsv.WebUrl;
                        objWPOutputBase.SiteCollection = ExceptionCsv.SiteCollection;
                        objWPOutputBase.WebApplication = ExceptionCsv.WebApplication;

                        if (System.IO.File.Exists(csvFile))
                        {
                            headerTransformWebPart = true;
                        }
                        FileUtility.WriteCsVintoFile(csvFile, objWPOutputBase, ref headerTransformWebPart);
                    }
                }
                Logger.AddMessageToTraceLogFile(Constants.Logging, "##############  Web Part  Trasnformation Utility Execution Completed  - DeleteWebPart_UsingCSV - InputCSV ##############");
                Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  - DeleteWebPart_UsingCSV - InputCSV ##############");
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] FUNCTION DeleteWebPart_UsingCSV. Exception Message: " + ex.Message);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "DeleteWebPart_UsingCSV", ex.GetType().ToString(), exceptionCommentsInfo1);

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception] FUNCTION DeleteWebPart_UsingCSV. Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        public bool DeleteWebPart(string webUrl, string pageUrl, string _storageKey, string outPutDirectory, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A", string ActionType = "web")
        {
            ClientContext clientContext = new ClientContext(webUrl);
            AuthenticationHelper ObjAuth = new AuthenticationHelper();
            bool isWebPartDeleted = false;
            string webPartXml = string.Empty;
            ExceptionCsv.WebUrl = webUrl;
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                //This function is Get Relative URL of the page
                string ServerRelativePageUrl = string.Empty;
                ServerRelativePageUrl = GetPageRelativeURL(webUrl, pageUrl, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                Guid storageKey = new Guid(GetWebPartID(_storageKey));

                if (ActionType.ToLower().Trim() == Constants.ActionType_Web.ToLower())
                {
                    //Initialized Exception and Logger. 
                    WebPart_Initialization(outPutDirectory);

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part Trasnformation Utility Execution Started - DeleteWebPart ##############");
                    Console.WriteLine("############## Web Part  Trasnformation Utility Execution Started - DeleteWebPart ##############");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                    Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[DeleteWebPart] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
                    Console.WriteLine("[DeleteWebPart] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
                }
                //Prepare Exception Comments
                exceptionCommentsInfo1 = "Web Url: " + webUrl + ", Storage Key : " + storageKey.ToString() + ", Page Url: " + ServerRelativePageUrl;

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.OnPremise.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][DeleteWebPart] GetNetworkCredentialAuthenticatedContext for WebUrl: " + webUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(webUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][DeleteWebPart] GetNetworkCredentialAuthenticatedContext for WebUrl: " + webUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.Online.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][DeleteWebPart] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + webUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(webUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][DeleteWebPart] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + webUrl);
                }

                if (clientContext != null)
                {
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    clientContext.ExecuteQuery();

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[DeleteWebPart] Successful authentication");
                    Console.WriteLine("[DeleteWebPart] Successful authentication");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][DeleteWebPart] Checking Out File ...");
                    Console.WriteLine("[START][DeleteWebPart] Checking Out File ...");


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
                            forceCheckOut = list.ForceCheckout;
                            enableVersioning = list.EnableVersioning;
                            enableMinorVersions = list.EnableMinorVersions;
                            enableModeration = list.EnableModeration;
                            dVisibility = list.DraftVersionVisibility;

                            Logger.AddMessageToTraceLogFile(Constants.Logging,
                                "[DeleteWebpart] List Details " + ServerRelativePageUrl + ". " +
                                "Force Check Out: " + forceCheckOut +
                                "Enable Versioning: " + enableVersioning +
                                "Enable Minor Versions: " + enableMinorVersions +
                                "Enable Moderation: " + enableModeration +
                                "Draft Version Visibility: " + dVisibility);

                            Logger.AddMessageToTraceLogFile(Constants.Logging,
                                "[DeleteWebpart] Removing Versioning");
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
                                clientContext.ExecuteQuery();
                            }
                            #endregion
                        }

                        catch (Exception ex)
                        {
                            ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "DeleteWebPart", ex.GetType().ToString(), exceptionCommentsInfo1);
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][DeleteWebPart] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);

                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("[EXCEPTION][DeleteWebPart] Exception Message: " + ex.Message);
                            Console.ForegroundColor = ConsoleColor.Gray;
                        
                        }
                    }

                    

                    try
                    {

                        if (DeleteWebPart(clientContext.Web, ServerRelativePageUrl, storageKey))
                        {
                            isWebPartDeleted = true;
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[DeleteWebPart] Successfully Deleted the WebPart");
                            Console.WriteLine("[DeleteWebPart] Successfully Deleted the WebPart");
                        }
                        else
                        {
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[DeleteWebPart] WebPart with StorageKey: " + storageKey + " does not exist in the Page: " + ServerRelativePageUrl);
                            Console.WriteLine("[DeleteWebPart] WebPart with StorageKey: " + storageKey + " does not exist in the Page: " + ServerRelativePageUrl);
                        }
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
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][DeleteWebPart]  File Checked in after successfully deleting the webpart.");
                    Console.WriteLine("[END][DeleteWebPart]  File Checked in after successfully deleting the webpart.");

                }
                
                if (ActionType.ToLower().Trim() ==  Constants.ActionType_Web.ToLower())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part Trasnformation Utility Execution Completed for Web - DeleteWebPart ##############");
                    Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  for Web - DeleteWebPart ##############");
                }

            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "DeleteWebPart", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][DeleteWebPart] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][DeleteWebPart] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            return isWebPartDeleted;
        }

        private bool DeleteWebPart(Web web, string serverRelativePageUrl, Guid storageKey)
        {
            bool isWebPartDeleted = false;
            var webPartPage = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

            web.Context.Load(webPartPage);
            web.Context.ExecuteQueryRetry();

            LimitedWebPartManager limitedWebPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            web.Context.Load(limitedWebPartManager.WebParts, wps => wps.Include(wp => wp.Id));
            web.Context.ExecuteQueryRetry();
            
            if (limitedWebPartManager.WebParts.Count >= 0)
            {
                foreach (WebPartDefinition webpartDef in limitedWebPartManager.WebParts)
                {
                    Microsoft.SharePoint.Client.WebParts.WebPart oWebPart = webpartDef.WebPart;
                    if (webpartDef.Id.Equals(storageKey))
                    {                        
                        webpartDef.DeleteWebPart();
                        web.Context.ExecuteQueryRetry();
                        isWebPartDeleted = true;
                        break;
                    }
                }
            }
            return isWebPartDeleted;
        }

        //sourceWebPartId - Used to update the content of the wikipage with new web part id [Fix for BugId - 95007] 
        private bool AddWebPart(Web web, string serverRelativePageUrl, WebPartEntity webPartEntity,string sourceWebPartId="")
        {
            bool isWebPartAdded = false;
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

            return isWebPartAdded;
        }

        //Used to update the content of the wikipage/PublishingPageContent with new web part id [Fix for BugId - 95007] 
        private void CheckForWikiFieldOrPublishingPageContentAndUpdate(WebPartDefinition webPart, Web web, Microsoft.SharePoint.Client.File webPartPage, string sourceWebPartId = "")
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
        
        private List GetPageList(ref ClientContext clientContext)
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
                    Log.Info("GetPageList", "Web: " + web.Url + "is a publishing web");
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
                        list = web.Lists.GetByTitle(Constants.TEAMSITE_PAGES_LIBRARY);

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
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "GetPageList", ex.GetType().ToString());
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][GetPageList] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][GetPageList] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            return list;
        }

        private bool IsPublishingWeb(ClientContext clientContext, Web web)
        {
            Logger.AddMessageToTraceLogFile(Constants.Logging, "Checking if the current web is a publishing web");
            Console.WriteLine("Checking if the current web is a publishing web");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "Checking for PublishingFeatureActivated ...");
            Console.WriteLine("Checking for PublishingFeatureActivated ...");
            
            var _IsPublished = false;
            var propName = "__PublishingFeatureActivated";

            try
            {

                //Ensure web properties are loaded
                if (!web.IsObjectPropertyInstantiated("AllProperties"))
                {
                    clientContext.Load(web, w => w.AllProperties);
                    clientContext.ExecuteQuery();
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "Ensure web properties are loaded...");
                    
                }
                //Verify whether publishing feature is activated 
                if (web.AllProperties.FieldValues.ContainsKey(propName))
                {
                    bool propVal;
                    Boolean.TryParse((string)web.AllProperties[propName], out propVal);

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "Verify whether publishing feature is activated...");
                    
                    _IsPublished = propVal;
                    return propVal;
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "IsPublishingWeb", ex.GetType().ToString());
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][IsPublishingWeb] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][IsPublishingWeb] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            return _IsPublished;
        }

        private bool IsLibraryExist(string pageLibraryName, ListCollection collList)
        {
            foreach (List oList in collList)
            {
                if (oList.Title.Equals(pageLibraryName))
                {
                    return true;
                }
            }
            return false;
        }

        private string GetPageRelativeURL(string WebUrl, string PageUrl, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            AuthenticationHelper ObjAuth = new AuthenticationHelper();
            ClientContext clientContext = null;
            string _relativePageUrl = string.Empty;

            if (WebUrl != "" || PageUrl != "")
            {
                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.OnPremise)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][GetPageRelativeURL] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(WebUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][GetPageRelativeURL] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.Online)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][GetPageRelativeURL] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(WebUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][GetPageRelativeURL] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                }
                if (clientContext != null)
                {
                    Web _Web = clientContext.Web;
                    clientContext.Load(_Web);
                    clientContext.ExecuteQuery();

                    Logger.AddMessageToTraceLogFile(true, "[GetPageRelativeURL] Web.ServerRelativeUrl: " + _Web.ServerRelativeUrl + " And PageUrl: " + PageUrl);

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
                    Logger.AddMessageToTraceLogFile(true, "[GetPageRelativeURL] RelativePageUrl Framed: " + _relativePageUrl);
                }
            }

            return _relativePageUrl;
        }

        private string GetWebPartID(string webPartID)
        {
            string _webPartID = string.Empty;

            string[] tempStr = webPartID.Split('_');

            if (tempStr.Length>5)
            {
                _webPartID = webPartID.Remove(0,tempStr[0].Length+1).Replace('_', '-');
            }
            else
            {
                _webPartID = webPartID.Replace('_', '-');
            }
           
            return _webPartID;
        }

        private string GetWebPartShortTypeName(string webPartType)
        {
            string _webPartType = string.Empty;

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
            return _webPartType;
        }

        private bool IsFeatureOnWeb(Guid FeatureID, ClientContext clientContext)
        {
            bool isFeatureAvailable = false;
           
            FeatureCollection features = clientContext.Web.Features;
            clientContext.Load(features);
            clientContext.ExecuteQuery();

            Feature feature = features.GetById(FeatureID);
            if (feature!=null)
            {
                clientContext.Load(feature);
                clientContext.ExecuteQuery();
                if (feature.DefinitionId != null)
                {
                    isFeatureAvailable = true;
                }                    
            }           

            return isFeatureAvailable;
        }

        private bool CheckWebPartOrAppPartPresenceInSite(ClientContext clientContext, string targetWebPartXmlFileName, string targetWebPartXmlFilePath)
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

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[CheckWebPartOrAppPartPresenceInSite] Checking for web part schema version");
                Console.WriteLine("[CheckWebPartOrAppPartPresenceInSite] Checking for web part schema version");

                if (webPartPropertiesXml.Contains("WebPart/v2"))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[CheckWebPartOrAppPartPresenceInSite] Web part schema version is V2");
                    Console.WriteLine("[CheckWebPartOrAppPartPresenceInSite] Web part schema version is V2");

                    XmlDataDocument xmldoc = new XmlDataDocument();
                    xmldoc.LoadXml(webPartPropertiesXml);
                    webPartType = GetWebPartShortTypeName(xmldoc.DocumentElement.GetElementsByTagName("TypeName").Item(0).FirstChild.Value);

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[CheckWebPartOrAppPartPresenceInSite] Web part Type: " + webPartType);
                    Console.WriteLine("[CheckWebPartOrAppPartPresenceInSite] Web part Type: " + webPartType);

                    xmldoc = null;
                }
                else
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[CheckWebPartOrAppPartPresenceInSite] Web part schema version is V3");
                    Console.WriteLine("[CheckWebPartOrAppPartPresenceInSite] Web part schema version is V3");

                    using (System.IO.FileStream fs = new System.IO.FileStream(targetWebPartXmlFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                    {
                        XmlReader reader = new XmlTextReader(fs);
                        XmlSerializer serializer = new XmlSerializer(typeof(webParts));
                        targetWebPart = (webParts)serializer.Deserialize(reader);
                        if (targetWebPart != null)
                        {
                            webPartType = GetWebPartShortTypeName(targetWebPart.webPart.metaData.type.name);

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[CheckWebPartOrAppPartPresenceInSite] Web part Type: " + webPartType);
                            Console.WriteLine("[CheckWebPartOrAppPartPresenceInSite] Web part Type: " + webPartType);
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
                            isWebPartInSite = IsFeatureOnWeb(featureID,clientContext);
                            break;
                        }
                    }

                }
                else
                {
                    Web web = clientContext.Site.RootWeb;
                    clientContext.Load(web, w => w.Url);
                    clientContext.ExecuteQuery();

                    List list = web.Lists.GetByTitle("Web Part Gallery");
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
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "CheckWebPartOrAppPartPresenceInSite", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][CheckWebPartOrAppPartPresenceInSite] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][CheckWebPartOrAppPartPresenceInSite] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
                return isWebPartInSite;
            }

            return isWebPartInSite;
        }
        private string GetWebPartPropertiesServiceCall(ClientContext clientContext, string storageKey, string pageUrl)
        {
            string webPartXml = string.Empty;
            var service = new Transformation.PowerShell.WebPartPagesService.WebPartPagesWebService();
            service.Url = clientContext.Web.Url + Constants.WEBPART_SERVICE;

            service.PreAuthenticate = true;

            service.Credentials = clientContext.Credentials;

            // Actual web service call which returns the information in string format
            webPartXml = service.GetWebPart2(pageUrl, storageKey.ToGuid(), Storage.Shared, SPWebServiceBehavior.Version3);

            return webPartXml;
        }

        private string GetTargetWebPartXmlFilePath(string webPartId, string targetDirectory)
        {
            string targetWebPartXmlFile = string.Empty;
            DirectoryInfo d = new DirectoryInfo(targetDirectory);
            FileInfo[] Files = d.GetFiles("*.xml");
            foreach (FileInfo file in Files)
            {
                if (file.Name.Contains(webPartId))
                {
                    targetWebPartXmlFile = file.FullName;
                    break;
                }
            }
            return targetWebPartXmlFile;
        }

        //sourceWebPartId - Used to update the content of the wikipage with new web part id [Fix for BugId - 95007] 
        public bool AddWebPart(string webUrl, string configuredWebPartFileName, string configuredWebPartXmlFilePath, string webPartZoneIndex, string webPartZoneID, string serverRelativePageUrl, string outPutDirectory, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A", string ActionType = "web", string sourceWebPartId = "")
        {
            bool isWebPartAdded = false;

            WebPartEntity webPart = new WebPartEntity();
            webPart.WebPartIndex = Convert.ToInt32(webPartZoneIndex);
            ClientContext clientContext = new ClientContext(webUrl);
            AuthenticationHelper ObjAuth = new AuthenticationHelper();
            string webPartXml = string.Empty;

            ExceptionCsv.WebUrl = webUrl;
            string exceptionCommentsInfo1 = string.Empty;
            try
            {
                if (ActionType.ToLower() == Constants.ActionType_Web.ToLower())
                {
                    //Initialized Exception and Logger. 
                    WebPart_Initialization(outPutDirectory);

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part Trasnformation Utility Execution Started - AddWebPart ##############");
                    Console.WriteLine("############## Web Part Trasnformation Utility Execution Started - AddWebPart ##############");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                    Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddWebPart] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
                    Console.WriteLine("[AddWebPart] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
                }

                //Prepare Exception Comments
                exceptionCommentsInfo1 = "Web Url: " + webUrl + ", Configured Web Part File Name: " + configuredWebPartFileName + " , Page Url: " + serverRelativePageUrl;

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.OnPremise.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][AddWebPart] GetNetworkCredentialAuthenticatedContext for WebUrl: " + webUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(webUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][AddWebPart] GetNetworkCredentialAuthenticatedContext for WebUrl: " + webUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper().Equals(Constants.Online.ToUpper()))
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][AddWebPart] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + webUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(webUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][AddWebPart] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + webUrl);
                }

                if (clientContext != null)
                {
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    clientContext.ExecuteQuery();

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddWebPart] Successful authentication");
                    Console.WriteLine("[AddWebPart] Successful authentication");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddWebPart] Checking for web part in the Web Part Gallery");
                    Console.WriteLine("[AddWebPart] Checking for web part in the Web Part Gallery");

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

                                Logger.AddMessageToTraceLogFile(Constants.Logging,
                                    "[AddWebPart] List Details " + serverRelativePageUrl + ". " +
                                    "Force Check Out: " + forceCheckOut +
                                    "Enable Versioning: " + enableVersioning +
                                    "Enable Minor Versions: " + enableMinorVersions +
                                    "Enable Moderation: " + enableModeration +
                                    "Draft Version Visibility: " + dVisibility);

                                Logger.AddMessageToTraceLogFile(Constants.Logging,
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
                                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "AddWebPart", ex.GetType().ToString(), exceptionCommentsInfo1);
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][AddWebPart] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);

                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("[EXCEPTION][AddWebPart] Exception Message: " + ex.Message);
                                Console.ForegroundColor = ConsoleColor.Gray;

                            }
                        }
                        

                        try
                        {

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddWebPart] Adding web part to the page at " + serverRelativePageUrl);
                            Console.WriteLine("[AddWebPart] Adding web part to the page at " + serverRelativePageUrl);
                            isWebPartAdded = AddWebPart(clientContext.Web, serverRelativePageUrl, webPart, sourceWebPartId);
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
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddWebPart] Successfully Added the newly configured WebPart");

                        Console.WriteLine("[AddWebPart] Successfully Added the newly configured WebPart");
                    }
                    else
                    {
                        throw new Exception("Target Webpart should be present in the site for the webpart to be added");
                    }
                }

                if (ActionType.ToLower() == Constants.ActionType_Web.ToLower())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part Trasnformation Utility Execution Completed for Web - AddWebPart ##############");
                    Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  for Web - AddWebPart ##############");
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "AddWebPart", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][AddWebPart] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][AddWebPart] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
                return isWebPartAdded;
            }
            return isWebPartAdded;
        }

        public void AddWebPart_UsingCSV(string sourceWebPartType, string targetWebPartFileName, string targetWebPartXmlDir, string usageFileName, string outPutFolder, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                WebPart_Initialization(outPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started - AddWebPart - For InputCSV ##############");
                Console.WriteLine("############## Web Part Trasnformation Utility Execution Started - AddWebPart - For InputCSV ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());
          
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddWebPart_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);
                Console.WriteLine("[AddWebPart_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);

                //Reading Input File
                IEnumerable<WebPartDiscoveryInput> objWPDInput;
                ReadWebPartUsageCSV(sourceWebPartType, usageFileName, outPutFolder, out objWPDInput, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                bool headerTransformWebPart = false;

                bool status = false;

                if (objWPDInput.Any())
                {
                    for (int i = 0; i < objWPDInput.Count(); i++)
                    {
                        try
                        {
                            WebPartDiscoveryInput objInput = objWPDInput.ElementAt(i);
                            SetExceptionCSVWebAppSiteColWebUrls(objInput.WebUrl.ToString(), SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                            //This is for Exception Comments:
                            exceptionCommentsInfo1 = "WebUrl: " + objInput.WebUrl + ", ZoneID: " + objInput.ZoneID;
                            //This is for Exception Comments:

                            //This function is Get Relative URL of the page
                            string _relativePageUrl = string.Empty;
                            _relativePageUrl = GetPageRelativeURL(objInput.WebUrl.ToString(), objInput.PageUrl.ToString(), SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                            string _storageKey = string.Empty;
                            _storageKey = GetWebPartID(objInput.StorageKey);

                            //sourceWebPartId - Used to update the content of the wikipage with new web part id [Fix for BugId - 95007] 
                            string _webPartId = string.Empty;
                            _webPartId = GetWebPartID(objInput.WebPartId);
                            //End

                            string _targetWebPartXml = string.Empty;
                            _targetWebPartXml = GetTargetWebPartXmlFilePath(_storageKey, targetWebPartXmlDir);

                            if (_targetWebPartXml != "")
                            {
                                status = AddWebPart(objInput.WebUrl, targetWebPartFileName, _targetWebPartXml, objInput.ZoneIndex, objInput.ZoneID, _relativePageUrl, outPutFolder, SharePointOnline_OR_OnPremise, UserName, Password, Domain, Constants.ActionType_CSV.ToLower(), _webPartId);
                            }
                            else
                            {
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[AddWebPart_UsingCSV] Target WebPartXml File: " + targetWebPartFileName + " does not exists for StorageKey" + _storageKey + "in Path(TargetWebPartXmlDir): " + targetWebPartXmlDir + " for page " + _relativePageUrl);
                                Console.WriteLine("[AddWebPart_UsingCSV] Target WebPartXml File: " + targetWebPartFileName + " does not exists for StorageKey" + _storageKey + "in Path(TargetWebPartXmlDir): " + targetWebPartXmlDir + " for page " + _relativePageUrl);
                            }
                            TranformWebPartStatusBase objWPOutputBase = new TranformWebPartStatusBase();
                            objWPOutputBase.WebUrl = objInput.WebUrl;
                            objWPOutputBase.WebPartType = objInput.WebPartType;
                            objWPOutputBase.ZoneID = objInput.ZoneID;
                            objWPOutputBase.ZoneIndex = objInput.ZoneIndex;
                            objWPOutputBase.WebPartId = objInput.WebPartId;
                            objWPOutputBase.PageUrl = objInput.PageUrl;

                            if (status)
                            {
                                objWPOutputBase.Status = "Successfully Added WebPart";
                            }
                            else
                            {
                                objWPOutputBase.Status = "Failed to Add WebPart";
                            }

                            FileUtility.WriteCsVintoFile(outPutFolder + @"\" + System.IO.Path.GetFileNameWithoutExtension(usageFileName) + "_AddOperationStatus.csv", objWPOutputBase, ref headerTransformWebPart);

                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Error in Processing Web:" + ExceptionCsv.WebUrl + " `\r\nError Details:" + ex.Message + " `\r\nExceptionComments:" + exceptionCommentsInfo1);
                            Console.ForegroundColor = ConsoleColor.Gray;
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception][ReplaceWebPart_UsingCSV] Exception Message: " + ex.Message);
                            ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "ReplaceWebPart_UsingCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                        }
                    }
                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "##############  Web Part  Trasnformation Utility Execution Completed  - AddWebPart - InputCSV ##############");
                Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  - AddWebPart - InputCSV ##############");
            }
            catch (Exception ex)
            {                
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] [AddWebPart_UsingCSV] Exception Message: " + ex.Message);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "AddWebPart_UsingCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception][AddWebPart_UsingCSV] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;

            }
        }

        //sourceWebPartId - Used to update the content of the wikipage with new web part id [Fix for BugId - 95007] 
        public bool ReplaceWebPart(string webUrl, string targetWebPartFileName, string targetWebPartXmlFile, string sourceWebPartStorageKey, string webPartZoneIndex, string webPartZoneID, string serverRelativePageUrl, string outPutDirectory, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A", string ActionType = "", string sourceWebPartId = "")
        {
            bool isWebPartReplaced = false;

            if (ActionType.ToLower() == Constants.ActionType_Web.ToLower())
            {
                //Initialized Exception and Logger. 
                WebPart_Initialization(outPutDirectory);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started - ReplaceWebPart ##############");
                Console.WriteLine("############## Web Part  Trasnformation Utility Execution Started ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceWebPart] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
                Console.WriteLine("[ReplaceWebPart] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + outPutDirectory);
            }

            if (DeleteWebPart(webUrl, serverRelativePageUrl, sourceWebPartStorageKey, outPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain, Constants.ActionType_All.ToLower()))
            {
                AddWebPart(webUrl, targetWebPartFileName, targetWebPartXmlFile, webPartZoneIndex, webPartZoneID, serverRelativePageUrl, outPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain, Constants.ActionType_All.ToLower(),sourceWebPartId);
                
                isWebPartReplaced = true;
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceWebPart] Successfully Replaced the newly configured WebPart");
                Console.WriteLine("[ReplaceWebPart] Successfully Replaced the newly configured WebPart");
            }

            if (ActionType.ToLower() == Constants.ActionType_Web.ToLower())
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part Trasnformation Utility Execution Completed for Web - ReplaceWebPart ##############");
                Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  for Web - ReplaceWebPart ##############");
            }

            return isWebPartReplaced;
        }

        public void ReplaceWebPart_UsingCSV(string sourceWebPartType, string targetWebPartXmlFilePath, string targetWebPartFileName, string targetWebPartXmlDir, string usageFileName, string outPutFolder, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A", string ActionType = "")
        {
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                if (ActionType.ToLower() != Constants.ActionType_All.ToLower())
                {
                    WebPart_Initialization(outPutFolder);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started - ReplaceWebPart_UsingCSV - For InputCSV ##############");
                    Console.WriteLine("############## Web Part Trasnformation Utility Execution Started - ReplaceWebPart_UsingCSV - For InputCSV ##############");

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                    Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceWebPart_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);
                    Console.WriteLine("[ReplaceWebPart_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);
                }

                //Reading Input File
                IEnumerable<WebPartDiscoveryInput> objWPDInput;
                ReadWebPartUsageCSV(sourceWebPartType, usageFileName, outPutFolder, out objWPDInput, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                bool headerTransformWebPart = false;

                if (objWPDInput.Any())
                {
                    //bool headerPageLayout = false;

                    for (int i = 0; i < objWPDInput.Count(); i++)
                    {
                        try
                        {
                            WebPartDiscoveryInput objInput = objWPDInput.ElementAt(i);
                            SetExceptionCSVWebAppSiteColWebUrls(objInput.WebUrl.ToString(), SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                            //This is for Exception Comments:
                            exceptionCommentsInfo1 = "WebUrl: " + objInput.WebUrl + ", ZoneID: " + objInput.ZoneID + ", Web PartID:" + objInput.WebPartId.ToString() + " ,PageUrl: " + objInput.PageUrl.ToString();
                            //This is for Exception Comments:

                            //This function is Get Relative URL of the page
                            string _relativePageUrl = string.Empty;
                            _relativePageUrl = GetPageRelativeURL(objInput.WebUrl.ToString(), objInput.PageUrl.ToString(), SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                            string _storageKey = string.Empty;
                            _storageKey = GetWebPartID(objInput.StorageKey);

                            //sourceWebPartId - Used to update the content of the wikipage with new web part id [Fix for BugId - 95007] 
                            string _webPartId = string.Empty;
                            _webPartId = GetWebPartID(objInput.WebPartId);
                            //End
                            bool status = false;

                            string sourceWebPartXmlFilePath = GetWebPartProperties(_relativePageUrl, _storageKey, objInput.WebUrl, outPutFolder, SharePointOnline_OR_OnPremise, UserName, Password, Domain, "CSVUpdates");

                            if (!string.IsNullOrEmpty(sourceWebPartXmlFilePath))
                            {
                                ConfigureNewWebPartXmlFile(targetWebPartXmlFilePath, sourceWebPartXmlFilePath, outPutFolder);

                                string _targetWebPartXml = string.Empty;
                                _targetWebPartXml = GetTargetWebPartXmlFilePath(_storageKey, targetWebPartXmlDir);

                                if (_targetWebPartXml != "")
                                {
                                    status = ReplaceWebPart(objInput.WebUrl, targetWebPartFileName, _targetWebPartXml, _storageKey, objInput.ZoneIndex, objInput.ZoneID, _relativePageUrl, outPutFolder, SharePointOnline_OR_OnPremise, UserName, Password, Domain, string.Empty, _webPartId);
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceWebPart_UsingCSV] Target WebPartXml File: " + targetWebPartFileName + " does not exists for StorageKey" + _storageKey + "in Path(TargetWebPartXmlDir): " + targetWebPartXmlDir + " for page " + _relativePageUrl);
                                    Console.WriteLine("[ReplaceWebPart_UsingCSV] Target WebPartXml File: " + targetWebPartFileName + " does not exists for StorageKey" + _storageKey + "in Path(TargetWebPartXmlDir): " + targetWebPartXmlDir + " for page " + _relativePageUrl);
                                    Console.ForegroundColor = ConsoleColor.Gray;
                                }
                            }
                            else
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReplaceWebPart_UsingCSV] Failed to get Source WebPart Properties: on Web: " + objInput.WebUrl + " for StorageKey " + _storageKey + "for page " + _relativePageUrl);
                                Console.WriteLine("\r\nFailed to get Source WebPart Properties: on Web: " + objInput.WebUrl + " for StorageKey " + _storageKey + "for page " + _relativePageUrl);
                                Console.ForegroundColor = ConsoleColor.Gray;
                            }

                            TranformWebPartStatusBase objWPOutputBase = new TranformWebPartStatusBase();
                            objWPOutputBase.WebApplication = ExceptionCsv.WebApplication;
                            objWPOutputBase.SiteCollection = ExceptionCsv.SiteCollection;
                            objWPOutputBase.WebUrl = objInput.WebUrl;
                            objWPOutputBase.WebPartType = objInput.WebPartType;
                            objWPOutputBase.ZoneID = objInput.ZoneID;
                            objWPOutputBase.ZoneIndex = objInput.ZoneIndex;
                            objWPOutputBase.WebPartId = objInput.WebPartId;
                            objWPOutputBase.PageUrl = objInput.PageUrl;

                            if (status)
                            {
                                objWPOutputBase.Status = "Successfully Replaced WebPart";
                            }
                            else
                            {
                                objWPOutputBase.Status = "Failed to Replace WebPart";
                            }

                            FileUtility.WriteCsVintoFile(outPutFolder + @"\" + System.IO.Path.GetFileNameWithoutExtension(usageFileName) + "_ReplaceOperationStatus.csv", objWPOutputBase, ref headerTransformWebPart);
                            
                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Error in Processing Web:" + ExceptionCsv.WebUrl + " `\r\nError Details:" + ex.Message + " `\r\nExceptionComments:" + exceptionCommentsInfo1);
                            Console.ForegroundColor = ConsoleColor.Gray;
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception][ReplaceWebPart_UsingCSV] Exception Message: " + ex.Message);
                            ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "ReplaceWebPart_UsingCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                        }
                    }

                }

                if (ActionType.ToLower() != Constants.ActionType_All.ToLower())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "##############  Web Part  Trasnformation Utility Execution Completed  - ReplaceWebPart_UsingCSV - InputCSV ##############");
                    Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  - ReplaceWebPart_UsingCSV - InputCSV ##############");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("[Exception][ReplaceWebPart_UsingCSV] Exception Message: " + ex.Message+", ExceptionComments:"+exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception][ReplaceWebPart_UsingCSV] Exception Message: " + ex.Message);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "ReplaceWebPart_UsingCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
            }
        }

        private void ReadWebPartUsageCSV(string sourceWebPartType, string usageFilePath, string outPutFolder, out IEnumerable<WebPartDiscoveryInput> objWPDInput, string SharePointOnline_OR_OnPremise = "OP", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReadWebPartUsageCSV] [START] Calling function ImportCsv.ReadMatchingColumns<WebPartDiscoveryInput>. WebPart Usage Discovery Input CSV file is available at " + outPutFolder + " and Input file name is " + Constants.WebPart_DiscoveryFile_Input);
            Console.WriteLine("[ReadWebPartUsageCSV] [START] Calling function ImportCsv.ReadMatchingColumns<WebPartDiscoveryInput>. WebPart Usage Discovery Input CSV file is available at " + usageFilePath);
            
            objWPDInput = null;
            objWPDInput = ImportCsv.ReadMatchingColumns<WebPartDiscoveryInput>(usageFilePath, Transformation.PowerShell.Common.Constants.CsvDelimeter);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ReadWebPartUsageCSV] [END] Read all the WebParts Usage Details from Discovery Usage File and saved in List - out IEnumerable<WebPartDiscoveryInput> objWPDInput, for processing.");
            Console.WriteLine("[ReadWebPartUsageCSV] [END] Read all the WebParts Usage Details from Discovery Usage File and saved in List - out IEnumerable<WebPartDiscoveryInput> objWPDInput, for processing.");

            try
            {
                if (objWPDInput.Any())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ReadWebPartUsageCSV - After Loading InputCSV ");
                    Console.WriteLine("[START] ReadWebPartUsageCSV - After Loading InputCSV");

                    objWPDInput = from p in objWPDInput
                                  where p.WebPartType == sourceWebPartType
                                  select p;
                    exceptionCommentsInfo1 = objWPDInput.ToString();

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ReadWebPartUsageCSV - After Loading InputCSV");
                    Console.WriteLine("[END] ReadWebPartUsageCSV - After Loading InputCSV");
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "ReadWebPartUsageCSV", ex.GetType().ToString(), exceptionCommentsInfo1);

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][ReadWebPartUsageCSV] Exception Message: " + ex.Message + ", Exception Comments:" + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        public void TransformWebPart_UsingCSV(string usageFileName, string sourceWebPartType, string targetWebPartFileName, string targetWebPartXmlFilePath, string outPutDirectory, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
             string exceptionCommentsInfo1 = string.Empty;

             try
             {
                 WebPart_Initialization(outPutDirectory);
                 
                 //Delete Replace OutPut File
                 FileUtility.DeleteFiles(outPutDirectory + @"\" + System.IO.Path.GetFileNameWithoutExtension(usageFileName) + "_ReplaceOperationStatus.csv");

                 Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Web Part  Trasnformation Utility Execution Started - TransformWebPart_UsingCSV ##############");
                 Console.WriteLine("############## Web Part Trasnformation Utility Execution Started - TransformWebPart_UsingCSV  ##############");

                 Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                 Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());
                 
                 Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][TransformWebPart_UsingCSV]");
                 Console.WriteLine("[START][TransformWebPart_UsingCSV]");
                 
                 Logger.AddMessageToTraceLogFile(Constants.Logging, "[TransformWebPart_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutDirectory);
                 Console.WriteLine("[TransformWebPart_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutDirectory);

                 string sourceWebPartXmlFilesDir = outPutDirectory + @"\" + Constants.SOURCE_WEBPART_XML_DIR;

                 if (!System.IO.Directory.Exists(sourceWebPartXmlFilesDir))
                 {
                     System.IO.Directory.CreateDirectory(sourceWebPartXmlFilesDir);
                 }

                 //Moved Below calls to ReplaceWebPart_UsingCSV - Issue #107616 

                 //GetWebPartProperties_UsingCSV
                 //GetWebPartProperties_UsingCSV(sourceWebPartType, usageFileName, outPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain,Constants.ActionType_All.ToLower());
                 
                 //ConfigureNewWebPartXml
                 //ConfigureNewWebPartXml(targetWebPartXmlFilePath, sourceWebPartXmlFilesDir, outPutDirectory, Constants.ActionType_All.ToLower());

                 //End Moved Below calls to ReplaceWebPart_UsingCSV - Issue #107616 

                 //ReplaceWebPart_UsingCSV
                 string targetWebPartXmlsDir = outPutDirectory + @"\" + Constants.TARGET_WEBPART_XML_DIR;
                 ReplaceWebPart_UsingCSV(sourceWebPartType, targetWebPartXmlFilePath, targetWebPartFileName, targetWebPartXmlsDir, usageFileName, outPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain, Constants.ActionType_All.ToLower());

                 Console.WriteLine("[END][TransformWebPart_UsingCSV] ");
                 Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][TransformWebPart_UsingCSV] ");

                 Logger.AddMessageToTraceLogFile(Constants.Logging, "##############  Web Part  Trasnformation Utility Execution Completed  - TransformWebPart_UsingCSV - InputCSV ##############");
                 Console.WriteLine("############## Web Part Trasnformation Utility Execution Completed  - TransformWebPart_UsingCSV - InputCSV ##############");
             }
             catch (Exception ex)
             {
                 ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "Web Part", ex.Message, ex.ToString(), "TransformWebPart_UsingCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                 Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][TransformWebPart_UsingCSV] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                 Console.ForegroundColor = ConsoleColor.Red;
                 Console.WriteLine("[EXCEPTION][TransformWebPart_UsingCSV] Exception Message: " + ex.Message);
                 Console.ForegroundColor = ConsoleColor.Gray;
             }
        }
        
        public void SetExceptionCSVWebAppSiteColWebUrls(string webUrl, string sharePointOnline_OR_OnPremise, string userName, string password, string domain)
        {
            ClientContext clientContext = null;
            AuthenticationHelper ObjAuth = new AuthenticationHelper();
            
            //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (sharePointOnline_OR_OnPremise.ToUpper() == Constants.OnPremise)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][SetExceptionCSVWebAppSiteColWebUrls] GetNetworkCredentialAuthenticatedContext for WebUrl: " + webUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(webUrl, userName, password, domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][SetExceptionCSVWebAppSiteColWebUrls] GetNetworkCredentialAuthenticatedContext for WebUrl: " + webUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (sharePointOnline_OR_OnPremise.ToUpper() == Constants.Online)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][SetExceptionCSVWebAppSiteColWebUrls] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + webUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(webUrl, userName, password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][SetExceptionCSVWebAppSiteColWebUrls] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + webUrl);
                }
                if (clientContext != null)
                {
                    Web web = clientContext.Web;
                    clientContext.Load(web, w => w.Url,
                                            w => w.ServerRelativeUrl);
                    clientContext.ExecuteQuery();

                    Site site = clientContext.Site;
                    clientContext.Load(site);
                    clientContext.ExecuteQuery();

                    ExceptionCsv.WebApplication = web.Url.Replace(web.ServerRelativeUrl, "");
                    ExceptionCsv.SiteCollection = site.Url;                
                    ExceptionCsv.WebUrl = web.Url;

                }
           
        }
    }
}
