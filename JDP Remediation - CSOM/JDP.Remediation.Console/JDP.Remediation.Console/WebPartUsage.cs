using JDP.Remediation.Console.Common.Base;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Utilities;
using JDP.Remediation.Console;
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
using JDP.Remediation.Console.WebPartPagesService;


namespace JDP.Remediation.Console
{
    public class WebPartUsage
    {
        public static string filePath = string.Empty;
        public static string outputPath = string.Empty;
        public static string timeStamp = string.Empty;
        public static void DoWork()
        {
            outputPath = Environment.CurrentDirectory;
            string webUrl = string.Empty;
            string webPartType = string.Empty;
            timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");

            //Trace Log TXT File Creation Command
            Logger.OpenLog("WebPartUsage", timeStamp);

            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Please enter Web Url : ");
            System.Console.ResetColor();
            webUrl = System.Console.ReadLine().ToLower();
            if (string.IsNullOrEmpty(webUrl))
            {
                Logger.LogErrorMessage("[WebPartUsage: DoWork]WebUrl should not be empty or null. Operation aborted...", true);
                return;
            }

            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Please enter Web Part type: ");
            System.Console.ResetColor();
            webPartType = System.Console.ReadLine().ToLower();
            if (string.IsNullOrEmpty(webPartType))
            {
                Logger.LogErrorMessage("[WebPartUsage: DoWork]WebPart Type should not be empty or null. Operation aborted...", true);
                return;
            }

            Logger.LogInfoMessage(String.Format("Process started {0}", DateTime.Now.ToString()), true);
            try
            {
                GetWebPartUsage(webPartType, webUrl, outputPath);
            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[WebPartUsage: DoWork]. Exception Message: " + ex.Message, true);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "WebPartUsage", ex.Message, ex.ToString(), "WebPartUsage: DoWork()", ex.GetType().ToString());
            }
            Logger.LogInfoMessage(String.Format("Process completed {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        public static void GetWebPartUsage(string webPartType, string webUrl, string outPutDirectory)
        {

            string exceptionCommentsInfo1 = string.Empty;
            ClientContext clientContext = null;
            bool headerWebPart = false;
            Web web = null;
            bool isfound = false;
            string webPartUsageFileName = outPutDirectory + "\\" + Constants.WEBPART_USAGE_ENTITY_FILENAME + timeStamp + Constants.CSVExtension;

            try
            {
                Logger.LogInfoMessage("[GetWebPartUsage] Finding WebPartUsage details for Web Part: " + webPartType + " in Web: " + webUrl);
                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    web = userContext.Web;
                    userContext.Load(web);
                    userContext.ExecuteQuery();
                    clientContext = userContext;

                    WebPartUsageEntity webPartUsageEntity = null;
                    List<WebPartUsageEntity> webPartUsage = new List<WebPartUsageEntity>();

                    //Prepare Exception Comments
                    exceptionCommentsInfo1 = "Web Url: " + clientContext.Web.Url + ", Web Part Type: " + webPartType;

                    if (clientContext != null)
                    {
                        List list = AddWebPart.GetPageList(ref clientContext);
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
                                try
                                {
                                    Microsoft.SharePoint.Client.File page = item.File;

                                    String pageUrl = page.ServerRelativeUrl;// item.FieldValues["EncodedAbsUrl"].ToString();

                                    Logger.LogInfoMessage("[GetWebPartUsage] Checking for the Web Part on the Page: " + page.Name);


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

                                        string actualWebPartType = AddWebPart.GetWebPartShortTypeName(WebPartTypeName);

                                        // only modify if we find the old web part
                                        if (actualWebPartType.Equals(webPartType, StringComparison.CurrentCultureIgnoreCase))
                                        {
                                            Logger.LogInfoMessage("[GetWebPartUsage] Found WebPart: " + webPartType + " in Page: " + page.Name);

                                            webPartUsageEntity = new WebPartUsageEntity();
                                            webPartUsageEntity.PageUrl = pageUrl;
                                            webPartUsageEntity.WebPartID = webPartDefinition.Id.ToString();
                                            webPartUsageEntity.WebURL = webUrl;
                                            webPartUsageEntity.WebPartTitle = webPart.Title;
                                            webPartUsageEntity.ZoneIndex = webPart.ZoneIndex.ToString();
                                            webPartUsageEntity.WebPartType = actualWebPartType;

                                            FileUtility.WriteCsVintoFile(webPartUsageFileName, webPartUsageEntity, ref headerWebPart);
                                            isfound = true;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "WebPartUsage", ex.Message, ex.ToString(), "GetWebPartUsage()", ex.GetType().ToString(), exceptionCommentsInfo1);
                                    Logger.LogErrorMessage("[GetWebPartUsage] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                                    System.Console.ForegroundColor = System.ConsoleColor.Red;
                                    Logger.LogErrorMessage("[GetWebPartUsage] Exception Message: " + ex.Message);
                                    System.Console.ResetColor();
                                }

                            }
                        }
                    }
                    //Default Pages
                    GetWebPartUsage_DefaultPages(webPartType, clientContext, outPutDirectory);
                    //Default Pages

                    if (isfound)
                        Logger.LogSuccessMessage("[GetWebPartUsage] WebPart Usage is exported to the file " + webPartUsageFileName, true);

                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "WebPartUsage", ex.Message, ex.ToString(), "GetWebPartUsage()", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.LogErrorMessage("[GetWebPartUsage] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[GetWebPartUsage] Exception Message: " + ex.Message);
                System.Console.ResetColor();
            }
        }

        public static string GetWebPartPropertiesServiceCall(ClientContext clientContext, string storageKey, string pageUrl)
        {
            string webPartXml = string.Empty;

            var service = new WebPartPagesService.WebPartPagesWebService();
            service.Url = clientContext.Web.Url + Constants.WEBPART_SERVICE;

            service.PreAuthenticate = true;

            service.Credentials = clientContext.Credentials;

            // Actual web service call which returns the information in string format
            webPartXml = service.GetWebPart2(pageUrl, storageKey.ToGuid(), Storage.Shared, SPWebServiceBehavior.Version3);

            return webPartXml;
        }

        public static void GetWebPartUsage_DefaultPages(string webPartType, ClientContext clientContext, string outPutDirectory)
        {
            ExceptionCsv.WebUrl = clientContext.Web.Url;
            string exceptionCommentsInfo1 = string.Empty;
            string webPartUsageFileName = outPutDirectory + "\\" + Constants.WEBPART_USAGE_ENTITY_FILENAME + timeStamp + Constants.CSVExtension;

            bool headerWebPart = false;
            bool isfound = false;
            try
            {
                string webUrl = clientContext.Web.Url;

                Logger.LogInfoMessage("[START][GetWebPartUsage_DefaultPages]");
                Logger.LogInfoMessage("[GetWebPartUsage_DefaultPages] Finding WebPartUsage details for Web Part: " + webPartType + " in Web: " + webUrl);

                WebPartUsageEntity webPartUsageEntity = null;
                List<WebPartUsageEntity> webPartUsage = new List<WebPartUsageEntity>();

                if (clientContext != null)
                {
                    clientContext.Load(clientContext.Web.RootFolder.Files);
                    clientContext.ExecuteQuery();

                    foreach (Microsoft.SharePoint.Client.File page in clientContext.Web.RootFolder.Files)
                    {
                        exceptionCommentsInfo1 = "Web Url: " + clientContext.Web.Url + ", Web Part Type: " + webPartType + ", PageTitle: " + page.ServerRelativeUrl;

                        try
                        {
                            if (Path.GetExtension(page.ServerRelativeUrl).Equals(".aspx", StringComparison.CurrentCultureIgnoreCase))
                            {
                                String pageUrl = page.ServerRelativeUrl;

                                Logger.LogInfoMessage("[GetWebPartUsage_DefaultPages] Checking for the Web Part on the Page: " + page.Name);

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

                                    string actualWebPartType = AddWebPart.GetWebPartShortTypeName(WebPartTypeName);

                                    // only modify if we find the old web part
                                    if (actualWebPartType.Equals(webPartType, StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        Logger.LogInfoMessage("[GetWebPartUsage_DefaultPages] Found WebPart: " + webPartType + " in Page: " + page.Name + ", " + page.ServerRelativeUrl);

                                        webPartUsageEntity = new WebPartUsageEntity();
                                        webPartUsageEntity.PageUrl = pageUrl;
                                        webPartUsageEntity.WebPartID = webPartDefinition.Id.ToString();
                                        webPartUsageEntity.WebURL = webUrl;
                                        webPartUsageEntity.WebPartTitle = webPart.Title;
                                        webPartUsageEntity.ZoneIndex = webPart.ZoneIndex.ToString();
                                        webPartUsageEntity.WebPartType = actualWebPartType;

                                        FileUtility.WriteCsVintoFile(webPartUsageFileName, webPartUsageEntity, ref headerWebPart);
                                        isfound = true;
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "WebPartUsage", ex.Message, ex.ToString(), "GetWebPartUsage_DefaultPages()", ex.GetType().ToString(), exceptionCommentsInfo1);
                            Logger.LogErrorMessage("[GetWebPartUsage_DefaultPages] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                            System.Console.ForegroundColor = System.ConsoleColor.Red;
                            Logger.LogErrorMessage("[GetWebPartUsage_DefaultPages] Exception Message: " + ex.Message);
                            System.Console.ResetColor();
                        }

                    }
                }

                if (isfound)
                    Logger.LogInfoMessage("[GetWebPartUsage_DefaultPages] Default Pages WebPart Usage is exported to the file " + webPartUsageFileName);

                Logger.LogInfoMessage("[END][GetWebPartUsage_DefaultPages]");

            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[GetWebPartUsage_DefaultPages] Exception Message: " + ex.Message);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, ExceptionCsv.WebUrl, "WebPartUsage", ex.Message, ex.ToString(), "GetWebPartUsage_DefaultPages()", ex.GetType().ToString(), exceptionCommentsInfo1);
            }
        }


    }
}
