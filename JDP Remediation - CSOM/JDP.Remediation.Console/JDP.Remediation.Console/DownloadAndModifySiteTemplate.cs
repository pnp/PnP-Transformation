using JDP.Remediation.Console;
using JDP.Remediation.Console.Common.Base;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Utilities;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace JDP.Remediation.Console
{
    public class DownloadAndModifySiteTemplate
    {
        public static List<string> lstContentTypeIDs;
        public static List<string> lstCustomFieldIDs;
        public static List<string> lstCustomErs;
        public static List<string> lstCustomFeatureIDs;
        public static string filePath = string.Empty;
        public static string outputPath = string.Empty;

        public static void DoWork()
        {
            bool processInputFile = false;
            bool processFarm = false;
            bool processSiteCollections = false;
            string webApplicationUrl = string.Empty;
            string siteTemplateInputFile = string.Empty;
            string siteCollectionUrlsList = string.Empty;
            string[] siteCollectionUrls = null;

            try
            {
                //Output files
                outputPath = Environment.CurrentDirectory;

                //Trace Log TXT File Creation Command
                Logger.OpenLog("DownloadAndModifySiteTemplate");
                Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: DoWork] Logger and Exception file will be available in path: " + outputPath);

                //User Options
                if (!ReadInputOptions(ref processInputFile, ref processFarm, ref processSiteCollections))
                {
                    System.Console.ForegroundColor = System.ConsoleColor.Red;
                    Logger.LogMessage("Invalid option selected. Operation aborted!");
                    System.Console.ResetColor();
                    return;
                }

                //Web Application Urls [If option 1 Selected]
                if (processFarm)
                {
                    if (!ReadWebApplication(ref webApplicationUrl))
                    {
                        System.Console.ForegroundColor = System.ConsoleColor.Red;
                        Logger.LogMessage("WebApplicationUrl is not valid. So, Operation aborted!");
                        System.Console.ResetColor();
                        return;
                    }
                }

                //SiteCollection Urls separated by comma(,) -  [If option 2 Selected]
                if (processSiteCollections)
                {
                    if (!ReadSiteCollectionList(ref siteCollectionUrlsList))
                    {
                        System.Console.ForegroundColor = System.ConsoleColor.Red;
                        Logger.LogMessage("SiteCollectionUrls is not valid. So, Operation aborted!");
                        System.Console.ResetColor();
                        return;
                    }
                    siteCollectionUrls = siteCollectionUrlsList.Split(',');
                }

                //Site Template CSV Path [If option 3 Selected]
                if (processInputFile)
                {
                    if (!ReadInputFile(ref siteTemplateInputFile))
                    {
                        System.Console.ForegroundColor = System.ConsoleColor.Red;
                        Logger.LogMessage("SiteTemplate input file is not valid or available. So, Operation aborted!");
                        Logger.LogMessage("Please enter path like: E.g. C:\\<Working Directory>\\<InputFile>.csv");
                        System.Console.ResetColor();
                        return;
                    }
                }

                //Validating Input File Path/Folder
                if (processInputFile || processFarm || processSiteCollections)
                {
                    //Input Files Path: EventReceivers.csv, ContentTypes.csv, CustomFields.csv and Features.csv
                    if (!ReadInputFilesPath())
                    {
                        System.Console.ForegroundColor = System.ConsoleColor.Red;
                        Logger.LogMessage("Input files directory is not valid. So, Operation aborted!");
                        System.Console.ResetColor();
                        return;
                    }
                    ReadInputFiles();
                }

                //Output File - Intermediate File, which will have info about Customization
                string csvFileName = outputPath + @"\" + Constants.SiteTemplateCustomizationUsage;
                bool headerOfCsv = true;
                List<SiteTemplateFTCAnalysisOutputBase> lstMissingSiteTempaltesInGalleryBase = new List<SiteTemplateFTCAnalysisOutputBase>();

                if (!(lstContentTypeIDs != null && lstContentTypeIDs.Any())
                    && !(lstCustomErs != null && lstCustomErs.Any())
                    && !(lstCustomFieldIDs != null && lstCustomFieldIDs.Any())
                    && !(lstCustomFeatureIDs != null && lstCustomFeatureIDs.Any()))
                {
                    System.Console.ForegroundColor = System.ConsoleColor.Red;
                    Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: DoWork] No records present in input files (Features.csv, EventReceivers.csv, ContentTypes.csv & CustomFields.csv) to check the customization. So, Operation aborted!", true);
                    WriteOutputReport(null, csvFileName, ref headerOfCsv);
                    System.Console.ResetColor();
                    return;
                }

                #region Site Template report based on input file
                if (processInputFile)
                {
                    if (System.IO.File.Exists(siteTemplateInputFile))
                    {
                        //Process SiteTemplateInputFile
                        ProcessSiteTemplateInputFile(siteTemplateInputFile, ref lstMissingSiteTempaltesInGalleryBase);
                        WriteOutputReport(lstMissingSiteTempaltesInGalleryBase, csvFileName, ref headerOfCsv);
                    }
                    else
                    {
                        Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: DoWork]. Exception Message: Site Template: " + filePath + @"\" + Constants.SiteTemplateCustomizationUsage + " is not present", true);
                    }
                }
                #endregion

                lstMissingSiteTempaltesInGalleryBase.Clear();

                #region SiteTemplate based on web application url
                if (processFarm)
                {
                    //Process WebApplicationUrl
                    ProcessWebApplicationUrl(webApplicationUrl, ref lstMissingSiteTempaltesInGalleryBase);
                    WriteOutputReport(lstMissingSiteTempaltesInGalleryBase, csvFileName, ref headerOfCsv);
                }
                #endregion

                lstMissingSiteTempaltesInGalleryBase.Clear();

                #region SiteTemplate based on SiteCollectionUrls list
                if (processSiteCollections)
                {
                    //Process SiteCollection Urls
                    ProcessSiteCollectionUrlsList(siteCollectionUrls, ref lstMissingSiteTempaltesInGalleryBase);
                    WriteOutputReport(lstMissingSiteTempaltesInGalleryBase, csvFileName, ref headerOfCsv);
                }
                #endregion

                if (processFarm || processInputFile || processSiteCollections)
                {
                    //Delete Downloaded SiteTemplates files/folders
                    DeleteDownloadedSiteTemplates();

                    System.Console.ForegroundColor = System.ConsoleColor.Green;
                    Logger.LogSuccessMessage("[DownloadAndModifySiteTemplate: DoWork] Successfully completed all SiteTemplate and output file is present at the path: "
                        + csvFileName, true);
                    System.Console.ResetColor();
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: DoWork]. Exception Message: " + ex.Message + ", Exception Comments: ", true);
            }
            finally
            {
                filePath = null;
                outputPath = null;
                lstContentTypeIDs = null;
                lstCustomErs = null;
                lstCustomFeatureIDs = null;
                lstCustomFieldIDs = null;
            }
            Logger.CloseLog();
        }

        public static bool DownloadSiteTemplate(string filePath, string SiteTemplateGalleryPath, string SiteTemplateName, string SiteCollection, string WebUrl)
        {
            Logger.LogInfoMessage("[DownloadSiteTemplate] Downloading the Site Template: " + SiteTemplateName, true);
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
            bool isDownloaded = false;

            try
            {
                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, SiteCollection))
                {
                    userContext.ExecuteQuery();

                    Site site = userContext.Site;
                    Web web = userContext.Web;
                    userContext.Load(site);
                    userContext.Load(web);
                    userContext.ExecuteQuery();

                    List siteTemplateGallery = site.GetCatalog(121);
                    userContext.Load(siteTemplateGallery);
                    userContext.Load(siteTemplateGallery.RootFolder);
                    userContext.ExecuteQuery();

                    string siteGalleryServerRelativeUrl = siteTemplateGallery.RootFolder.ServerRelativeUrl;

                    //To remove errors resulting from Eval Sites
                    if (siteGalleryServerRelativeUrl == SiteTemplateGalleryPath && SiteCollection == WebUrl)
                    {
                        string fileUrl = SiteTemplateGalleryPath + "/" + SiteTemplateName; ///sites/EvalSitetesting-eval

                        FileInformation info = Microsoft.SharePoint.Client.File.OpenBinaryDirect(userContext, fileUrl);
                        string fileName = fileUrl.Substring(fileUrl.LastIndexOf("/") + 1);

                        var fileNamePath = Path.Combine(filePath, fileName);
                        using (var fileStream = System.IO.File.Create(fileNamePath))
                        {
                            info.Stream.CopyTo(fileStream);
                            isDownloaded = true;
                            Logger.LogInfoMessage("[DownloadSiteTemplate] Successfully Downloaded Site Template " + SiteTemplateName, true);
                        }
                    }
                    else
                    {
                        Logger.LogInfoMessage("[DownloadSiteTemplate] Download Failed for " + SiteTemplateName + ". SiteGalleryPath is not present in the current Site Collection: ", true);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DownloadSiteTemplate]. Exception Message: " + ex.Message + ", Exception Comments: SiteGalleryPath is not present in the current Site Collection", true);
            }
            return isDownloaded;
        }

        private static bool ProcessWspFile(string filePath, string solFileName, ref SiteTemplateFTCAnalysisOutputBase objSiteCustOutput)
        {
            string receiversXml = string.Empty;
            string contentTypesXml = string.Empty;
            string customFieldsTypesXml = string.Empty;
            bool isCustomizationPresent = false;
            bool isCustomContentType = false;
            bool isCustomEventReceiver = false;
            bool isCustomSiteColumn = false;
            bool isCustomFeature = false;

            string fileName = objSiteCustOutput.SiteTemplateName;
            string downloadPath = filePath + @"\" + Constants.DownloadPathSiteTemplates;

            Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Processing the Site Template: " + objSiteCustOutput.SiteTemplateName, true);
            try
            {
                string newFilePath = solFileName.Replace(".wsp", ".cab");
                if (System.IO.File.Exists(newFilePath))
                    System.IO.File.Delete(newFilePath);
                System.IO.File.Move(solFileName, newFilePath);

                var destDir = newFilePath.Substring(0, newFilePath.LastIndexOf(@"\"));
                Directory.SetCurrentDirectory(destDir);
                string newFileName = newFilePath.Substring(newFilePath.LastIndexOf(@"\") + 1);

                FileInfo solFileObj = new FileInfo(newFileName);
                Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Extracting the Site Template: " + objSiteCustOutput.SiteTemplateName, true);
                string cmd = "/e /a /y /L \"" + newFileName.Replace(".", "_") + "\" \"" + newFileName + "\"";
                ProcessStartInfo pI = new ProcessStartInfo("extrac32.exe", cmd);
                Process p = Process.Start(pI);
                p.WaitForExit();
                string cabDir = newFilePath.Replace(".", "_");
                Directory.SetCurrentDirectory(newFileName.Replace(".", "_"));
                string extractedPath = solFileName.Replace(".wsp", "_cab");
                Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Extracted the Site Template: " + objSiteCustOutput.SiteTemplateName + "to path: " + extractedPath, true);

                string[] webTempFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "Onet.xml", SearchOption.AllDirectories);

                if (webTempFiles.Length > 0)
                {
                    string fileNameWithoutExtension = fileName.Substring(0, fileName.LastIndexOf("."));

                    string listInstanceFolderName = extractedPath + @"\" + fileNameWithoutExtension + "ListInstances";

                    string[] featureFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "Feature.xml", SearchOption.AllDirectories);

                    #region Custom Features

                    Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Searching for Onet.xml file for finding customized features in path: " + webTempFiles.ElementAt(0), true);

                    if (lstCustomFeatureIDs != null && lstCustomFeatureIDs.Count > 0)
                    {
                        if (webTempFiles.Count() > 0)
                        {
                            #region Custom SiteFeatures
                            try
                            {
                                CheckCustomFeature(webTempFiles.ElementAt(0), "/Project/Configurations/Configuration/SiteFeatures/Feature", ref isCustomFeature, solFileName);
                            }
                            catch (Exception ex)
                            {
                                Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWspFile]. Exception Message: " + ex.Message + ", Exception Comments: Exception while reading Custom Site Features tag", true);
                            }
                            #endregion

                            #region Custom WebFeatures
                            if (!isCustomFeature)
                            {
                                try
                                {
                                    CheckCustomFeature(webTempFiles.ElementAt(0), "/Project/Configurations/Configuration/WebFeatures/Feature", ref isCustomFeature, solFileName);
                                }
                                catch (Exception ex)
                                {
                                    Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWspFile]. Exception Message: " + ex.Message + ", Exception Comments: Exception while reading Custom Web Features tag", true);
                                }
                            }
                            #endregion
                        }

                        if (featureFiles.Count() > 0 && !isCustomFeature)
                        {
                            for (int i = 0; i < featureFiles.Count(); i++)
                            {
                                try
                                {
                                    Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Searching for Feature.xml file for finding customized features in path: " + featureFiles.ElementAt(i), true);

                                    CheckCustomFeature(featureFiles.ElementAt(i), "/Feature", ref isCustomFeature, solFileName);
                                }
                                catch (Exception ex)
                                {
                                    Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWspFile]. Exception Message: " + ex.Message + ", Exception Comments: Exception while reading Features tag", true);
                                }
                            }
                        }
                    }
                    #endregion

                    #region Web EventReceivers
                    Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Searching for WebEventReceivers Folder for finding customized Event Receivers in path: " + extractedPath, true);
                    IEnumerable<string> ERlist = Directory.GetDirectories(extractedPath).Where(s => s.EndsWith("WebEventReceivers"));

                    if (lstCustomErs != null && lstCustomErs.Count > 0)
                    {
                        if (ERlist.Count() > 0)
                        {
                            if (ERlist.ElementAt(0).EndsWith(@"\"))
                                receiversXml = ERlist.ElementAt(0) + "Elements.xml";
                            else
                                receiversXml = ERlist.ElementAt(0) + @"\" + "Elements.xml";

                            CheckCustomEventReceiver(receiversXml, "/Elements/Receivers", ref isCustomEventReceiver, solFileName);
                        }
                    }
                    #endregion


                    Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Searching for ListInstances Folder for finding customized elements in path: " + extractedPath, true);
                    IEnumerable<string> list = Directory.GetDirectories(extractedPath).Where(s => s.EndsWith("ListInstances"));

                    if (list.Count() > 0 & !isCustomEventReceiver)
                    {
                        #region Custom List EventReceivers

                        if (list.ElementAt(0).EndsWith(@"\"))
                            receiversXml = list.ElementAt(0) + "Elements.xml";
                        else
                            receiversXml = list.ElementAt(0) + @"\" + "Elements.xml";


                        CheckCustomEventReceiver(receiversXml, "/Elements/Receivers", ref isCustomEventReceiver, solFileName);

                        #endregion

                        //Reading ElementContentTypes.xml for Searching Content Types
                        if (list.ElementAt(0).EndsWith(@"\"))
                            contentTypesXml = list.ElementAt(0) + "ElementsContentType.xml";
                        else
                            contentTypesXml = list.ElementAt(0) + @"\" + "ElementsContentType.xml";

                        if (System.IO.File.Exists(contentTypesXml))
                        {
                            Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Searching for customized Content Types and Customized Event Receivers or Site Columns associated in Content Types in: " + contentTypesXml, true);
                            var reader = new XmlTextReader(contentTypesXml);

                            reader.Namespaces = false;
                            reader.Read();
                            XmlDocument doc2 = new XmlDocument();
                            doc2.Load(reader);

                            //Initiallizing all the nodes required to check
                            XmlNodeList xmlDocReceivers = doc2.SelectNodes("/Elements/ContentType");

                            #region Custom ContentType Event Receviers
                            if (lstCustomErs != null && lstCustomErs.Count > 0 && !isCustomEventReceiver)
                            {
                                for (int i = 0; i < xmlDocReceivers.Count; i++)
                                {
                                    try
                                    {
                                        if (xmlDocReceivers[i].HasChildNodes)
                                        {
                                            var docList = xmlDocReceivers[i]["XmlDocuments"];
                                            if (docList != null)
                                            {
                                                XmlNodeList xmlDocList = docList.ChildNodes;

                                                for (int j = 0; j < xmlDocList.Count; j++)
                                                {
                                                    try
                                                    {
                                                        if (xmlDocList[j].Attributes["NamespaceURI"] != null)
                                                        {
                                                            var namespaceURl = xmlDocList[j].Attributes["NamespaceURI"].Value;
                                                            if (namespaceURl.Contains("http://schemas.microsoft.com/sharepoint/events"))
                                                            {
                                                                XmlNodeList child = xmlDocList[j].ChildNodes;

                                                                if (child != null && child.Count > 0)
                                                                {
                                                                    for (int x = 0; x < child.Count; x++)
                                                                    {
                                                                        XmlNodeList receiverChilds = child[x].ChildNodes;
                                                                        for (int y = 0; y < receiverChilds.Count; y++)
                                                                        {
                                                                            try
                                                                            {
                                                                                if (receiverChilds[y].HasChildNodes)
                                                                                {
                                                                                    string ctAssemblyValue = receiverChilds[y]["Assembly"].InnerText;
                                                                                    if (lstCustomErs.Where(c => ctAssemblyValue.Equals(c, StringComparison.CurrentCultureIgnoreCase)).Any())
                                                                                    {
                                                                                        isCustomEventReceiver = true;
                                                                                        Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Customized Event Receiver in Content Type Found for: " + objSiteCustOutput.SiteTemplateName, true);
                                                                                        break;
                                                                                    }
                                                                                }
                                                                            }
                                                                            catch (Exception ex)
                                                                            {
                                                                                Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWspFile]. Exception Message: " + ex.Message + ", Exception Comments: Exception while reading Receivers tag in content types", true);
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWspFile]. Exception Message: " + ex.Message + ", Exception Comments: Exception while reading Receivers tag in content types", true);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWspFile]. Exception Message: " + ex.ToString() + ", " + ex.Message, true);
                                    }
                                }
                            }
                            #endregion

                            #region Custom ContentTypes
                            if (lstContentTypeIDs != null && lstContentTypeIDs.Count > 0)
                            {

                                //Iterate all ContentTypes in manifest.xml
                                for (int i = 0; i < xmlDocReceivers.Count; i++)
                                {
                                    try
                                    {
                                        var docList = xmlDocReceivers[i].Attributes["ID"].Value;

                                        //Remove contenttype tag if ContentTypeId present in custom ContentTypes file ContentTypes.csv
                                        if (lstContentTypeIDs.Where(c => docList.StartsWith(c)).Any())
                                        {
                                            isCustomContentType = true;
                                            Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Customized Content Type Found for: " + objSiteCustOutput.SiteTemplateName, true);
                                            break;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWspFile]. Exception Message: " + ex.Message + ", Exception Comments: Exception while reading content types", true);
                                    }
                                }
                            }
                            #endregion

                            #region CustomFields in ContentTypes
                            if (lstCustomFieldIDs != null && lstCustomFieldIDs.Count > 0)
                            {
                                //Checking Site Columns presence in Content Types
                                for (int i = 0; i < xmlDocReceivers.Count; i++)
                                {
                                    try
                                    {
                                        var fieldRefs = xmlDocReceivers[i]["FieldRefs"];
                                        if (fieldRefs != null)
                                        {
                                            XmlNodeList xmlFieldRefList = fieldRefs.ChildNodes;

                                            for (int j = 0; j < xmlFieldRefList.Count; j++)
                                            {
                                                try
                                                {
                                                    string fieldRefId = xmlFieldRefList[j].Attributes["ID"].Value;
                                                    if (lstCustomFieldIDs.Where(c => fieldRefId.Equals(c)).Any())
                                                    {
                                                        isCustomSiteColumn = true;
                                                        Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Customized Site Column associated with Content Type Found for: " + objSiteCustOutput.SiteTemplateName, true);
                                                        break;
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWspFile]. Exception Message: " + ex.Message + ", Exception Comments: Exception while reading Site Columns tag in content types", true);
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWspFile]. Exception Message: " + ex.Message + ", Exception Comments: Exception while reading Site Columns tag in content types", true);
                                    }
                                }
                            }
                            #endregion

                            reader.Dispose();
                        }
                        //Reading ElementContentTypes.xml for Searching Custom Fields
                        if (list.ElementAt(0).EndsWith(@"\"))
                            customFieldsTypesXml = list.ElementAt(0) + "ElementsFields.xml";
                        else
                            customFieldsTypesXml = list.ElementAt(0) + @"\" + "ElementsFields.xml";

                        if (!isCustomSiteColumn && System.IO.File.Exists(customFieldsTypesXml) && !isCustomSiteColumn)
                        {
                            Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Searching for customized Site Columns in List in: " + customFieldsTypesXml, true);
                            var reader = new XmlTextReader(customFieldsTypesXml);

                            reader.Namespaces = false;
                            reader.Read();
                            XmlDocument doc3 = new XmlDocument();
                            doc3.Load(reader);
                            XmlNodeList xmlFields = doc3.SelectNodes("/Elements/Field");

                            #region Custom Fields
                            if (lstCustomFieldIDs != null && lstCustomFieldIDs.Count > 0)
                            {
                                for (int i = 0; i < xmlFields.Count; i++)
                                {
                                    try
                                    {
                                        var fieldList = xmlFields[i].Attributes["ID"].Value;

                                        //Remove contenttype tag if ContentTypeId present in custom ContentTypes file ContentTypes.csv
                                        if (lstCustomFieldIDs.Where(c => fieldList.Equals(c)).Any())
                                        {
                                            isCustomSiteColumn = true;
                                            Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Customized Site Column Found for: " + objSiteCustOutput.SiteTemplateName, true);
                                            break;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWspFile]. Exception Message: " + ex.Message + ", Exception Comments: Exception while reading Site Columns", true);
                                    }
                                }
                            }
                            #endregion
                            reader.Dispose();
                        }
                    }
                    if (isCustomContentType || isCustomEventReceiver || isCustomSiteColumn || isCustomFeature)
                    {
                        objSiteCustOutput.IsCustomizationPresent = "YES";
                        isCustomizationPresent = true;

                        if (lstCustomErs != null && lstCustomErs.Count > 0)
                            objSiteCustOutput.IsCustomizedEventReceiver = isCustomEventReceiver ? "YES" : "NO";
                        else
                            objSiteCustOutput.IsCustomizedEventReceiver = Constants.NoInputFile;

                        if (lstContentTypeIDs != null && lstContentTypeIDs.Count > 0)
                            objSiteCustOutput.IsCustomizedContentType = isCustomContentType ? "YES" : "NO";
                        else
                            objSiteCustOutput.IsCustomizedContentType = Constants.NoInputFile;

                        if (lstCustomFieldIDs != null && lstCustomFieldIDs.Count > 0)
                            objSiteCustOutput.IsCustomizedSiteColumn = isCustomSiteColumn ? "YES" : "NO";
                        else
                            objSiteCustOutput.IsCustomizedSiteColumn = Constants.NoInputFile;

                        if (lstCustomFeatureIDs != null && lstCustomFeatureIDs.Count > 0)
                            objSiteCustOutput.IsCustomizedFeature = isCustomFeature ? "YES" : "NO";
                        else
                            objSiteCustOutput.IsCustomizedFeature = Constants.NoInputFile;
                    }
                    else
                    {
                        isCustomizationPresent = false;
                    }
                }
                else
                {
                    isCustomizationPresent = false;
                }

            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWspFile]. Exception Message: " + ex.Message, true);
            }

            Directory.SetCurrentDirectory(downloadPath);
            System.IO.DirectoryInfo directory = new System.IO.DirectoryInfo(downloadPath);
            foreach (System.IO.FileInfo file in directory.GetFiles()) file.Delete();
            foreach (System.IO.DirectoryInfo subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
            return isCustomizationPresent;
        }

        public static bool GetCustomizedSiteTemplate(ref SiteTemplateFTCAnalysisOutputBase objSiteCustOutput, DataRow siteTemplateRow,
            Microsoft.SharePoint.Client.File ltFile, string siteCollection, string webAppUrl)
        {
            bool isCustomizationPresent = false;
            string fileName = string.Empty;
            string siteGalleryPath = string.Empty;

            try
            {
                fileName = ltFile.Name;
                siteGalleryPath = ltFile.ServerRelativeUrl.Substring(0, ltFile.ServerRelativeUrl.LastIndexOf('/'));
                objSiteCustOutput.SiteCollection = siteCollection;
                objSiteCustOutput.WebApplication = GetWebapplicationUrlFromSiteCollectionUrl(siteCollection);
                objSiteCustOutput.WebUrl = siteCollection;
                objSiteCustOutput.SiteTemplateName = fileName;
                objSiteCustOutput.SiteTemplateGalleryPath = siteGalleryPath;

                bool isDownloaded = DownloadSiteTemplate(outputPath + @"\" + Constants.DownloadPathSiteTemplates,
                    siteGalleryPath, fileName, objSiteCustOutput.WebUrl, objSiteCustOutput.WebUrl);
                if (isDownloaded)
                {
                    isCustomizationPresent = ProcessWspFile(outputPath, outputPath + @"\" + Constants.DownloadPathSiteTemplates + @"\" + fileName,
                        ref objSiteCustOutput);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: GetCustomizedSiteTemplate]. Exception Message: " + ex.Message + ", Exception Comments: ", true);
            }
            return isCustomizationPresent;
        }

        public static void WriteOutputReport(List<SiteTemplateFTCAnalysisOutputBase> ltSiteTemplateOutputBase, string csvFileName, ref bool headerOfCsv)
        {
            try
            {
                Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: WriteOutputReport] Writing the Output file " + Constants.SiteTemplateCustomizationUsage, true);
                if (System.IO.File.Exists(csvFileName))
                    System.IO.File.Delete(csvFileName);
                if (ltSiteTemplateOutputBase != null && ltSiteTemplateOutputBase.Any())
                {
                    //Export the result(Missing Workflow Details) in CSV file                   
                    FileUtility.WriteCsVintoFile(csvFileName, ref ltSiteTemplateOutputBase, ref headerOfCsv);
                }
                else
                {
                    headerOfCsv = false;
                    SiteTemplateFTCAnalysisOutputBase objSiteTemplatesNoInstancesFound = new SiteTemplateFTCAnalysisOutputBase();
                    FileUtility.WriteCsVintoFile(csvFileName, objSiteTemplatesNoInstancesFound, ref headerOfCsv);
                    objSiteTemplatesNoInstancesFound = null;
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: WriteOutputReport] Exception Message: " + ex.Message + ", Exception Comments: ", true);
            }
        }

        public static void ProcessSiteCollectionUrl(string siteCollectionUrl,
            ref List<SiteTemplateFTCAnalysisOutputBase> lstMissingSiteTempaltesInGalleryBase, string webApplicationUrl)
        {
            try
            {
                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, siteCollectionUrl))
                {
                    userContext.ExecuteQuery();
                    Web web = userContext.Web;
                    Folder folder = userContext.Web.GetFolderByServerRelativeUrl("_catalogs/solutions");
                    userContext.Load(web.Folders);
                    userContext.Load(folder);
                    userContext.Load(folder.Files);
                    userContext.Load(web);
                    userContext.ExecuteQuery();

                    // Loop through all the site templates    
                    foreach (Microsoft.SharePoint.Client.File stFile in folder.Files)
                    {
                        try
                        {
                            System.Console.WriteLine("File Name: " + stFile.Name);
                            SiteTemplateFTCAnalysisOutputBase objSiteCustOutput = new SiteTemplateFTCAnalysisOutputBase();
                            string siteTemplatePath = string.Empty;
                            string siteTemplateGalleryPath = stFile.ServerRelativeUrl.Substring(0, stFile.ServerRelativeUrl.LastIndexOf('/'));

                            bool isCustomizationPresent = false;

                            isCustomizationPresent = GetCustomizedSiteTemplate(ref objSiteCustOutput, null, stFile, siteCollectionUrl, webApplicationUrl);

                            if (isCustomizationPresent)
                            {
                                userContext.Load(stFile.Author);
                                userContext.Load(stFile.ModifiedBy);
                                userContext.ExecuteQuery();
                                objSiteCustOutput.CreatedBy = stFile.Author.LoginName;
                                objSiteCustOutput.CreatedDate = stFile.TimeCreated.ToString();
                                objSiteCustOutput.ModifiedBy = stFile.ModifiedBy.LoginName;
                                objSiteCustOutput.ModifiedDate = stFile.TimeLastModified.ToString();
                                lstMissingSiteTempaltesInGalleryBase.Add(objSiteCustOutput);
                            }
                            objSiteCustOutput = null;
                        }
                        catch (Exception ex)
                        {
                            Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: DoWork]. Exception Message: " + ex.Message + ", Exception Comments: ", true);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: DoWork]. Exception Message: " + ex.Message + ", Exception Comments: ", true);
            }
        }

        public static void CheckCustomFeature(string xmlFilePath, string featureNodePath, ref bool isCustomFeature, string siteTemplateName)
        {
            if (System.IO.File.Exists(xmlFilePath))
            {
                var reader = new XmlTextReader(xmlFilePath);
                try
                {
                    reader.Namespaces = false;
                    reader.Read();
                    XmlDocument doc = new XmlDocument();
                    doc.Load(reader);

                    //Initiallizing all the nodes required to check
                    XmlNodeList siteFeatureNodes = doc.SelectNodes(featureNodePath);

                    for (int j = 0; j < siteFeatureNodes.Count; j++)
                    {
                        try
                        {
                            string featureID = siteFeatureNodes[j].Attributes["ID"].Value;
                            if (featureID.StartsWith("{"))
                            {
                                featureID = featureID.TrimStart('{');
                                featureID = featureID.TrimEnd('}');
                            }
                            if (lstCustomFeatureIDs.Where(c => c.Contains(featureID.ToLower())).Any())
                            {
                                isCustomFeature = true;
                                Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: CheckCustomFeature] Customized Feature Found for: " + siteTemplateName, true);
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: CheckCustomFeature]. Exception Message: " + ex.Message + ", Exception Comments: Exception while reading Features tag", true);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWspFile]. Exception Message: " + ex.Message + ", Exception Comments: Exception while reading Features tag", true);
                }
                finally
                {
                    reader.Dispose();
                }
            }
        }

        public static void CheckCustomEventReceiver(string xmlFilePath, string erNodePath, ref bool isCustomEventReceiver, string siteTemplateName)
        {
            if (System.IO.File.Exists(xmlFilePath))
            {
                Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessWspFile] Searching for customized Web/Site Event Receivers in: " + xmlFilePath, true);
                var reader = new XmlTextReader(xmlFilePath);
                try
                {
                    reader.Namespaces = false;
                    reader.Read();
                    XmlDocument doc = new XmlDocument();
                    doc.Load(reader);

                    //Initiallizing all the nodes required to check
                    XmlNodeList receiverNodes = doc.SelectNodes(erNodePath);

                    //Chcecking for Custom Event Receivers
                    if (receiverNodes != null && receiverNodes.Count > 0)
                    {
                        for (int i = 0; i < receiverNodes.Count; i++)
                        {
                            XmlNodeList receiverChilds = receiverNodes[i].ChildNodes;
                            for (int j = 0; j < receiverChilds.Count; j++)
                            {
                                try
                                {
                                    if (receiverChilds[j].HasChildNodes)
                                    {
                                        string assemblyValue = receiverChilds[j]["Assembly"].InnerText;
                                        if (lstCustomErs.Where(c => assemblyValue.Equals(c, StringComparison.CurrentCultureIgnoreCase)).Any())
                                        {
                                            isCustomEventReceiver = true;
                                            Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: CheckCustomEventReceiver] Customized Web/Site/List Event Receiver Found for: " + siteTemplateName, true);
                                            break;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: CheckCustomEventReceiver]. Exception Message: " + ex.Message + ", Exception Comments: Exception while reading Web/Site Receivers tag", true);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWspFile]. Exception Message: " + ex.Message + ", Exception Comments: Exception while reading Features tag", true);
                }
                finally
                {
                    reader.Dispose();
                }
            }
        }

        public static void ReadInputFiles()
        {
            IEnumerable<ContentTypeInput> objCtInput;
            IEnumerable<CustomFieldInput> objFtInput;
            IEnumerable<FeatureInput> objFRInput;
            IEnumerable<EventReceiverInput> objErInput;

            try
            {
                //Content Type Input
                if (System.IO.File.Exists(filePath + @"\" + Constants.ContentTypeInput))
                {
                    Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ReadInputFiles] Reading the ContentTypes.csv Input file", true);
                    objCtInput = ImportCSV.ReadMatchingColumns<ContentTypeInput>(filePath + @"\" + Constants.ContentTypeInput, Constants.CsvDelimeter);
                    lstContentTypeIDs = objCtInput.Select(c => c.ContentTypeID).ToList();
                    lstContentTypeIDs = lstContentTypeIDs.Distinct().ToList();
                }
                else
                {
                    Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ReadInputFiles]. Exception Message: " + filePath + @"\" + Constants.ContentTypeInput + " is not present", true);
                }
                //Custom Field Input
                if (System.IO.File.Exists(filePath + @"\" + Constants.CustomFieldsInput))
                {
                    Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ReadInputFiles] Reading the CustomFields.csv Input file", true);
                    objFtInput = ImportCSV.ReadMatchingColumns<CustomFieldInput>(filePath + @"\" + Constants.CustomFieldsInput, Constants.CsvDelimeter);
                    lstCustomFieldIDs = objFtInput.Select(c => c.ID).ToList();
                    lstCustomFieldIDs = lstCustomFieldIDs.Distinct().ToList();
                }
                else
                {
                    Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ReadInputFiles]. Exception Message: " + filePath + @"\" + Constants.CustomFieldsInput + " is not present", true);
                }
                //Features Input
                if (System.IO.File.Exists(filePath + @"\" + Constants.FeaturesInput))
                {
                    Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ReadInputFiles] Reading the Features.csv Input file", true);
                    objFRInput = ImportCSV.ReadMatchingColumns<FeatureInput>(filePath + @"\" + Constants.FeaturesInput, Constants.CsvDelimeter);
                    lstCustomFeatureIDs = objFRInput.Select(c => c.FeatureID.ToLower()).ToList();
                    lstCustomFeatureIDs = lstCustomFeatureIDs.Distinct().ToList();
                }
                else
                {
                    Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ReadInputFiles]. Exception Message: " + filePath + @"\" + Constants.FeaturesInput + " is not present", true);
                }
                //EventReceivers Input
                if (System.IO.File.Exists(filePath + @"\" + Constants.EventReceiversInput))
                {
                    Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ReadInputFiles] Reading the EventReceivers.csv Input file", true);
                    objErInput = ImportCSV.ReadMatchingColumns<EventReceiverInput>(filePath + @"\" + Constants.EventReceiversInput, Constants.CsvDelimeter);
                    lstCustomErs = objErInput.Select(c => c.Assembly.ToLower()).ToList();
                    lstCustomErs = lstCustomErs.Distinct().ToList();
                }
                else
                {
                    Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ReadInputFiles]. Exception Message: " + filePath + @"\" + Constants.EventReceiversInput + " is not present", true);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ReadInputFiles]. Exception Message: " + ex.Message + ", Exception Comments: ", true);
            }
            finally
            {
                objCtInput = null;
                objErInput = null;
                objFtInput = null;
                objFRInput = null;
            }
        }

        public static bool ReadInputOptions(ref bool processInputFile, ref bool processFarm, ref bool processSiteCollections)
        {
            string processOption = string.Empty;

            Logger.LogMessage("Please select any of following options:");
            Logger.LogMessage("1 - Process with Auto-generated Site Collection Report");
            Logger.LogMessage("2 - Process with PreMT/Discovery SiteTemplate Report");
            Logger.LogMessage("3 - Process with SiteCollectionUrls separated by comma (,)");

            processOption = System.Console.ReadLine();

            if (processOption.Equals("2"))
                processInputFile = true;
            else if (processOption.Equals("1"))
                processFarm = true;
            else if (processOption.Equals("3"))
                processSiteCollections = true;
            else
                return false;
            return true;
        }

        public static bool ReadWebApplication(ref string webApplicationUrl)
        {
            System.Console.ForegroundColor = ConsoleColor.Yellow;
            Logger.LogMessage("!!! NOTE !!!");
            Logger.LogMessage("This operation is intended for use only in PPE; use on PROD at your own risk.");
            Logger.LogMessage("This operation is based on Search Service Result. it would be possible list of Site returned would be different than actual due to permission issue or stale crawl data.");
            Logger.LogMessage("For PROD, it is safer to generate the report via the o365 Self-Service Admin Portal.");

            System.Console.ResetColor();

            Logger.LogMessage("Please enter any one of the web application URL for Context making: ");
            webApplicationUrl = System.Console.ReadLine();

            if (string.IsNullOrEmpty(webApplicationUrl))
                return false;
            return true;
        }

        public static bool ReadSiteCollectionList(ref string siteCollectionUrlsList)
        {
            Logger.LogMessage("Enter SiteCollection URLs separated by comma (,): ");
            siteCollectionUrlsList = System.Console.ReadLine();

            if (string.IsNullOrEmpty(siteCollectionUrlsList))
                return false;
            return true;
        }

        public static bool ReadInputFile(ref string SiteTemplateInputFile)
        {
            Logger.LogMessage("Enter Complete Input File Path of Site Template Report Either Pre-Scan OR Discovery Report.");
            SiteTemplateInputFile = System.Console.ReadLine();
            Logger.LogMessage("[DownloadAndModifySiteTemplate: ReadInputFile] Entered Input File of Site Template Data " + SiteTemplateInputFile, false);
            if (string.IsNullOrEmpty(SiteTemplateInputFile) || !System.IO.File.Exists(SiteTemplateInputFile))
                return false;
            return true;
        }

        public static bool ReadInputFilesPath()
        {
            Logger.LogMessage("Enter the directory of input files for customization analysis (Features.csv, EventReceivers.csv, ContentTypes.csv and CustomFields.csv)");
            Logger.LogMessage("Please refer document for how to create input files to analyze the customization. These files are required to find what customization we are looking inside a template");
            filePath = System.Console.ReadLine();
            Logger.LogMessage("[DownloadAndModifySiteTemplate: ReadInputFilesPath] Entered Input files directory: " + filePath, false);
            if (string.IsNullOrEmpty(filePath) || !System.IO.Directory.Exists(filePath))
                return false;
            return true;
        }

        public static void DeleteDownloadedSiteTemplates()
        {
            try
            {
                //Delete DownloadedSiteTemplate directory if exists
                if (Directory.Exists(outputPath + @"\" + Constants.DownloadPathSiteTemplates))
                {
                    if (Environment.CurrentDirectory.Equals(outputPath + @"\" + Constants.DownloadPathSiteTemplates))
                    {
                        Environment.CurrentDirectory = outputPath;
                        Directory.Delete(outputPath + @"\" + Constants.DownloadPathSiteTemplates, true);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: DeleteDownloadedSiteTemplates]. Exception Message: " + ex.Message + ", Exception Comments: ", true);
            }
        }

        public static void ProcessSiteTemplateInputFile(string siteTemplateInputFile, ref List<SiteTemplateFTCAnalysisOutputBase> lstMissingSiteTempaltesInGalleryBase)
        {
            if (System.IO.File.Exists(siteTemplateInputFile))
            {
                try
                {
                    Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessSiteTemplateInputFile] Checking the Customization Status for Site Templates", true);

                    Logger.LogInfoMessage("[DownloadAndModifySiteTemplate: ProcessSiteTemplateInputFile] Reading the Site Templates Input file: " + siteTemplateInputFile, true);


                    DataTable dtSiteTemplatesInput = new DataTable();
                    dtSiteTemplatesInput = ImportCSV.Read(siteTemplateInputFile, Constants.CsvDelimeter);

                    //Get distinct SiteCollectionUrls from InputFile
                    List<string> lstSiteCollectionUrls = dtSiteTemplatesInput.AsEnumerable()
                                                    .Select(r => r.Field<string>("SiteCollection"))
                                                    .ToList();
                    lstSiteCollectionUrls = lstSiteCollectionUrls.Distinct().ToList();
                    foreach (string siteCollection in lstSiteCollectionUrls)
                    {
                        string webApplicationUrl = string.Empty;
                        try
                        {
                            Logger.LogInfoMessage("Processing the site: " + siteCollection, true);
                            webApplicationUrl = GetWebapplicationUrlFromSiteCollectionUrl(siteCollection);

                            //Record SiteCollection Url in SiteCollections.txt
                            ProcessSiteCollectionUrl(siteCollection, ref lstMissingSiteTempaltesInGalleryBase, webApplicationUrl);
                        }
                        catch (Exception ex)
                        {
                            Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessSiteTemplateInputFile]. Exception Message: " + ex.Message + ", Exception Comments: ", true);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessSiteTemplateInputFile]. Exception Message: " + ex.Message + ", Exception Comments: ", true);
                }
            }
            else
            {
                Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessSiteTemplateInputFile]. Exception Message: Site Template: " + filePath + @"\" + Constants.SiteTemplateCustomizationUsage + " is not present", true);
            }
        }

        public static void ProcessWebApplicationUrl(string webApplicationUrl, ref List<SiteTemplateFTCAnalysisOutputBase> lstMissingSiteTempaltesInGalleryBase)
        {
            try
            {
                //Delete SiteCollections.txt file if it already exists
                if (System.IO.File.Exists(outputPath + @"\" + Constants.SiteCollectionsTextFile))
                    System.IO.File.Delete(outputPath + @"\" + Constants.SiteCollectionsTextFile);
                //Create SiteCollections.txt file
                System.IO.StreamWriter file = new System.IO.StreamWriter(outputPath + @"\" + Constants.SiteCollectionsTextFile);

                List<SiteEntity> siteUrls = GenerateSiteCollectionReport.GetAllSites(webApplicationUrl);
                if (siteUrls != null && siteUrls.Count > 0)
                {
                    foreach (SiteEntity siteUrlEntity in siteUrls)
                    {
                        try
                        {
                            string siteCollection = siteUrlEntity.Url;
                            Logger.LogInfoMessage("Processing the site: " + siteUrlEntity.Url, true);
                            //Record SiteCollection Url in SiteCollections.txt
                            file.WriteLine(siteUrlEntity.Url);
                            ProcessSiteCollectionUrl(siteCollection, ref lstMissingSiteTempaltesInGalleryBase, webApplicationUrl);
                        }
                        catch (Exception ex)
                        {
                            Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWebApplicationUrl]. Exception Message: " + ex.Message + ", Exception Comments: ", true);
                        }
                    }
                }
                file.Close();
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessWebApplicationUrl]. Exception Message: " + ex.Message + ", Exception Comments: ", true);
            }
        }

        public static void ProcessSiteCollectionUrlsList(string[] siteCollectionUrls, ref List<SiteTemplateFTCAnalysisOutputBase> lstMissingSiteTempaltesInGalleryBase)
        {
            try
            {
                foreach (string siteCollectionUrl in siteCollectionUrls)
                {
                    string siteCollection = siteCollectionUrl.Trim();
                    string webApplicationUrl = string.Empty;
                    try
                    {
                        Logger.LogInfoMessage("Processing the site: " + siteCollection, true);
                        webApplicationUrl = GetWebapplicationUrlFromSiteCollectionUrl(siteCollection);
                        ProcessSiteCollectionUrl(siteCollection, ref lstMissingSiteTempaltesInGalleryBase, webApplicationUrl);
                    }
                    catch (Exception ex)
                    {
                        Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessSiteCollectionUrlsList]. Exception Message: " + ex.Message + ", Exception Comments: ", true);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DownloadAndModifySiteTemplate: ProcessSiteCollectionUrlsList]. Exception Message: " + ex.Message + ", Exception Comments: ", true);
            }
        }

        public static string GetWebapplicationUrlFromSiteCollectionUrl(string siteCollection)
        {
            Uri uri;
            try
            {
                uri = new Uri(siteCollection);
                if (uri != null)
                {
                    return uri.Scheme + @"://" + uri.Host;
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[DownloadAndModifyListTemplate: GetWebapplicationUrlFromSiteCollectionUrl]. Exception Message: " + ex.Message + ", Exception Comments: ", true);
            }
            finally
            {
                uri = null;
            }
            return string.Empty;
        }
    }
}
