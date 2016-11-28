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
    public class ReplaceWebPart
    {
        public static string filePath = string.Empty;
        public static string outputPath = string.Empty;
        public static string timeStamp = string.Empty;

        public static void DoWork()
        {
            timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");
            outputPath = Environment.CurrentDirectory;
            string webPartsInputFile = string.Empty;
            string webpartType = string.Empty;
            string targetWebPartFileName = string.Empty;
            string targetWebPartXmlFilePath = string.Empty;


            //Trace Log TXT File Creation Command
            Logger.OpenLog("ReplaceWebPart", timeStamp);

            if (!ReadInputFile(ref webPartsInputFile))
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("Webparts input file is not valid or available. So, Operation aborted!");
                Logger.LogErrorMessage("Please enter path like: E.g. C:\\<Working Directory>\\<InputFile>.csv");
                System.Console.ResetColor();
                return;
            }

            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Please enter Source Webpart Type :");
            System.Console.ResetColor();
            webpartType = System.Console.ReadLine().ToLower();
            if (string.IsNullOrEmpty(webpartType))
            {
                Logger.LogErrorMessage("[ReplaceWebPart: DoWork]Webpart Type should not be empty or null. Operation aborted...", true);
                return;
            }

            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Please enter Target WebPart File Name :");
            System.Console.ResetColor();
            targetWebPartFileName = System.Console.ReadLine().ToLower();
            if (string.IsNullOrEmpty(targetWebPartFileName))
            {
                Logger.LogErrorMessage("[ReplaceWebPart: DoWork]Target WebPart File Name should not be empty or null. Operation aborted...", true);
                return;
            }

            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Please enter Target WebPart Xml File Path :");
            System.Console.ResetColor();
            targetWebPartXmlFilePath = System.Console.ReadLine().ToLower();

            if (string.IsNullOrEmpty(targetWebPartXmlFilePath) || !System.IO.File.Exists(targetWebPartXmlFilePath))
            {
                Logger.LogErrorMessage("[ReplaceWebPart: DoWork]Target WebPart Xml File Path is not valid or available. Operation aborted...", true);
                return;
            }
            Logger.LogInfoMessage(String.Format("Process started {0}", DateTime.Now.ToString()), true);
            try
            {
                TransformWebPart_UsingCSV(webPartsInputFile, webpartType, targetWebPartFileName, targetWebPartXmlFilePath, outputPath);
            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[ReplaceWebPart: DoWork]. Exception Message: " + ex.Message, true);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "ReplaceWebPart", ex.Message, ex.ToString(), "ReplaceWebPart: DoWork()", ex.GetType().ToString());
            }
            Logger.LogInfoMessage(String.Format("Process completed {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }
        public static void TransformWebPart_UsingCSV(string usageFileName, string sourceWebPartType, string targetWebPartFileName, string targetWebPartXmlFilePath, string outPutDirectory)
        {
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                WebPart_Initialization(outputPath);

                string sourceWebPartXmlFilesDir = outPutDirectory + @"\" + Constants.SOURCE_WEBPART_XML_DIR;

                if (!System.IO.Directory.Exists(sourceWebPartXmlFilesDir))
                {
                    System.IO.Directory.CreateDirectory(sourceWebPartXmlFilesDir);
                }


                //ReplaceWebPart_UsingCSV
                string targetWebPartXmlsDir = outPutDirectory + @"\" + Constants.TARGET_WEBPART_XML_DIR;
                ReplaceWebPart_UsingCSV(sourceWebPartType, targetWebPartXmlFilePath, targetWebPartFileName, targetWebPartXmlsDir, usageFileName, outPutDirectory);
            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[TransformWebPart_UsingCSV] Exception Message: " + ex.Message);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "ReplaceWebPart", ex.Message, ex.ToString(), "TransformWebPart_UsingCSV()", ex.GetType().ToString(), exceptionCommentsInfo1);
            }
        }

        public static bool ReadInputFile(ref string webPartsInputFile)
        {
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            Logger.LogMessage("Enter Complete Input File Path of Webparts Report Either Pre-Scan OR Discovery Report:");
            System.Console.ResetColor();
            webPartsInputFile = System.Console.ReadLine();
            System.Console.WriteLine("[ReadInputFile] Entered Input File of Webpart " + webPartsInputFile, false);
            if (string.IsNullOrEmpty(webPartsInputFile) || !System.IO.File.Exists(webPartsInputFile))
                return false;
            return true;
        }
        public static void ReplaceWebPart_UsingCSV(string sourceWebPartType, string targetWebPartXmlFilePath, string targetWebPartFileName, string targetWebPartXmlDir, string usageFileName, string outPutFolder)
        {
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                //Reading Input File
                IEnumerable<WebPartDiscoveryInput> objWPDInput;
                ReadWebPartUsageCSV(sourceWebPartType, usageFileName, outPutFolder, out objWPDInput);

                bool headerTransformWebPart = false;

                if (objWPDInput.Any())
                {
                    //bool headerPageLayout = false;

                    for (int i = 0; i < objWPDInput.Count(); i++)
                    {
                        try
                        {
                            WebPartDiscoveryInput objInput = objWPDInput.ElementAt(i);
                            SetExceptionCSVWebAppSiteColWebUrls(objInput.WebUrl.ToString());

                            //This is for Exception Comments:
                            exceptionCommentsInfo1 = "WebUrl: " + objInput.WebUrl + ", ZoneID: " + objInput.ZoneID + ", Web PartID:" + objInput.WebPartId.ToString() + " ,PageUrl: " + objInput.PageUrl.ToString();
                            //This is for Exception Comments:

                            //This function is Get Relative URL of the page
                            string _relativePageUrl = string.Empty;
                            _relativePageUrl = GetPageRelativeURL(objInput.WebUrl.ToString(), objInput.PageUrl.ToString());

                            string _storageKey = string.Empty;
                            _storageKey = GetWebPartID(objInput.StorageKey);

                            //sourceWebPartId - Used to update the content of the wikipage with new web part id [Fix for BugId - 95007] 
                            string _webPartId = string.Empty;
                            _webPartId = GetWebPartID(objInput.WebPartId);
                            //End
                            bool status = false;

                            string sourceWebPartXmlFilePath = WebPartProperties.GetWebPartProperties(_relativePageUrl, _storageKey, objInput.WebUrl, outPutFolder);

                            if (!string.IsNullOrEmpty(sourceWebPartXmlFilePath))
                            {
                                ConfigureNewWebPartXmlFile(targetWebPartXmlFilePath, sourceWebPartXmlFilePath, outPutFolder);

                                string _targetWebPartXml = string.Empty;
                                _targetWebPartXml = GetTargetWebPartXmlFilePath(_storageKey, targetWebPartXmlDir);

                                if (_targetWebPartXml != "")
                                {
                                    status = ReplaceWebPartInPage(objInput.WebUrl, targetWebPartFileName, _targetWebPartXml, _storageKey, objInput.ZoneIndex, objInput.ZoneID, _relativePageUrl, outPutFolder, _webPartId);
                                }
                                else
                                {
                                    System.Console.ForegroundColor = System.ConsoleColor.Red;
                                    Logger.LogErrorMessage("[ReplaceWebPart_UsingCSV] Target WebPartXml File: " + targetWebPartFileName + " does not exists for StorageKey" + _storageKey + "in Path(TargetWebPartXmlDir): " + targetWebPartXmlDir + " for page " + _relativePageUrl);
                                    System.Console.ResetColor();
                                }
                            }
                            else
                            {
                                System.Console.ForegroundColor = ConsoleColor.Red;
                                Logger.LogErrorMessage("[ReplaceWebPart_UsingCSV] Failed to get Source WebPart Properties: on Web: " + objInput.WebUrl + " for StorageKey " + _storageKey + "for page " + _relativePageUrl);
                                System.Console.ResetColor();
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
                            objWPOutputBase.ExecutionDateTime = DateTime.Now.ToString();

                            if (status)
                            {
                                objWPOutputBase.Status = Constants.Success;
                            }
                            else
                            {
                                objWPOutputBase.Status = Constants.Failure;
                            }
                            if (!System.IO.File.Exists(outputPath + @"\" + Constants.ReplaceWebPartStatusFileName + timeStamp + Constants.CSVExtension))
                            {
                                headerTransformWebPart = false;
                            }
                            else
                                headerTransformWebPart = true;

                            FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.ReplaceWebPartStatusFileName + timeStamp + Constants.CSVExtension, objWPOutputBase, ref headerTransformWebPart);

                        }
                        catch (Exception ex)
                        {
                            System.Console.ForegroundColor = ConsoleColor.Red;
                            Logger.LogErrorMessage("Error in Processing Web:" + ExceptionCsv.WebUrl + " `\r\nError Details:" + ex.Message + " `\r\nExceptionComments:" + exceptionCommentsInfo1);
                            System.Console.ResetColor();
                            ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "ReplaceWebPart", ex.Message, ex.ToString(), "ReplaceWebPart_UsingCSV()", ex.GetType().ToString(), exceptionCommentsInfo1);
                        }
                    }

                }
                else
                {
                    Logger.LogInfoMessage("Source WebPart Type: " + sourceWebPartType + " does not exist in the Input file: " + usageFileName, true);
                }

            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[ReplaceWebPart_UsingCSV] Exception Message: " + ex.Message);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "ReplaceWebPart", ex.Message, ex.ToString(), "ReplaceWebPart_UsingCSV()", ex.GetType().ToString(), exceptionCommentsInfo1);
            }
        }

        private static void ReadWebPartUsageCSV(string sourceWebPartType, string usageFilePath, string outPutFolder, out IEnumerable<WebPartDiscoveryInput> objWPDInput)
        {
            string exceptionCommentsInfo1 = string.Empty;

            Logger.LogInfoMessage("[ReadWebPartUsageCSV] [START] Calling function ImportCsv.ReadMatchingColumns<WebPartDiscoveryInput>");

            objWPDInput = null;
            objWPDInput = ImportCSV.ReadMatchingColumns<WebPartDiscoveryInput>(usageFilePath, Constants.CsvDelimeter);

            Logger.LogInfoMessage("[ReadWebPartUsageCSV] [END] Read all the WebParts Usage Details from Discovery Usage File and saved in List - out IEnumerable<WebPartDiscoveryInput> objWPDInput, for processing.");

            try
            {
                if (objWPDInput.Any())
                {
                    Logger.LogInfoMessage("[START] ReadWebPartUsageCSV - After Loading InputCSV ");

                    objWPDInput = from p in objWPDInput
                                  where p.WebPartType.ToLower() == sourceWebPartType.ToLower()
                                  select p;
                    exceptionCommentsInfo1 = objWPDInput.ToString();

                    Logger.LogInfoMessage("[END] ReadWebPartUsageCSV - After Loading InputCSV");
                }
            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[ReadWebPartUsageCSV] Exception Message: " + ex.Message + ", Exception Comments:" + exceptionCommentsInfo1);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "ReplaceWebPart", ex.Message, ex.ToString(), "ReadWebPartUsageCSV()", ex.GetType().ToString(), exceptionCommentsInfo1);
            }
        }

        private static string GetPageRelativeURL(string WebUrl, string PageUrl)
        {
            ClientContext clientContext = null;
            string _relativePageUrl = string.Empty;
            Web web = null;
            try
            {
                if (WebUrl != "" || PageUrl != "")
                {

                    using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, WebUrl))
                    {
                        web = userContext.Web;
                        userContext.Load(web);
                        userContext.ExecuteQuery();
                        clientContext = userContext;

                        Logger.LogInfoMessage("[GetPageRelativeURL] Web.ServerRelativeUrl: " + web.ServerRelativeUrl + " And PageUrl: " + PageUrl);

                        //Issue: Found in MARS Retraction Process, the root web ServerRelativeUrl would result "/" only
                        //Hence appending "/" would throw exception for ServerRelativeUrl parameter
                        if (web.ServerRelativeUrl.ToString().Equals("/"))
                        {
                            _relativePageUrl = web.ServerRelativeUrl.ToString() + PageUrl;
                        }
                        else if (!PageUrl.Contains(web.ServerRelativeUrl))
                        {
                            _relativePageUrl = web.ServerRelativeUrl.ToString() + "/" + PageUrl;
                        }
                        else
                        {
                            _relativePageUrl = PageUrl;
                        }
                        Logger.LogInfoMessage("[GetPageRelativeURL] RelativePageUrl Framed: " + _relativePageUrl);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[GetPageRelativeURL] Exception Message: " + ex.Message);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, WebUrl, "ReplaceWebPart", ex.Message, ex.ToString(), "GetPageRelativeURL()", ex.GetType().ToString());
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
                Logger.LogErrorMessage("[GetWebPartID]. Exception Message: " + ex.Message, true);
            }
            return _webPartID;
        }
        public static void ConfigureNewWebPartXmlFile(string targetWebPartXmlFilePath, string sourceXmlFile, string OutPutDirectory)
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

            Logger.LogInfoMessage("[ConfigureNewWebPartXmlFile] Configuring target webpart with source web part");

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
                            Logger.LogInfoMessage("[ConfigureNewWebPartXmlFile] Property:" + oOTBWBProperty.name + " matched in both " + System.IO.Path.GetFileNameWithoutExtension(sourceWebPartXmlFilePath) + " and " + System.IO.Path.GetFileNameWithoutExtension(targetWebPartXmlFilePath));
                            //Console.WriteLine("[ConfigureNewWebPartXmlFile] Property:" + oOTBWBProperty.name + " matched in both " + System.IO.Path.GetFileNameWithoutExtension(sourceWebPartXmlFilePath) + " and " + System.IO.Path.GetFileNameWithoutExtension(targetWebPartXmlFilePath));
                            break;
                        }
                    }
                    if (!isUpdatePoperty)
                    {
                        notUpdatedProperties.data.properties.SetValue(customeWBProperty, i);
                        Logger.LogInfoMessage("[ConfigureNewWebPartXmlFile] Property:" + customeWBProperty.name + "doesn't matched in both " + System.IO.Path.GetFileNameWithoutExtension(sourceWebPartXmlFilePath) + " and " + System.IO.Path.GetFileNameWithoutExtension(targetWebPartXmlFilePath));
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
                Logger.LogInfoMessage("[ConfigureNewWebPartXmlFile] New Configured web part is created at " + newWebPartXmlFilePath);

                Logger.LogInfoMessage("[ConfigureNewWebPartXmlFile] The properties which are not configured are extracted into a new file at " + notUpdatedPropertiesXmlFilePath);

                string result = newWebPartXmlFilePath + ";" + notUpdatedPropertiesXmlFilePath;
                string[] filePaths = result.Split(';');
                if (filePaths.Count() > 1)
                {
                    if (!String.IsNullOrEmpty(filePaths[0].Trim()))
                    {
                        Logger.LogInfoMessage("[ConfigureNewWebPartXmlFile] New Configured Web Part Xml File:" + filePaths[0]);
                    }
                    if (!String.IsNullOrEmpty(filePaths[1].Trim()))
                    {
                        Logger.LogInfoMessage("[ConfigureNewWebPartXmlFile] Custom Properties not configured in the new web part xml file:" + filePaths[1]);
                    }
                }
                Logger.LogSuccessMessage("[ConfigureNewWebPartXmlFile] Successfully configured web part");


            }
            catch (Exception ex)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("[ConfigureNewWebPartXmlFile] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                System.Console.ResetColor();
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "ReplaceWebPart", ex.Message, ex.ToString(), "ConfigureNewWebPartXmlFile()", ex.GetType().ToString());
            }
        }

        private static string GetTargetWebPartXmlFilePath(string webPartId, string targetDirectory)
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
        public static bool ReplaceWebPartInPage(string webUrl, string targetWebPartFileName, string targetWebPartXmlFile, string sourceWebPartStorageKey, string webPartZoneIndex, string webPartZoneID, string serverRelativePageUrl, string outPutDirectory, string sourceWebPartId = "")
        {
            bool isWebPartReplaced = false;

            if (DeleteWebparts.DeleteWebPart(webUrl, serverRelativePageUrl, sourceWebPartStorageKey))
            {
                System.Console.ForegroundColor = System.ConsoleColor.Green;
                Logger.LogSuccessMessage("[ReplaceWebPart]Successfully Deleted WebPart ");
                System.Console.ResetColor();
                if (AddWebPart.AddWebPartToPage(webUrl, targetWebPartFileName, targetWebPartXmlFile, webPartZoneIndex, webPartZoneID, serverRelativePageUrl, outPutDirectory, sourceWebPartId))
                {
                    System.Console.ForegroundColor = System.ConsoleColor.Green;
                    isWebPartReplaced = true;
                    Logger.LogSuccessMessage("[ReplaceWebPart]Successfully Added WebPart ");
                    Logger.LogSuccessMessage("[ReplaceWebPart] Successfully Replaced the newly configured WebPart and output file is present in the path: " + outPutDirectory, true);
                    System.Console.ResetColor();
                }
            }

            return isWebPartReplaced;
        }
        public static void SetExceptionCSVWebAppSiteColWebUrls(string webUrl)
        {
            ClientContext clientContext = null;
            Web web = null;

            using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
            {
                web = userContext.Web;
                userContext.Load(web, w => w.Url,
                                        w => w.ServerRelativeUrl);
                userContext.ExecuteQuery();
                clientContext = userContext;

                Site site = clientContext.Site;
                clientContext.Load(site);
                clientContext.ExecuteQuery();

                ExceptionCsv.WebApplication = web.Url.Replace(web.ServerRelativeUrl, "");
                ExceptionCsv.SiteCollection = site.Url;
                ExceptionCsv.WebUrl = web.Url;

            }

        }

        public static void WebPart_Initialization(string DiscoveryUsage_OutPutFolder)
        {
            //Excception CSV Creation Command
            ExceptionCsv objException = ExceptionCsv.CurrentInstance;
            objException.CreateLogFile(DiscoveryUsage_OutPutFolder);

            ExceptionCsv.WebApplication = Constants.NotApplicable;
            ExceptionCsv.SiteCollection = Constants.NotApplicable;
            ExceptionCsv.WebUrl = Constants.NotApplicable;

        }

    }
}
