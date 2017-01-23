using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.IO;
using System.Xml;

using Microsoft.SharePoint.Client;
using UdcxRemediation.Console.Common.Base;
using UdcxRemediation.Console.Common.CSV;

namespace UdcxRemediation.Console
{
    public class CommentUDCXFileNodes
    {
        public static void DoWork(string inputFileSpec)
        {
            Logger.OpenLog("CommentUDCXFileNodes");
            Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);
            Logger.LogInfoMessage(inputFileSpec, true);

            List<UdcxReportOutput> _WriteUDCList = null;

            Logger.LogInfoMessage(String.Format("AppSettings:"), true);
            Logger.LogInfoMessage(String.Format("- AppSettings[UseAppModel] = {0}", Program.UseAppModel), true);
            if (Program.UseAppModel == true)
            {
                Logger.LogInfoMessage(String.Format("- AppId = {0}", ConfigurationManager.AppSettings["ClientId"].ToString()), true);
            }
            else
            {
                string adminUsername = String.Format("{0}{1}", (String.IsNullOrEmpty(Program.AdminDomain) ? "" : String.Format("{0}\\", Program.AdminDomain)), Program.AdminUsername);
                Logger.LogInfoMessage(String.Format("- Admin Username = {0}", adminUsername), true);
            }

            IEnumerable<UdcxReportInput> udcxCSVRows = ImportCSV.ReadMatchingColumns<UdcxReportInput>(inputFileSpec, Constants.CsvDelimeter);
            if (udcxCSVRows != null)
            {
                try
                {
                    var authRows = udcxCSVRows.Where(x => x.Authentication != null && x.Authentication != Constants.ErrorStatus && x.Authentication.Length > 0);
                    if (authRows != null && authRows.Count() > 0)
                    {
                        _WriteUDCList = new List<UdcxReportOutput>();
                        Logger.LogInfoMessage(String.Format("Preparing to process a total of {0} Udcx Files ...", authRows.Count()), true);

                        foreach (UdcxReportInput udcxFileInput in authRows)
                        {
                            CommentUDCXFileNode(udcxFileInput, _WriteUDCList);
                        }

                        GenerateStatusReport(_WriteUDCList);
                    }
                    else
                    {
                        Logger.LogInfoMessage("No UDCX File records with authentication nodes found in '" + inputFileSpec + "' File ", true);
                    }                  
                }
                catch (Exception ex)
                {
                    Logger.LogErrorMessage(String.Format("CommentUDCXFileNode() failed: Error={0}", ex.Message), true);
                }

                Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            }
            else
            {
                Logger.LogInfoMessage("No UDCX File records found in '" + inputFileSpec + "' File ", true);
            }

            Logger.CloseLog();
        }

        private static void CommentUDCXFileNode(UdcxReportInput udcxFileInput, List<UdcxReportOutput> _WriteUDCList)
        {
            if (udcxFileInput == null)
            {
                return;
            }

            string siteUrl = udcxFileInput.SiteUrl;
            string webUrl = udcxFileInput.WebUrl;
            string dirName = udcxFileInput.DirName;
            string leafName = udcxFileInput.LeafName;
            string authentication = udcxFileInput.Authentication;

            UdcxReportOutput udcxOutput = new UdcxReportOutput();
            udcxOutput.SiteUrl = siteUrl;
            udcxOutput.WebUrl = webUrl;
            udcxOutput.DirName = dirName;
            udcxOutput.LeafName = leafName;
            udcxOutput.Authentication = authentication;

            if (dirName.EndsWith("/"))
            {
                dirName = dirName.TrimEnd(new char[] { '/' });
            }
            if (leafName.StartsWith("/"))
            {
                leafName = leafName.TrimStart(new char[] { '/' });
            }
            string serverRelativeFolderPath = "/" + dirName;
            string serverRelativeFilePath = "/" + dirName + '/' + leafName;

            try
            {
                Logger.LogInfoMessage(String.Format("Processing UCDX File [{0}/{1}] of Web [{2}] ...", dirName, leafName, webUrl), true);

                // IMPORTANT: Open the webUrl, not the siteUrl, so Folder.Files.Add() can properly process files of child webs
                using (ClientContext userContext = Helper.CreateClientContextBasedOnAuthMode(Program.UseAppModel, Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    Web web = userContext.Web;
                    userContext.Load(web);
                    userContext.ExecuteQueryRetry();

                    XNamespace xns = "http://schemas.microsoft.com/office/infopath/2006/udc";
                    XDocument xmlDoc = null;

                    Logger.LogInfoMessage(String.Format("Getting contents of UCDX File [{0}] from Web [{1}] ...", serverRelativeFilePath, webUrl), false);
                    // Approach to read File contents depends on Auth Model chosen
                    if (Program.UseAppModel == true)
                    {
                        string originalFileContents = SafeGetFileAsString(web, serverRelativeFilePath);
                        if (String.IsNullOrEmpty(originalFileContents))
                        {
                            Logger.LogErrorMessage(String.Format("Could not get contents of UCDX File"), false);
                            udcxOutput.Status = Constants.ErrorStatus + ": could not get file contents.";
                            _WriteUDCList.Add(udcxOutput);
                            return;
                        }

                        xmlDoc = XDocument.Load(new StringReader(originalFileContents));
                    }
                    else
                    {
                        FileInformation info = Microsoft.SharePoint.Client.File.OpenBinaryDirect(userContext, serverRelativeFilePath);
                        xmlDoc = XDocument.Load(XmlReader.Create(info.Stream));
                    }
                    Logger.LogInfoMessage(String.Format("Got contents of UCDX File"), false);

                    XElement authElem = xmlDoc.Root.Element(xns + "ConnectionInfo").Element(xns + "Authentication");
                    if (authElem != null)
                    {
                        string authData = authElem.ToString();
                        authData = authData.Replace("<udc:Authentication xmlns:udc=\"" + xns + "\">", "<udc:Authentication>");
                        authElem.ReplaceWith(new XComment(authData));

                        string saveUdcxContent = xmlDoc.Declaration.ToString() + xmlDoc.ToString();

                        using (MemoryStream contentStream = new MemoryStream())
                        {
                            StreamWriter writer = new StreamWriter(contentStream);
                            writer.Write(saveUdcxContent);
                            writer.Flush();
                            contentStream.Position = 0;

                            Logger.LogInfoMessage(String.Format("Saving contents of UCDX File [{0}] to Web [{1}] ...", serverRelativeFilePath, webUrl), false);
                            Folder targetFolder = null;

                            // grab the parent folder in preparation for the file upload...
                            Logger.LogInfoMessage(String.Format("Getting folder [{0}] of Web [{1}] ...", serverRelativeFolderPath, webUrl), false);
                            try
                            {
                                targetFolder = web.GetFolderByServerRelativeUrl(serverRelativeFolderPath);
                                userContext.Load(targetFolder);
                                userContext.ExecuteQueryRetry();

                                Logger.LogInfoMessage(String.Format("Got folder"), false);
                            }
                            catch (Exception ex)
                            {
                                Logger.LogErrorMessage(String.Format("CommentUDCXFileNode() failed for UDCX File [{0}/{1}] of Web [{2}]: Reason={3}; Error={4}", dirName, leafName, webUrl,
                                    "Upload Folder was not Found.",
                                    "[" + ex.Message + "] | [" + ex.HResult + "] | [" + ex.Source + "] | [" + ex.StackTrace + "] | [" + ex.TargetSite + "]"), false);
                                udcxOutput.Status = Constants.ErrorStatus + ": Upload Folder was not Found.";
                                _WriteUDCList.Add(udcxOutput);
                                return;
                            }

                            // check-out the file (if needed) in preparation for the file upload...
                            try
                            {
                                Logger.LogInfoMessage(String.Format("Checking out file [{0}] ...", leafName), false);
                                web.CheckOutFile(serverRelativeFilePath);
                                Logger.LogInfoMessage(String.Format("Checked out file"), false);
                            }
                            catch (Exception ex)
                            {
                                Logger.LogErrorMessage(String.Format("CommentUDCXFileNode() failed for UDCX File [{0}/{1}] of Web [{2}]: Reason={3}; Error={4}", dirName, leafName, webUrl,
                                    "File Checkout failed.",
                                    "[" + ex.Message + "] | [" + ex.HResult + "] | [" + ex.Source + "] | [" + ex.StackTrace + "] | [" + ex.TargetSite + "]"), false);
                                udcxOutput.Status = Constants.ErrorStatus + ": File Checkout failed.";
                                _WriteUDCList.Add(udcxOutput);
                                return;
                            }

                            // upload the modified file...
                            Logger.LogInfoMessage(String.Format("Uploading file [{0}] ...", leafName), false);
                            Microsoft.SharePoint.Client.File targetFile = null;

                            // Approach to save File contents depends on Auth Model chosen
                            if (Program.UseAppModel == true)
                            {
                                try
                                {
                                    targetFile = targetFolder.UploadFile(leafName, contentStream, true);
                                    Logger.LogInfoMessage(String.Format("Uploaded file"), false);
                                }
                                catch (Exception ex)
                                {
                                    Logger.LogErrorMessage(String.Format("CommentUDCXFileNode() failed for UDCX File [{0}/{1}] of Web [{2}]: Reason={3}; Error={4}", dirName, leafName, webUrl,
                                        "File Upload failed.",
                                        "[" + ex.Message + "] | [" + ex.HResult + "] | [" + ex.Source + "] | [" + ex.StackTrace + "] | [" + ex.TargetSite + "]"), false);
                                    udcxOutput.Status = Constants.ErrorStatus + ": File Upload failed.";
                                    _WriteUDCList.Add(udcxOutput);
                                    return;
                                }
                            }
                            else
                            {
                                try
                                {
                                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(userContext, serverRelativeFilePath, contentStream, true);

                                    targetFile = web.GetFileByServerRelativeUrl(serverRelativeFilePath);
                                    web.Context.Load(targetFile);
                                    web.Context.ExecuteQueryRetry();
                                }
                                catch (Exception ex)
                                {
                                    Logger.LogErrorMessage(String.Format("CommentUDCXFileNode() failed for UDCX File [{0}/{1}] of Web [{2}]: Reason={3}; Error={4}", dirName, leafName, webUrl,
                                        "File Upload failed.",
                                        "[" + ex.Message + "] | [" + ex.HResult + "] | [" + ex.Source + "] | [" + ex.StackTrace + "] | [" + ex.TargetSite + "]"), false);
                                    udcxOutput.Status = Constants.ErrorStatus + ": File Upload failed.";
                                    _WriteUDCList.Add(udcxOutput);
                                    return;
                                }
                            }

                            // publish the modified file (executes check-in, publish, and approval as needed)...
                            try
                            {
                                Logger.LogInfoMessage(String.Format("Publishing file [{0}] ...", leafName), false);
                                targetFile.PublishFileToLevel(FileLevel.Published);
                                Logger.LogInfoMessage(String.Format("Published file"), false);
                            }
                            catch (Exception ex)
                            {
                                Logger.LogErrorMessage(String.Format("CommentUDCXFileNode() failed for UDCX File [{0}/{1}] of Web [{2}]: Reason={3}; Error={4}", dirName, leafName, webUrl,
                                    "File Publish failed.",
                                    "[" + ex.Message + "] | [" + ex.HResult + "] | [" + ex.Source + "] | [" + ex.StackTrace + "] | [" + ex.TargetSite + "]"), false);
                                udcxOutput.Status = Constants.ErrorStatus + ": File Publish failed.";
                                _WriteUDCList.Add(udcxOutput);
                                return;
                            }
                            Logger.LogInfoMessage(String.Format("Saved contents of UCDX File [{0}] to Web [{1}]", serverRelativeFilePath, webUrl), false);

                            udcxOutput.Status = Constants.SuccessStatus;
                            Logger.LogSuccessMessage(String.Format("Updated UCDX File [{0}/{1}] of Web [{2}]", dirName, leafName, webUrl), false);
                        }
                    }
                    else
                    {
                        udcxOutput.Status = Constants.NoAuthNodeFound;
                        Logger.LogWarningMessage(String.Format("Skipped UCDX File [{0}/{1}] of Web [{2}]: Reason={3}", dirName, leafName, webUrl, Constants.NoAuthNodeFound), false);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("CommentUDCXFileNode() failed for UDCX File [{0}/{1}] of Web [{2}]: Error={3}", dirName, leafName, webUrl, ex.Message), false);
                udcxOutput.Status = Constants.ErrorStatus;
                udcxOutput.ErrorDetails = ex.Message;
            }

            _WriteUDCList.Add(udcxOutput);
        }

        private static void GenerateStatusReport(List<UdcxReportOutput> _WriteUDCList)
        {
            string reportFileName = Environment.CurrentDirectory + "\\" + Constants.UdcxReport;

            Common.Utilities.FileUtility.WriteCsVintoFile(reportFileName, ref _WriteUDCList);
        }

        private static string SafeGetFileAsString(Web web, string serverRelativeFilePath)
        {
            try
            {
                return web.GetFileAsString(serverRelativeFilePath);
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("SafeGetFileAsString() failed for File [{0}] of Web [{1}]: Error={2}", serverRelativeFilePath, web.Url, ex.Message), false);
                return String.Empty;
            }
        }

    }
}
