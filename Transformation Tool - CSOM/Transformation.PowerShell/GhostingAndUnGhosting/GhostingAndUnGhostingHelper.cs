using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Transformation.PowerShell.Common;
using Transformation.PowerShell.Common.Utilities;

namespace Transformation.PowerShell.GhostingAndUnGhosting
{
    public class GhostingAndUnGhostingHelper
    {
        private void GhostingAndUnGhosting_Initialization(string OutPutFolder, string Type)
        {
            //Excception CSV Creation Command
            ExceptionCsv objException = ExceptionCsv.CurrentInstance;
            objException.CreateLogFile(OutPutFolder);

            ExceptionCsv.WebApplication = Constants.NotApplicable;
            ExceptionCsv.SiteCollection = Constants.NotApplicable;
            ExceptionCsv.WebUrl = Constants.NotApplicable;

            //Trace Log TXT File Creation Command
            Logger objTraceLogs = Logger.CurrentInstance;
            objTraceLogs.CreateLogFile(OutPutFolder);
            //Trace Log TXT File Creation Command

            //Delete Output Files
            if (Type.ToUpper() == "UNGHOST")
            {
                FileUtility.DeleteFiles(OutPutFolder + @"\" + Constants.UnGhosting_Output);
            }
            else if (Type.ToUpper() == "DOWNLOAD")
            {
                FileUtility.DeleteFiles(OutPutFolder + @"\" + Constants.UnGhosting_DownloadFileReport);
            }
        }
        public void UnGhostFile(string absoluteFilePath, string outPutFolder, string OperationType, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "NA", string Password = "NA", string Domain = "NA")
        {
            string fileName = string.Empty;
            string newFileName = string.Empty;
            string directoryName = string.Empty;
            bool headerCSVColumns = false;
            string exceptionCommentsInfo1 = string.Empty;

            GhostingAndUnGhosting_Initialization(outPutFolder, "UNGHOST");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Un Ghosting - Trasnformation Utility Execution Started - For Web ##############");
            Console.WriteLine("############## Un Ghosting - Trasnformation Utility Execution Started - For Web ##############");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
            Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: UnGhostFile");
            Console.WriteLine("[START] ::: UnGhostFile");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[UnGhostFile] Initiated Logger and Exception Class. Logger and Exception file will be available in path " + outPutFolder);
            Console.WriteLine("[UnGhostFile] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);

            try
            {
                exceptionCommentsInfo1 = "FilePath: " + absoluteFilePath;
                AuthenticationHelper ObjAuth = new AuthenticationHelper();
                ClientContext clientContext = null;

                Uri fileUrl = new Uri(absoluteFilePath);

                clientContext = new ClientContext(fileUrl.GetComponents(UriComponents.SchemeAndServer, UriFormat.UriEscaped));
                Uri siteUrl = Web.WebUrlFromPageUrlDirect(clientContext, fileUrl);
                clientContext = new ClientContext(siteUrl);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[UnGhostFile] WebUrl is " + siteUrl.ToString());
                Console.WriteLine("[UnGhostFile] WebUrl is " + siteUrl.ToString());

                ExceptionCsv.WebUrl = siteUrl.ToString();

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.OnPremise)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UnGhostFile] GetNetworkCredentialAuthenticatedContext for WebUrl: " + siteUrl.ToString());
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(siteUrl.ToString(), UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UnGhostFile] GetNetworkCredentialAuthenticatedContext for WebUrl: " + siteUrl.ToString());
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.Online)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][UnGhostFile] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + siteUrl.ToString());
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(siteUrl.ToString(), UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][UnGhostFile] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + siteUrl.ToString());
                }

                if (clientContext != null)
                {
                    Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByServerRelativeUrl(fileUrl.AbsolutePath);
                    clientContext.Load(file);
                    clientContext.ExecuteQuery();
                    directoryName = GetLibraryName(fileUrl.ToString(), siteUrl.ToString(), fileName);

                    Folder folder = clientContext.Web.GetFolderByServerRelativeUrl(directoryName);
                    clientContext.Load(folder);
                    clientContext.ExecuteQuery();

                    fileName = file.Name;
                    newFileName = GetNextFileName(fileName);
                    string path = System.IO.Directory.GetCurrentDirectory();
                    string downloadedFilePath = path + "\\" + newFileName;

                    using (WebClient myWebClient = new WebClient())
                    {
                        myWebClient.Credentials = CredentialCache.DefaultNetworkCredentials;
                        myWebClient.DownloadFile(absoluteFilePath, downloadedFilePath);
                    }

                    Microsoft.SharePoint.Client.File uploadedFile = FileFolderExtensions.UploadFile(folder, newFileName, downloadedFilePath, true);
                    if (uploadedFile.CheckOutType.Equals(CheckOutType.Online))
                    {
                        uploadedFile.CheckIn("File is UnGhotsed and Updated", CheckinType.MinorCheckIn);                        
                    }
                    clientContext.Load(uploadedFile);
                    clientContext.ExecuteQuery();

                    bool UnGhostFile_Status = false;
                    if (OperationType.ToUpper().Trim().Equals("MOVE"))
                    {
                        uploadedFile.MoveTo(directoryName + fileName, MoveOperations.Overwrite);
                        clientContext.ExecuteQuery();

                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[UnGhostFile] Created the new version of the file " + fileName + " using MOVE operation");
                        Console.WriteLine("[UnGhostFile] Created the new version of the file " + fileName + " using MOVE operation");
                        UnGhostFile_Status = true;
                    }
                    else if (OperationType.ToUpper().Trim().Equals("COPY"))
                    {
                        uploadedFile.CopyTo(directoryName + fileName, true);
                        clientContext.ExecuteQuery();

                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[UnGhostFile] Created the new version of the file " + fileName + " using COPY operation");
                        Console.WriteLine("[UnGhostFile] Created the new version of the file " + fileName + " using COPY operation");
                        UnGhostFile_Status = true;
                    }
                    else
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[UnGhostFile] The Operation input in not provided to unghost the file " + fileName + "");
                        Console.WriteLine("[UnGhostFile] The Operation input in not provided to unghost the file " + fileName + "");
                    }

                    //If Un-Ghost File Operation is Successful
                    if (UnGhostFile_Status)
                    {
                        GhostingAndUnGhostingBase objUGBase = new GhostingAndUnGhostingBase();
                        objUGBase.FileName = fileName;
                        objUGBase.FilePath = absoluteFilePath;
                        objUGBase.WebUrl = siteUrl.ToString();
                        objUGBase.SiteCollection = Constants.NotApplicable;
                        objUGBase.WebApplication = Constants.NotApplicable;

                        if (objUGBase != null)
                        {
                            FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.UnGhosting_Output, objUGBase,
                                ref headerCSVColumns);
                        }

                        //Deleting the files, which is downloaded to Un-Ghost the file
                        if (System.IO.File.Exists(downloadedFilePath))
                        {
                            System.IO.File.Delete(downloadedFilePath);
                        }
                        //Deleting the files, which is downloaded to Un-Ghost the file
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] UnGhostFile. Exception Message: " + ex.Message);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "UnGhost", ex.Message, ex.ToString(), "UnGhostFile", ex.GetType().ToString(), exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception] UnGhostFile. Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: UnGhostFile");
            Console.WriteLine("[END] ::: UnGhostFile");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "############## UnGhostFile - Trasnformation Utility Execution Completed for Web ##############");
            Console.WriteLine("############## UnGhostFile - Trasnformation Utility Execution Completed for Web ##############");
        }
        public void DownloadFileFromHive(string absoluteFilePath, string outPutFolder, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "NA", string Password = "NA", string Domain = "NA")
        {
            bool headerCSVColumns = false;
            string exceptionCommentsInfo1 = string.Empty;

            GhostingAndUnGhosting_Initialization(outPutFolder, "DOWNLOAD");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "############## DownloadFileFromHive - Trasnformation Utility Execution Started - For Web ##############");
            Console.WriteLine("############## DownloadFileFromHive - Trasnformation Utility Execution Started - For Web ##############");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
            Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: DownloadFileFromHive");
            Console.WriteLine("[START] ::: DownloadFileFromHive");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[DownloadFileFromHive] Initiated Logger and Exception Class. Logger and Exception file will be available in path " + outPutFolder);
            Console.WriteLine("[DownloadFileFromHive] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);

            try
            {
                exceptionCommentsInfo1 = "FilePath: " + absoluteFilePath;
                string fileName = Path.GetFileName(absoluteFilePath);               

                using (WebClient myWebClient = new WebClient())
                {
                    //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                    if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.OnPremise)
                    {
                        myWebClient.Credentials = new System.Net.NetworkCredential(UserName, Password, Domain);
                    }
                    //SharePointOnline  => OL (Online)
                    else if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.Online)
                    {
                        AuthenticationHelper ObjAuth = new AuthenticationHelper();
                        var spoPassword = ObjAuth.GetSecureString(Password);
                        myWebClient.Credentials = new SharePointOnlineCredentials(UserName, spoPassword);
                    }
                    myWebClient.Credentials = new System.Net.NetworkCredential(UserName, Password, Domain);
                    myWebClient.DownloadFile(absoluteFilePath, outPutFolder + "\\" + fileName);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[DownloadFileFromHive] File is downloaded to the Directory : " + outPutFolder);
                    Console.WriteLine("[DownloadFileFromHive] File is downloaded to the Directory : " + outPutFolder);
                }

                DownloadFileBase objDFBase = new DownloadFileBase();
                objDFBase.GivenFilePath = absoluteFilePath;
                objDFBase.FileName = fileName;
                objDFBase.DownloadedFilePath = outPutFolder + "\\" + fileName;

                objDFBase.WebUrl = Constants.NotApplicable;
                objDFBase.SiteCollection = Constants.NotApplicable;
                objDFBase.WebApplication = Constants.NotApplicable;

                if (objDFBase != null)
                {
                    FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.UnGhosting_DownloadFileReport, objDFBase,
                        ref headerCSVColumns);
                }
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] DownloadFileFromHive. Exception Message: " + ex.Message);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "DownloadFileFromHive", ex.Message, ex.ToString(), "DownloadFileFromHive", ex.GetType().ToString(), exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception] DownloadFileFromHive. Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: DownloadFileFromHive");
            Console.WriteLine("[END] ::: DownloadFileFromHive");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "############## DownloadFileFromHive - Trasnformation Utility Execution Completed for Web ##############");
            Console.WriteLine("############## DownloadFileFromHive - Trasnformation Utility Execution Completed for Web ##############");
        }
        
        public string GetLibraryName(string absolutePath,string siteUrl,string fileName)
        {
            string directoryName = string.Empty;
            string result = string.Empty;
            directoryName = absolutePath.Substring(siteUrl.Length+1);
            string[] tempDirectoryName = directoryName.Split('/');
            for(int i=0;i<tempDirectoryName.Length-1;i++)
            {
                result = result+tempDirectoryName[i] + "/";
            }

            return result; 
        }
        public string GetNextFileName(string fileName)
        {
            string extension = Path.GetExtension(fileName);
            string name = Path.GetFileNameWithoutExtension(fileName);
            string nextUniqueFilename = string.Empty;


            if (name.Contains("-Copy"))
            { // Need to increase the existing index by one or add first index

                int iIndexOfOpenBraces = name.LastIndexOf('(');
                int iIndexOfCloseBraces = name.LastIndexOf(')');
                string ContentAfterOpenBrace = name.Substring(iIndexOfOpenBraces + 1, (iIndexOfCloseBraces - iIndexOfOpenBraces) - 1);

                // check if content after Open Braces is a number, if so increase index by one, if not add the number (1)
                int iCurrentIndex;
                bool bIsContentAfterLastUnderscoreIsNumber = int.TryParse(ContentAfterOpenBrace, out iCurrentIndex);
                if (bIsContentAfterLastUnderscoreIsNumber)
                {
                    iCurrentIndex++;
                    string sContentBeforUnderscore = name.Substring(0, iIndexOfOpenBraces);
                    nextUniqueFilename = string.Format("{0}({1}){2}", sContentBeforUnderscore, iCurrentIndex++, extension);
                }
                else
                {
                    nextUniqueFilename = string.Format("{0}-Copy({1}){2}", name, "1", extension);
                }
            }
            else
            { // No "_Copy" in file name. Simple add first index along with "_Copy"
                nextUniqueFilename = string.Format("{0}-Copy({1}){2}", name, "1", extension);
            }

            return nextUniqueFilename;
        }
       
    }
}
