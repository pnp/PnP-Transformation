using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Base;
using JDP.Remediation.Console.Common.Utilities;

namespace JDP.Remediation.Console
{
    /*
        This operation reads an input file for a list of custom master pages and updates the Device Channel Mappings file of the web that contains 
        the master page file to ensure all references to the custom master page have been removed.  Each reference is reset to use the name of the 
        Site Master Page currently in use on the web.

        This operation is helpful in eliminating the last references to a custom master page file; doing so allows the file to be deleted from the 
        Master Page Gallery of the web.
    */
    public class ResetDeviceChannelMappingFiles
    {
        private static string csvOutputFileSpec = String.Empty;
        private static bool csvOutputFileHasHeader = false;

        private static bool ReadInputFile(ref string inputFile)
        {
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            Logger.LogInfoMessage(String.Format("Please enter the complete file path to the input file (e.g., C:\\<Working Directory>\\{0})",
                Constants.LockedMasterPagesFilesInputFileName
                ));
            System.Console.ForegroundColor = System.ConsoleColor.Yellow;
            System.Console.ResetColor();

            inputFile = System.Console.ReadLine();
            Logger.LogInfoMessage(String.Format("ReadInputFile(): User-specified input file path: [{0}]", inputFile), false);

            if (string.IsNullOrEmpty(inputFile) || !System.IO.File.Exists(inputFile))
            {
                return false;
            }
            return true;
        }

        public static void DoWork()
        {
            string timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");

            csvOutputFileSpec = Environment.CurrentDirectory + "\\ResetDeviceChannelMappingFiles-" + timeStamp + Constants.CSVExtension;
            csvOutputFileHasHeader = System.IO.File.Exists(csvOutputFileSpec);

            Logger.OpenLog("ResetDeviceChannelMappingFiles", timeStamp);
            Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);

            string inputFileSpec = String.Empty;
            if (!ReadInputFile(ref inputFileSpec))
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage(String.Format("Input file [{0}] does not exist.", inputFileSpec), true);
                Logger.LogInfoMessage(String.Format("Scan aborted {0}", DateTime.Now.ToString()), true);
                Logger.CloseLog();
                System.Console.ResetColor();
                return;
            }

            // The Locked Master Pages input file is essentially a filtered instance of the Missing Setup Files input file.
            IEnumerable<LockedMasterPageFilesInput> objInputLockedMasterPageFiles = ImportCSV.ReadMatchingColumns<LockedMasterPageFilesInput>(inputFileSpec, Constants.CsvDelimeter);
            if (objInputLockedMasterPageFiles == null || objInputLockedMasterPageFiles.Count() == 0)
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage(String.Format("Input file [{0}] is empty.", inputFileSpec), true);
                Logger.LogInfoMessage(String.Format("Scan aborted {0}", DateTime.Now.ToString()), true);
                Logger.CloseLog();
                System.Console.ResetColor();
                return;
            }

            Logger.LogInfoMessage(String.Format("Preparing to process a total of {0} master page files ...", objInputLockedMasterPageFiles.Count()), true);
            try
            {
                foreach (LockedMasterPageFilesInput masterPageFile in objInputLockedMasterPageFiles)
                {
                    ResetMappingFile(masterPageFile);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ResetDeviceChannelMappingFiles() failed: Error={0}", ex.Message), true);
                ExceptionCsv.WriteException(
                    Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, 
                    "MappingFile", 
                    ex.Message, ex.ToString(), "DoWork", ex.GetType().ToString(), 
                    "Exception occured while processing input file."
                    );
            }

            Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        private static void ResetMappingFile(LockedMasterPageFilesInput masterPageFile)
        {
            if (masterPageFile == null)
            {
                return;
            }

            string setupFileDirName = masterPageFile.SetupFileDirName;
            string setupFileName = masterPageFile.SetupFileName;
            string setupFileExtension = masterPageFile.SetupFileExtension;
            string siteUrl = masterPageFile.SiteCollection;
            string webAppUrl = masterPageFile.WebApplication;
            string webUrl = masterPageFile.WebUrl;

            LockedMasterPageFilesOutput csvObject = new LockedMasterPageFilesOutput();
            csvObject.SetupFileDirName = setupFileDirName;
            csvObject.SetupFileName = setupFileName;
            csvObject.WebApplication = webAppUrl;
            csvObject.WebUrl = webUrl;
            csvObject.ExecutionDateTime = DateTime.Now.ToString();
            csvObject.Status = Constants.Success;

            // clean the inputs
            if (setupFileDirName.EndsWith("/"))
            {
                setupFileDirName = setupFileDirName.TrimEnd(new char[] { '/' });
            }
            if (setupFileName.StartsWith("/"))
            {
                setupFileName = setupFileName.TrimStart(new char[] { '/' });
            }
            if (webUrl.EndsWith("/"))
            {
                webUrl = webUrl.TrimEnd(new char[] { '/' });
            }
            if (webAppUrl.EndsWith("/"))
            {
                webAppUrl = webAppUrl.TrimEnd(new char[] { '/' });
            }

            // e.g., "https://ppeTeams.contoso.com/sites/test/_catalogs/masterpage/Sample.master"
            string targetFilePath = setupFileDirName + '/' + setupFileName;

            // e.g., "https://ppeTeams.contoso.com/sites/test/_catalogs/masterpage/__DeviceChannelMappings.aspx"
            string mappingFilePath = setupFileDirName + '/' + Constants.DeviceChannelMappingFileName;

            // e.g., "/_catalogs/masterpage/Sample.master"
            // e.g., "/_catalogs/masterpage/folder/Sample.master"
            // e.g., "/sites/testSite/_catalogs/masterpage/Sample.master"
            // e.g., "/sites/testSite/_catalogs/masterpage/folder/Sample.master"
            // e.g., "/sites/testSite/childWeb/_catalogs/masterpage/Sample.master"
            // e.g., "/sites/testSite/childWeb/_catalogs/masterpage/folder/Sample.master"
            string serverRelativeMappingFilePath = mappingFilePath.Substring(webAppUrl.Length);

            if (setupFileExtension.Equals("master", StringComparison.InvariantCultureIgnoreCase) == false)
            {
                // ignore anything that is not a master page file
                Logger.LogWarningMessage(String.Format("Skipping file [not a Master Page]: {0}", targetFilePath), true);
                return;
            }

            try
            {
                Logger.LogInfoMessage(String.Format("Processing File: {0} ...", targetFilePath), true);

                Logger.LogInfoMessage(String.Format(" setupFileDirName= {0} ...", setupFileDirName), false);
                Logger.LogInfoMessage(String.Format(" setupFileName= {0} ...", setupFileName), false);
                Logger.LogInfoMessage(String.Format(" targetFilePath= {0} ...", targetFilePath), false);
                Logger.LogInfoMessage(String.Format(" mappingFilePath= {0} ...", mappingFilePath), false);
                Logger.LogInfoMessage(String.Format(" serverRelativeFilePath= {0} ...", serverRelativeMappingFilePath), false);
                Logger.LogInfoMessage(String.Format(" webAppUrl= {0} ...", webAppUrl), false);
                Logger.LogInfoMessage(String.Format(" webUrl= {0} ...", webUrl), false);

                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    Web web = userContext.Web;
                    userContext.Load(web);
                    userContext.ExecuteQuery();

                    // Get Device Channel Mapping File
                    string originalFileContents = Helper.SafeGetFileAsString(web, serverRelativeMappingFilePath);
                    if (String.IsNullOrEmpty(originalFileContents))
                    {
                        // bail if Mapping file not found..
                        Logger.LogWarningMessage(String.Format("Skipping file: [{0}]; Mapping File not found: [{1}]", targetFilePath, serverRelativeMappingFilePath), true);
                        return;
                    }

                    string tempFileContents = originalFileContents.ToLower();
                    if (tempFileContents.Contains(setupFileName.ToLower()) == false)
                    {
                        // bail if MP file not referenced in Mapping file..
                        Logger.LogWarningMessage(String.Format("Skipping file: [{0}]; Mapping File does not reference the master page file: [{1}]", targetFilePath, serverRelativeMappingFilePath), true);
                        return;
                    }

                    // grab the current master page settings for the web.
                    Helper.MasterPageInfo mpi = Helper.GetMasterPageInfo(web);
                    string currentCustomMasterPageUrl = mpi.CustomMasterPageUrl;
                    string currentCustomMasterPage = currentCustomMasterPageUrl.Substring(currentCustomMasterPageUrl.LastIndexOf("/") + 1);

                    // Edit Device Channel Mapping File so it now references the correct/current master page files.
                    // TODO: this is a case-sensitive operation; add case-insensitive logic if it becomes necessary...
                    string updatedFileContents = originalFileContents.Replace(setupFileName, currentCustomMasterPage);

                    // Did the case-sensitive replacement operation fail?
                    tempFileContents = updatedFileContents.ToLower();
                    if (tempFileContents.Contains(setupFileName.ToLower()) == true)
                    {
                        // bail if replacement operation failed due to case-sensitivity
                        Logger.LogErrorMessage(String.Format("ResetMappingFile() failed for file {0}: Error={1}", 
                            targetFilePath, "Casing of Master Page References in Mapping File does not match the casing of the Master Page [" + setupFileName + "] specified in the input file"), 
                            true
                            );
                        Logger.LogWarningMessage(String.Format("Update the casing of the Master Page entry [{0}] of the input file to match the casing used in the Mapping File [Contents={1}]", setupFileName, originalFileContents), false);
                        return;
                    }

                    Logger.LogInfoMessage(String.Format("Reset Mapping File [{0}] to reference [{1}]", serverRelativeMappingFilePath, currentCustomMasterPage), true);

                    // Upload Modified Device Channel Mapping File
                    Helper.UploadDeviceChannelMappingFile(web, serverRelativeMappingFilePath, updatedFileContents, "File reset by Transformation Console");

                    // Backup Original Device Channel Mapping File
                    Helper.UploadDeviceChannelMappingFile(web, serverRelativeMappingFilePath + ".bak", originalFileContents, "Backup created by Transformation Console");

                    csvObject.MappingFile = serverRelativeMappingFilePath;
                    csvObject.MappingBackup = serverRelativeMappingFilePath + ".bak";
                    csvObject.MappingMasterPageRef = currentCustomMasterPage;
                    csvObject.Status = Constants.Success;
                    FileUtility.WriteCsVintoFile(csvOutputFileSpec, csvObject, ref csvOutputFileHasHeader);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ResetMappingFile() failed for file {0}: Error={1}", targetFilePath, ex.Message), true);
                ExceptionCsv.WriteException(
                    webAppUrl, siteUrl, webUrl,
                    "MappingFile", 
                    ex.Message, ex.ToString(), "ResetMappingFile", ex.GetType().ToString(), 
                    String.Format("ResetMappingFile() failed for file {0}", targetFilePath)
                    );
            }
        }
    }
}
