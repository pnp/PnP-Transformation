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
    public class DeleteMissingSetupFiles
    {
        public static void DoWork()
        {
            string timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");
            string setupInputFile = string.Empty;
            Logger.OpenLog("DeleteSetupFiles", timeStamp);

            //if (!ShowInformation())
            //    return;

            if (!ReadInputFile(ref setupInputFile))
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("Setup Files input file is not valid or available. So, Operation aborted!");
                Logger.LogErrorMessage("Please enter path like: E.g. C:\\<Working Directory>\\<InputFile>.csv");
                System.Console.ResetColor();
                return;
            }

            string inputFileSpec = setupInputFile;

            if (System.IO.File.Exists(inputFileSpec))
            {
                Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);
                IEnumerable<MissingSetupFilesInput> objInputMissingSetupFiles = ImportCSV.ReadMatchingColumns<MissingSetupFilesInput>(inputFileSpec, Constants.CsvDelimeter);
                if (objInputMissingSetupFiles != null && objInputMissingSetupFiles.Any())
                {
                    try
                    {
                        string csvFile = Environment.CurrentDirectory + @"\" + Constants.DeleteSetupFileStatus + timeStamp + Constants.CSVExtension;
                        if (System.IO.File.Exists(csvFile))
                            System.IO.File.Delete(csvFile);
                        Logger.LogInfoMessage(String.Format("Preparing to delete a total of {0} files ...", objInputMissingSetupFiles.Cast<Object>().Count()), true);

                        foreach (MissingSetupFilesInput missingFile in objInputMissingSetupFiles)
                        {
                            DeleteMissingFile(missingFile, csvFile);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogErrorMessage(String.Format("[DeleteSetupFiles: DoWork] failed: Error={0}", ex.Message), true);
                        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "SetupFile", ex.Message,
                            ex.ToString(), "DoWork", ex.GetType().ToString(), "Exception occured while reading input file");
                    }
                }
                else
                {
                    Logger.LogInfoMessage("There is nothing to delete from the '" + inputFileSpec + "' File ", true);
                }
                Logger.LogInfoMessage(String.Format("Scan Completed {0}", DateTime.Now.ToString()), true);
            }
            else
                Logger.LogErrorMessage(String.Format("[DeleteSetupFiles: DoWork]: Input file {0} is not available", inputFileSpec), true);
            Logger.CloseLog();
        }

        private static void DeleteMissingFile(MissingSetupFilesInput missingFile, string csvFile)
        {
            bool headerSetupFile = false;
            if (missingFile == null)
            {
                return;
            }

            string setupFileDirName = missingFile.SetupFileDirName;
            string setupFileName = missingFile.SetupFileName;
            string webAppUrl = missingFile.WebApplication;
            string webUrl = missingFile.WebUrl;

            MissingSetupFilesOutput objSetupOP = new MissingSetupFilesOutput();
            objSetupOP.SetupFileDirName = setupFileDirName;
            objSetupOP.SetupFileName = setupFileName;
            objSetupOP.WebApplication = webAppUrl;
            objSetupOP.WebUrl = webUrl;
            objSetupOP.ExecutionDateTime = DateTime.Now.ToString();

            if (webUrl.IndexOf("http", StringComparison.InvariantCultureIgnoreCase) < 0)
            {
                // ignore the header row in case it is still present
                return;
            }

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

            // e.g., "/_catalogs/masterpage/Sample.master"
            // e.g., "/_catalogs/masterpage/folder/Sample.master"
            // e.g., "/sites/testSite/_catalogs/masterpage/Sample.master"
            // e.g., "/sites/testSite/_catalogs/masterpage/folder/Sample.master"
            // e.g., "/sites/testSite/childWeb/_catalogs/masterpage/Sample.master"
            // e.g., "/sites/testSite/childWeb/_catalogs/masterpage/folder/Sample.master"
            string serverRelativeFilePath = targetFilePath.Substring(webAppUrl.Length);

            try
            {
                Logger.LogInfoMessage(String.Format("Processing File: {0} ...", targetFilePath), true);

                //Logger.LogInfoMessage(String.Format("-setupFileDirName= {0}", setupFileDirName), false);
                //Logger.LogInfoMessage(String.Format("-setupFileName= {0}", setupFileName), false);
                //Logger.LogInfoMessage(String.Format("-targetFilePath= {0}", targetFilePath), false);
                //Logger.LogInfoMessage(String.Format("-webAppUrl= {0}", webAppUrl), false);
                //Logger.LogInfoMessage(String.Format("-webUrl= {0}", webUrl), false);
                //Logger.LogInfoMessage(String.Format("-serverRelativeFilePath= {0}", serverRelativeFilePath), false);

                // we have to open the web because Helper.DeleteFileByServerRelativeUrl() needs to update the web in order to commit the change
                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    Web web = userContext.Web;
                    userContext.Load(web);
                    userContext.ExecuteQuery();

                    if (Helper.DeleteFileByServerRelativeUrl(web, serverRelativeFilePath))
                    {
                        Logger.LogInfoMessage(setupFileName + " deleted successfully and output file is present in the path: " + Environment.CurrentDirectory);
                        objSetupOP.Status = Constants.Success;
                    }
                    else
                        objSetupOP.Status = Constants.Failure;
                    if (System.IO.File.Exists(csvFile))
                    {
                        headerSetupFile = true;
                    }
                    FileUtility.WriteCsVintoFile(csvFile, objSetupOP, ref headerSetupFile);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[DeleteSetupFiles: DoWork] failed for {0}: Error={1}", targetFilePath, ex.Message), true);
                ExceptionCsv.WriteException(webAppUrl, Constants.NotApplicable, webUrl, "SetupFile", ex.Message, ex.ToString(), "DeleteMissingFile",
                    ex.GetType().ToString(), String.Format("DeleteSetupFiles > DeleteMissingFile() failed for {0}", targetFilePath));
            }
        }

        private static bool ShowInformation()
        {
            bool doContinue = false;
            string option = string.Empty;
            System.Console.WriteLine(Constants.MissingSetupFilesInputFileName + " file needs to be present in current working directory (where JDP.Remediation.Console.exe is present) for SetupFile cleanup ");
            System.Console.WriteLine("Please make sure you verify the data before executing Clean-up option as Cleaned Setup files can't be rollback.");
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Press 'y' to proceed further. Press any key to go for Clean-Up Menu.");
            option = System.Console.ReadLine().ToLower();
            System.Console.ResetColor();
            if (option.Equals("y", StringComparison.OrdinalIgnoreCase))
                doContinue = true;
            return doContinue;
        }

        private static bool ReadInputFile(ref string setupInputFile)
        {
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            Logger.LogMessage("Enter Complete Input File Path of Pre-Scan Setup File <" + Constants.MissingSetupFilesInputFileName + "> Report:");
            System.Console.ForegroundColor = System.ConsoleColor.Yellow;
            System.Console.WriteLine("Please make sure you verify the data before executing Clean-up option as cleaned Setup Files can't be rollback.");
            System.Console.ResetColor();
            setupInputFile = System.Console.ReadLine();
            Logger.LogMessage("[DeleteSetupFiles : ReadInputFile] Entered Input File of Setup File Data " + setupInputFile, false);
            if (string.IsNullOrEmpty(setupInputFile) || !System.IO.File.Exists(setupInputFile))
                return false;
            return true;
        }
    }
}
