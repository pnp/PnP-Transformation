using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Base;

namespace JDP.Remediation.Console
{
    public class DeleteMissingSetupFiles
    {
        public static void DoWork()
        {
            Logger.OpenLog("DeleteMissingSetupFiles");
            if (!ShowInformation())
                return;
            Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);
            string inputFileSpec = Environment.CurrentDirectory + "\\" + Constants.MissingSetupFilesInputFileName;
            IEnumerable<MissingSetupFilesInput> objInputMissingSetupFiles = ImportCSV.ReadMatchingColumns<MissingSetupFilesInput>(inputFileSpec, Constants.CsvDelimeter);
            if (objInputMissingSetupFiles != null)
            {
                try
                {
                    Logger.LogInfoMessage(String.Format("Preparing to delete a total of {0} files ...", objInputMissingSetupFiles.Cast<Object>().Count()), true);

                    foreach (MissingSetupFilesInput missingFile in objInputMissingSetupFiles)
                    {
                        DeleteMissingFile(missingFile);
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogErrorMessage(String.Format("DeleteMissingSetupFiles() failed: Error={0}", ex.Message), true);
                }
            }
            else
            {
                Logger.LogInfoMessage("There is nothing to delete from the '" + inputFileSpec + "' File ", true);
            }
            Logger.LogInfoMessage(String.Format("Scan Completed {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        private static void DeleteMissingFile(MissingSetupFilesInput missingFile)
        {

            if (missingFile == null)
            {
                return;
            }

            string setupFileDirName = missingFile.SetupFileDirName;
            string setupFileName = missingFile.SetupFileName;
            string webAppUrl = missingFile.WebApplication;
            string webUrl = missingFile.WebUrl;

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

                    Helper.DeleteFileByServerRelativeUrl(web, serverRelativeFilePath);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("DeleteMissingFile() failed for {0}: Error={1}", targetFilePath, ex.Message), false);
            }
        }

        private static bool ShowInformation()
        {
            bool doContinue = false;
            string option = string.Empty;
            System.Console.WriteLine(Constants.MissingSetupFilesInputFileName + " file needs to be present in current working directory (where JDP.Remediation.Console.exe is present) for SetupFile cleanup ");
            System.Console.WriteLine("Please make sure you verify the data before executing Clean-up option as Cleaned Setup files can't be rollback.");
            System.Console.WriteLine("Press 'y' to proceed further. Press any key to go for Clean-Up Menu.");
            option = System.Console.ReadLine().ToLower();
            if (option.Equals("y", StringComparison.OrdinalIgnoreCase))
                doContinue = true;
            return doContinue;
        }
    }
}
