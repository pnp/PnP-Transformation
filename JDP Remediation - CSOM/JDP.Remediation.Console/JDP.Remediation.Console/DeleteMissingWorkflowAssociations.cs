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
    public class DeleteMissingWorkflowAssociations
    {
        public static void DoWork()
        {
            try
            {
                Logger.OpenLog("DeleteMissingWorkflowAssociations");
                System.Console.WriteLine("Enter the path of input file PreMT_MissingWorkflowAssociations.csv");
                string filePath = System.Console.ReadLine();
                if (string.IsNullOrEmpty(filePath) || !System.IO.Directory.Exists(filePath))
                {
                    Logger.LogWarningMessage("Input FilePath '" + filePath + "' is not valid", true);
                    filePath = Environment.CurrentDirectory;
                    Logger.LogInfoMessage("Correct Input FilePath is not provided so it changed to current environment '" + filePath + "'", true);
                }

                Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);
                string inputFileSpec = filePath + "\\" + Constants.MissingWorkflowAssociationsInputFileName;
                if (System.IO.File.Exists(inputFileSpec))
                {
                    IEnumerable<MissingWorkflowAssociationsInput> objInputMissingWorkflowAssociations = ImportCSV.ReadMatchingColumns<MissingWorkflowAssociationsInput>(inputFileSpec, Constants.CsvDelimeter);
                    if (objInputMissingWorkflowAssociations != null)
                    {
                        try
                        {
                            Logger.LogInfoMessage(String.Format("\nPreparing to delete a total of {0} files ...", objInputMissingWorkflowAssociations.Cast<Object>().Count()), true);

                            foreach (MissingWorkflowAssociationsInput missingFile in objInputMissingWorkflowAssociations)
                            {
                                DeleteMissingFile(missingFile);
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.LogErrorMessage(String.Format("DeleteMissingWorkflowAssociationFiles() failed: Error={0}", ex.Message), true);
                        }
                    }
                    else
                    {
                        Logger.LogInfoMessage("There is nothing to delete from the '" + inputFileSpec + "' File ", true);
                    }
                }
                else
                {
                    Logger.LogErrorMessage("The input file " + inputFileSpec + " is not present", true);
                }

            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("DeleteMissingWorkflowAssociationFiles() failed: Error={0}", ex.Message), true);
            }
            Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        private static void DeleteMissingFile(MissingWorkflowAssociationsInput missingFile)
        {

            if (missingFile == null)
            {
                return;
            }

            string wfFileDirName = missingFile.DirName;
            string wfFileName = missingFile.LeafName;
            string webAppUrl = missingFile.WebApplication;
            string webUrl = missingFile.WebUrl;

            if (webUrl.IndexOf("http", StringComparison.InvariantCultureIgnoreCase) < 0)
            {
                // ignore the header row in case it is still present
                return;
            }

            // clean the inputs
            if (wfFileDirName.EndsWith("/"))
            {
                wfFileDirName = wfFileDirName.TrimEnd(new char[] { '/' });
            }
            if (!wfFileDirName.StartsWith("/"))
            {
                wfFileDirName = "/" + wfFileDirName;
            }
            if (wfFileName.StartsWith("/"))
            {
                wfFileName = wfFileName.TrimStart(new char[] { '/' });
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
            string serverRelativeFilePath = wfFileDirName + '/' + wfFileName;

            try
            {
                Logger.LogInfoMessage(String.Format("\n\nProcessing Workflow Association File: {0} ...", webAppUrl + serverRelativeFilePath), true);

                // we have to open the web because Helper.DeleteFileByServerRelativeUrl() needs to update the web in order to commit the change
                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    Web web = userContext.Web;
                    userContext.Load(web);
                    userContext.ExecuteQuery();

                    Helper.DeleteFileByServerRelativeUrl(web, serverRelativeFilePath);
                    //Logger.LogInfoMessage(targetFile.Name + " deleted successfully");
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("DeleteMissingWorkflowAssociationFile() failed for {0}: Error={1}", serverRelativeFilePath, ex.Message), false);
            }
        }

        private static bool ShowInformation()
        {
            bool doContinue = false;
            string option = string.Empty;
            System.Console.WriteLine(Constants.MissingWorkflowAssociationsInputFileName + " file needs to be present in current working directory (where JDP.Remediation.Console.exe is present) for Workflow Associations cleanup ");
            System.Console.WriteLine("Please make sure you verify the data before executing Clean-up option as cleaned Workflow Associations can't be rollback.");
            System.Console.WriteLine("Press 'y' to proceed further. Press any key to go for Clean-Up Menu.");
            option = System.Console.ReadLine().ToLower();
            if (option.Equals("y", StringComparison.OrdinalIgnoreCase))
                doContinue = true;
            return doContinue;
        }
    }
}