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
    public class DeleteMissingWorkflowAssociations
    {
        public static void DoWork()
        {
            try
            {
                string timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");
                Logger.OpenLog("DeleteWorkflowAssociations", timeStamp);

                if (!ShowInformation())
                    return;


                string inputFileSpec = Environment.CurrentDirectory + @"\" + Constants.MissingWorkflowAssociationsInputFileName;
                if (System.IO.File.Exists(inputFileSpec))
                {
                    Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);
                    IEnumerable<MissingWorkflowAssociationsInput> objInputMissingWorkflowAssociations = ImportCSV.ReadMatchingColumns<MissingWorkflowAssociationsInput>(inputFileSpec, Constants.CsvDelimeter);
                    if (objInputMissingWorkflowAssociations != null && objInputMissingWorkflowAssociations.Any())
                    {
                        try
                        {
                            string csvFile = Environment.CurrentDirectory + @"/" + Constants.DeleteWorkflowAssociationsStatus + timeStamp + Constants.CSVExtension;
                            if (System.IO.File.Exists(csvFile))
                                System.IO.File.Delete(csvFile);
                            Logger.LogInfoMessage(String.Format("\n[DeleteMissingWorkflowAssociations: DoWork] Preparing to delete a total of {0} files ...", objInputMissingWorkflowAssociations.Cast<Object>().Count()), true);

                            foreach (MissingWorkflowAssociationsInput missingFile in objInputMissingWorkflowAssociations)
                            {
                                DeleteMissingFile(missingFile, csvFile);
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.LogErrorMessage(String.Format("[DeleteMissingWorkflowAssociations: DoWork] failed: Error={0}", ex.Message), true);
                            ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "WorkflowAssociations", ex.Message,
                                ex.ToString(), "DoWork", ex.GetType().ToString(), "Exception occured while reading input file");
                        }
                    }
                    else
                    {
                        Logger.LogInfoMessage("[DeleteMissingWorkflowAssociations: DoWork] There is nothing to delete from the '" + inputFileSpec + "' File ", true);
                    }
                    Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
                }
                else
                {
                    Logger.LogErrorMessage("[DeleteMissingWorkflowAssociations: DoWork] The input file " + inputFileSpec + " is not present", true);
                }

            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[DeleteMissingWorkflowAssociations: DoWork] failed: Error={0}", ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "WorkflowAssociations", ex.Message,
                    ex.ToString(), "DoWork", ex.GetType().ToString(), "Exception occured while reading input file");
            }
            Logger.CloseLog();
        }

        private static void DeleteMissingFile(MissingWorkflowAssociationsInput missingFile, string csvFile)
        {
            bool headerWAOP = false;
            MissingWorkflowAssociationsOutput objWFOP = new MissingWorkflowAssociationsOutput();
            if (missingFile == null)
            {
                return;
            }

            string wfFileDirName = missingFile.DirName;
            string wfFileName = missingFile.LeafName;
            string webAppUrl = missingFile.WebApplication;
            string webUrl = missingFile.WebUrl;

            objWFOP.DirName = wfFileDirName;
            objWFOP.LeafName = wfFileName;
            objWFOP.WebApplication = webAppUrl;
            objWFOP.WebUrl = webUrl;
            objWFOP.SiteCollection = missingFile.SiteCollection;

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
                Logger.LogInfoMessage(String.Format("\n\n[DeleteMissingWorkflowAssociations: DeleteMissingFile] Processing Workflow Association File: {0} ...", webAppUrl + serverRelativeFilePath), true);

                // we have to open the web because Helper.DeleteFileByServerRelativeUrl() needs to update the web in order to commit the change
                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    Web web = userContext.Web;
                    userContext.Load(web);
                    userContext.ExecuteQuery();

                    if (Helper.DeleteFileByServerRelativeUrl(web, serverRelativeFilePath))
                        objWFOP.Status = Constants.Success;
                    else
                        objWFOP.Status = Constants.Failure;
                    //Logger.LogInfoMessage(targetFile.Name + " deleted successfully");
                }

                if (System.IO.File.Exists(csvFile))
                {
                    headerWAOP = true;
                }
                FileUtility.WriteCsVintoFile(csvFile, objWFOP, ref headerWAOP);
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[DeleteMissingWorkflowAssociations: DeleteMissingFile] failed for {0}: Error={1}", serverRelativeFilePath, ex.Message), true);
                ExceptionCsv.WriteException(webAppUrl, Constants.NotApplicable, webUrl, "WorkflowAssociations", ex.Message, ex.ToString(), "DeleteMissingFile",
                    ex.GetType().ToString(), String.Format("DeleteMissingWorkflowAssociationFile() failed for {0}: Error={1}", serverRelativeFilePath, ex.Message));
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