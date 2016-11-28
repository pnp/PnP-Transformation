using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Reflection;
using System.Security;
using System.Text;
using System.Threading.Tasks;

using OfficeDevPnP.Core;

using Microsoft.SharePoint.Client;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Base;
using JDP.Remediation.Console.Common.Utilities;

namespace JDP.Remediation.Console
{
    public class DeleteMissingFeatures
    {
        public static void DoWork()
        {
            string timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");
            string featureInputFile = string.Empty;
            Logger.OpenLog("DeleteFeatures", timeStamp);

            //if (!ShowInformation())
            //    return;

            if (!ReadInputFile(ref featureInputFile))
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("Features input file is not valid or available. So, Operation aborted!");
                Logger.LogErrorMessage("Please enter path like: E.g. C:\\<Working Directory>\\<InputFile>.csv");
                System.Console.ResetColor();
                return;
            }

            string inputFileSpec = featureInputFile;

            if (System.IO.File.Exists(inputFileSpec))
            {
                Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);

                IEnumerable<MissingFeaturesInput> objInputMissingFeatures = ImportCSV.ReadMatchingColumns<MissingFeaturesInput>(inputFileSpec, Constants.CsvDelimeter);

                if (objInputMissingFeatures != null && objInputMissingFeatures.Any())
                {
                    try
                    {
                        string csvFile = Environment.CurrentDirectory + @"\" + Constants.DeleteFeatureStatus + timeStamp + Constants.CSVExtension;
                        if (System.IO.File.Exists(csvFile))
                            System.IO.File.Delete(csvFile);
                        Logger.LogInfoMessage(String.Format("Preparing to delete a total of {0} features ...", objInputMissingFeatures.Cast<Object>().Count()), true);

                        foreach (MissingFeaturesInput missingFeature in objInputMissingFeatures)
                        {
                            DeleteMissingFeature(missingFeature, csvFile);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogErrorMessage(String.Format("[DeleteMissingFeatures: DoWork] failed: Error={0}", ex.Message), true);
                        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Feature", ex.Message, ex.ToString(),
                            "DoWork", ex.GetType().ToString(), "Exception occured while reading input file");
                    }
                }
                else
                {
                    Logger.LogInfoMessage("There is nothing to delete from the '" + inputFileSpec + "' File ", true);

                }
                Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            }
            else
                Logger.LogErrorMessage(String.Format("[DeleteMissingFeatures: DoWork] failed: Input file {0} is not available", inputFileSpec), false);

            Logger.CloseLog();
        }

        private static void DeleteMissingFeature(MissingFeaturesInput missingFeature, string csvFile)
        {
            bool headerFeatureOP = false;
            MissingFeatureOutput objFeatureOP = new MissingFeatureOutput();
            if (missingFeature == null)
            {
                return;
            }

            string featureId = missingFeature.FeatureId;
            string targetUrl = string.Empty;
            objFeatureOP.FeatureId = featureId; ;
            objFeatureOP.Scope = missingFeature.Scope;
            objFeatureOP.WebApplication = missingFeature.WebApplication;
            objFeatureOP.ExecutionDateTime = DateTime.Now.ToString();

            if (missingFeature.Scope == "Site")
            {
                targetUrl = missingFeature.SiteCollection;
                objFeatureOP.SiteCollection = targetUrl;
                objFeatureOP.WebUrl = Constants.NotApplicable;
            }
            else
            {
                targetUrl = missingFeature.WebUrl;
                objFeatureOP.SiteCollection = missingFeature.SiteCollection;
                objFeatureOP.WebUrl = missingFeature.WebUrl;
            }

            if (targetUrl.IndexOf("http", StringComparison.InvariantCultureIgnoreCase) < 0)
            {
                // ignore the header row in case it is still present
                return;
            }

            Guid featureDefinitionId = new Guid(featureId);
            try
            {
                Logger.LogInfoMessage(String.Format("Processing Feature {0} on {1} {2} ...", featureId, missingFeature.Scope, targetUrl), true);

                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, targetUrl))
                {
                    //RemoveFeature(userContext, featureDefinitionId);
                    switch (missingFeature.Scope.ToLower())
                    {
                        case "site":
                            userContext.Load(userContext.Site);
                            userContext.ExecuteQuery();
                            if (RemoveFeatureFromSite(userContext, featureDefinitionId))
                            {
                                Logger.LogSuccessMessage(String.Format("Deleted Feature {0} from site {1} and output file is present in the path: {2}", featureDefinitionId.ToString(), userContext.Site.Url, Environment.CurrentDirectory), true);
                                objFeatureOP.Status = Constants.Success;
                            }
                            else
                            {
                                Logger.LogErrorMessage(String.Format("[DeleteMissingFeatures: DeleteMissingFeature] Could not delete site Feature {0}; feature not found", featureDefinitionId.ToString()), true);
                                objFeatureOP.Status = Constants.Failure;
                            }
                            break;

                        case "web":
                            userContext.Load(userContext.Web);
                            userContext.ExecuteQuery();
                            if (RemoveFeatureFromWeb(userContext, featureDefinitionId))
                            {
                                Logger.LogSuccessMessage(String.Format("Deleted Feature {0} from web {1} and output file is present in the path: {2}", featureDefinitionId.ToString(), userContext.Web.Url, Environment.CurrentDirectory), true);
                                objFeatureOP.Status = Constants.Success;
                            }
                            else
                            {
                                Logger.LogErrorMessage(String.Format("[DeleteMissingFeatures: DeleteMissingFeature] Could not delete web Feature {0}; feature not found", featureDefinitionId.ToString()), true);
                                objFeatureOP.Status = Constants.Failure;
                            }
                            break;
                    }
                }
                if (System.IO.File.Exists(csvFile))
                {
                    headerFeatureOP = true;
                }
                FileUtility.WriteCsVintoFile(csvFile, objFeatureOP, ref headerFeatureOP);
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[DeleteMissingFeatures: DeleteMissingFeature] failed for Feature {0} on {1} {2}; Error={3}", featureDefinitionId.ToString(), missingFeature.Scope, targetUrl, ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, targetUrl, targetUrl, "Feature", ex.Message, ex.ToString(), "DeleteMissingFeature",
                    ex.GetType().ToString(), String.Format("DeleteMissingFeature() failed for Feature {0} on {1} {2}", featureDefinitionId.ToString(), missingFeature.Scope, targetUrl));
            }
        }

        private static void RemoveFeature(ClientContext userContext, Guid featureDefinitionId)
        {
            userContext.Load(userContext.Site);
            userContext.Load(userContext.Web);
            userContext.ExecuteQuery();

            // The scope of the feature is unknown; search the web-scoped features first; if the feature is not found, then search the site-scoped features

            if (RemoveFeatureFromWeb(userContext, featureDefinitionId))
            {
                Logger.LogSuccessMessage(String.Format("Deleted Feature {0} from web {1}", featureDefinitionId.ToString(), userContext.Web.Url), true);
                return;
            }

            Logger.LogInfoMessage(String.Format("Feature was not found in the web-scoped features; trying the site-scoped features..."), false);

            if (RemoveFeatureFromSite(userContext, featureDefinitionId))
            {
                Logger.LogSuccessMessage(String.Format("Deleted Feature {0} from site {1}", featureDefinitionId.ToString(), userContext.Site.Url), true);
                return;
            }

            Logger.LogErrorMessage(String.Format("Could not delete Feature {0}; feature not found", featureDefinitionId.ToString()), false);
        }

        private static bool RemoveFeatureFromSite(ClientContext userContext, Guid featureDefinitionId)
        {
            try
            {
                FeatureCollection features = userContext.Site.Features;
                ClearObjectData(features);

                userContext.Load(features);
                userContext.ExecuteQuery();

                //DumpFeatures(features, "site");

                Feature targetFeature = features.GetById(featureDefinitionId);
                targetFeature.EnsureProperties(f => f.DefinitionId);

                if (targetFeature == null || !targetFeature.IsPropertyAvailable("DefinitionId") || targetFeature.ServerObjectIsNull.Value)
                {
                    Logger.LogInfoMessage(String.Format("Could not delete Feature {0}; feature not found in site.Features", featureDefinitionId.ToString()), false);
                    return false;
                }
                features.Remove(featureDefinitionId, true);

                // commit the changes
                userContext.Load(features);
                userContext.ExecuteQuery();
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[DeleteMissingFeatures: RemoveFeatureFromSite] failed for Feature {0} on site {1}; Error={2}", featureDefinitionId.ToString(), userContext.Site.Url, ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, userContext.Site.Url, Constants.NotApplicable, "Feature", ex.Message, ex.ToString(), "RemoveFeatureFromSite",
                    ex.GetType().ToString(), String.Format("RemoveFeatureFromSite() failed for Feature {0} on site {1}", featureDefinitionId.ToString(), userContext.Site.Url));
                return false;
            }
        }

        private static bool RemoveFeatureFromWeb(ClientContext userContext, Guid featureDefinitionId)
        {
            try
            {
                FeatureCollection features = userContext.Web.Features;
                ClearObjectData(features);

                userContext.Load(features);
                userContext.ExecuteQuery();

                //DumpFeatures(features, "web");

                Feature targetFeature = features.GetById(featureDefinitionId);
                targetFeature.EnsureProperties(f => f.DefinitionId);

                if (targetFeature == null || !targetFeature.IsPropertyAvailable("DefinitionId") || targetFeature.ServerObjectIsNull.Value)
                {
                    Logger.LogInfoMessage(String.Format("Could not delete Feature {0}; feature not found in web.Features", featureDefinitionId.ToString()), false);
                    return false;
                }

                features.Remove(featureDefinitionId, true);

                // commit the changes
                userContext.Load(features);
                userContext.ExecuteQuery();
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[DeleteMissingFeatures: RemoveFeatureFromWeb] failed for Feature {0} on web {1}; Error={2}", featureDefinitionId.ToString(), userContext.Web.Url, ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, userContext.Web.Url, "Feature", ex.Message, ex.ToString(), "RemoveFeatureFromWeb",
                    ex.GetType().ToString(), String.Format("RemoveFeatureFromWeb() failed for Feature {0} on web {1}", featureDefinitionId.ToString(), userContext.Web.Url));
                return false;
            }
        }

        private static void ClearObjectData(ClientObject clientObject)
        {
            PropertyInfo info_ClientObject = typeof(ClientObject).GetProperty("ObjectData", BindingFlags.NonPublic | BindingFlags.Instance);

            var objectData = (ClientObjectData)info_ClientObject.GetValue(clientObject, new object[0]);
            objectData.MethodReturnObjects.Clear();
        }

        private static void DumpFeatures(FeatureCollection features, string scope)
        {
            Logger.LogInfoMessage(String.Format("Dumping [{0}] {1}-scoped features...", features.Count, scope), false);
            foreach (Feature f in features)
            {
                Logger.LogInfoMessage(f.DefinitionId.ToString(), false);
            }
            Logger.LogInfoMessage("-------------------------------------------------", false);
        }

        private static bool ShowInformation()
        {
            bool doContinue = false;
            string option = string.Empty;
            System.Console.WriteLine(" Features Input File (Pre-Scan OR Discovery Report) needs to be present in current working directory (where JDP.Remediation.Console.exe is present) for Feature cleanup ");
            System.Console.WriteLine("Please make sure you verify the data before executing Clean-up option as cleaned Features can't be rollback.");
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Press 'y' to proceed further. Press any key to go for Clean-Up Menu.");
            System.Console.ResetColor();
            option = System.Console.ReadLine().ToLower();
            if (option.Equals("y", StringComparison.OrdinalIgnoreCase))
                doContinue = true;
            return doContinue;
        }

        private static bool ReadInputFile(ref string featureInputFile)
        {
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            Logger.LogMessage("Enter Complete Input File Path of Features Report Either Pre-Scan OR Discovery Report:");
            System.Console.ForegroundColor = System.ConsoleColor.Yellow;
            System.Console.WriteLine("Please make sure you verify the data before executing Clean-up option as cleaned Features can't be rollback.");
            System.Console.ResetColor();
            featureInputFile = System.Console.ReadLine();
            Logger.LogMessage("[DeleteFeatures : ReadInputFile] Entered Input File of Feature Data " + featureInputFile, false);
            if (string.IsNullOrEmpty(featureInputFile) || !System.IO.File.Exists(featureInputFile))
                return false;
            return true;
        }
    }
}
