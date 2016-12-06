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
using JDP.Remediation.Console.Common.Base;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Utilities;

namespace JDP.Remediation.Console
{
    public class DeleteMissingEventReceivers
    {
        public static void DoWork()
        {
            string timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");
            string eventReceiverInputFile = string.Empty;
            Logger.OpenLog("DeleteEventReceivers", timeStamp);

            //if (!ShowInformation())
            //  return;

            if (!ReadInputFile(ref eventReceiverInputFile))
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("Event Receivers input file is not valid or available. So, Operation aborted!");
                Logger.LogErrorMessage("Please enter path like: E.g. C:\\<Working Directory>\\<InputFile>.csv");
                System.Console.ResetColor();
                return;
            }

            string inputFileSpec = eventReceiverInputFile;

            if (System.IO.File.Exists(inputFileSpec))
            {
                Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);
                //Read Input file
                IEnumerable<MissingEventReceiversInput> objInputMissingEventReceivers = ImportCSV.ReadMatchingColumns<MissingEventReceiversInput>(inputFileSpec, Constants.CsvDelimeter);
                if (objInputMissingEventReceivers != null && objInputMissingEventReceivers.Any())
                {
                    try
                    {
                        string csvFile = Environment.CurrentDirectory + @"/" + Constants.DeleteEventReceiversStatus + timeStamp + Constants.CSVExtension;
                        if (System.IO.File.Exists(csvFile))
                            System.IO.File.Delete(csvFile);
                        Logger.LogInfoMessage(String.Format("Preparing to delete a total of {0} event receivers ...", objInputMissingEventReceivers.Cast<Object>().Count()), true);

                        foreach (MissingEventReceiversInput MissingEventReceiver in objInputMissingEventReceivers)
                        {
                            DeleteMissingEventReceiver(MissingEventReceiver, csvFile);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogErrorMessage(String.Format("[DeleteMissingEventReceivers: DoWork] failed: Error={0}", ex.Message), true);
                        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "EventReceiver", ex.Message,
                            ex.ToString(), "DoWork", ex.GetType().ToString(), "Exception occured while reading input file");
                    }
                }
                else
                {
                    Logger.LogInfoMessage("There is nothing to delete from the '" + inputFileSpec + "' File ", true);

                }
                Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            }
            else
                Logger.LogErrorMessage(String.Format("[DeleteMissingEventReceivers: DoWork]: Input file {0} is not available", inputFileSpec), true);

            Logger.CloseLog();
        }

        private static void DeleteMissingEventReceiver(MissingEventReceiversInput MissingEventReceiver, string csvFile)
        {
            bool headerEROP = false;
            MissingEventReceiversOutput objEROP = new MissingEventReceiversOutput();
            if (MissingEventReceiver == null)
            {
                return;
            }

            string assemblyName = MissingEventReceiver.Assembly;
            string eventName = MissingEventReceiver.EventName;
            string hostId = MissingEventReceiver.HostId;
            string hostTypeInfo = MissingEventReceiver.HostType;
            string siteUrl = MissingEventReceiver.SiteCollection;
            string webUrl = MissingEventReceiver.WebUrl;
            objEROP.Assembly = assemblyName;
            objEROP.EventName = eventName;
            objEROP.HostId = hostId;
            objEROP.HostType = hostTypeInfo;
            objEROP.SiteCollection = siteUrl;
            objEROP.WebUrl = webUrl;
            objEROP.WebApplication = MissingEventReceiver.WebApplication;
            objEROP.ExecutionDateTime = DateTime.Now.ToString();

            if (webUrl.IndexOf("http", StringComparison.InvariantCultureIgnoreCase) < 0)
            {
                // ignore the header row in case it is still present
                return;
            }

            try
            {
                Logger.LogInfoMessage(String.Format(" Processing Event Receiver [{0}] of Assembly [{1}] ...", eventName, assemblyName), true);

                //Logger.LogInfoMessage(String.Format("-assemblyName= {0}", assemblyName), false);
                //Logger.LogInfoMessage(String.Format("-eventName= {0}", eventName), false);
                //Logger.LogInfoMessage(String.Format("-hostId= {0}", hostId), false);
                //Logger.LogInfoMessage(String.Format("-hostTypeInfo= {0}", hostTypeInfo), false);
                //Logger.LogInfoMessage(String.Format("-siteUrl= {0}", siteUrl), false);
                //Logger.LogInfoMessage(String.Format("-webUrl= {0}", webUrl), false);

                int hostType = -1;
                if (Int32.TryParse(hostTypeInfo, out hostType) == false)
                {
                    Logger.LogErrorMessage(String.Format("[DeleteMissingEventReceivers: DeleteMissingEventReceiver] failed for Event Receiver [{0}]: Error= Unknown HostType value [{1}] ", eventName, hostTypeInfo), false);
                    return;
                }

                switch (hostType)
                {
                    case 0:
                        if (DeleteSiteEventReceiver(siteUrl, eventName, assemblyName))
                            objEROP.Status = Constants.Success;
                        else
                            objEROP.Status = Constants.Failure;
                        break;
                    case 1:
                        if (DeleteWebEventReceiver(webUrl, eventName, assemblyName))
                            objEROP.Status = Constants.Success;
                        else
                            objEROP.Status = Constants.Failure;
                        break;
                    case 2:
                        if (DeleteListEventReceiver(webUrl, hostId, eventName, assemblyName))
                            objEROP.Status = Constants.Success;
                        else
                            objEROP.Status = Constants.Failure;
                        break;

                    default:
                        Logger.LogErrorMessage(String.Format("[DeleteMissingEventReceivers: DeleteMissingEventReceiver] failed for Event Receiver [{0}] of Assembly [{1}]: Error= Unknown HostType value [{2}] ", eventName, assemblyName, hostTypeInfo), false);
                        objEROP.Status = Constants.Failure;
                        return;
                }

                if (System.IO.File.Exists(csvFile))
                {
                    headerEROP = true;
                }
                FileUtility.WriteCsVintoFile(csvFile, objEROP, ref headerEROP);
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[DeleteMissingEventReceivers: DeleteMissingEventReceiver] failed for Event Receiver [{0}] of Assembly [{1}]: Error={2}", eventName, assemblyName, ex.Message), true);
                ExceptionCsv.WriteException(MissingEventReceiver.WebApplication, siteUrl, webUrl, "EventReceiver", ex.Message, ex.ToString(), "DeleteMissingEventReceiver",
                    ex.GetType().ToString(), String.Format("[DeleteMissingEventReceivers: DeleteMissingEventReceiver] failed for Event Receiver [{0}] of Assembly [{1}]: Error={2}", eventName, assemblyName, ex.Message));
            }
        }

        private static bool DeleteSiteEventReceiver(string siteUrl, string eventName, string assemblyName)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Deleting SITE Event Receiver [{0}] from site {1} ...", eventName, siteUrl), true);

                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, siteUrl))
                {
                    Site site = userContext.Site;
                    EventReceiverDefinitionCollection receivers = site.EventReceivers;
                    userContext.Load(site);
                    userContext.Load(receivers);
                    userContext.ExecuteQuery();

                    foreach (EventReceiverDefinition receiver in receivers)
                    {
                        if (receiver.ReceiverName.Equals(eventName, StringComparison.InvariantCultureIgnoreCase) &&
                            receiver.ReceiverAssembly.Equals(assemblyName, StringComparison.InvariantCultureIgnoreCase)
                            )
                        {
                            receiver.DeleteObject();
                            userContext.ExecuteQuery();
                            Logger.LogSuccessMessage(String.Format("[DeleteMissingEventReceivers: DeleteSiteEventReceiver] Deleted SITE Event Receiver [{0}] from site {1} and output file is present in the path: {2}", eventName, siteUrl, Environment.CurrentDirectory), true);
                            return true;
                        }
                    }
                    Logger.LogErrorMessage(String.Format("[DeleteMissingEventReceivers: DeleteSiteEventReceiver] DeleteSiteEventReceiver() failed for Event Receiver [{0}] on site {1}; Error=Event Receiver not Found.", eventName, siteUrl), true);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[DeleteMissingEventReceivers: DeleteSiteEventReceiver] failed for Event Receiver [{0}] on site {1}; Error={2}", eventName, siteUrl, ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, siteUrl, Constants.NotApplicable, "EventReceiver", ex.Message, ex.ToString(), "DeleteSiteEventReceiver",
                    ex.GetType().ToString(), String.Format("[DeleteMissingEventReceivers: DeleteSiteEventReceiver] failed for Event Receiver [{0}] on site {1}", eventName, siteUrl));
                return false;
            }
            return false;
        }
        private static bool DeleteWebEventReceiver(string webUrl, string eventName, string assemblyName)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Deleting WEB Event Receiver [{0}] from web {1} ...", eventName, webUrl), true);

                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    Web web = userContext.Web;
                    EventReceiverDefinitionCollection receivers = web.EventReceivers;
                    userContext.Load(web);
                    userContext.Load(receivers);
                    userContext.ExecuteQuery();

                    foreach (EventReceiverDefinition receiver in receivers)
                    {
                        if (receiver.ReceiverName.Equals(eventName, StringComparison.InvariantCultureIgnoreCase) &&
                            receiver.ReceiverAssembly.Equals(assemblyName, StringComparison.InvariantCultureIgnoreCase)
                            )
                        {
                            receiver.DeleteObject();
                            userContext.ExecuteQuery();
                            Logger.LogSuccessMessage(String.Format("[DeleteMissingEventReceivers: DeleteWebEventReceiver] Deleted WEB Event Receiver [{0}] from web {1} and output file is present in the path: {2}", eventName, webUrl, Environment.CurrentDirectory), true);
                            return true;
                        }
                    }
                    Logger.LogErrorMessage(String.Format("[DeleteMissingEventReceivers: DeleteWebEventReceiver] DeleteWebEventReceiver() failed for Event Receiver [{0}] on web {1}; Error=Event Receiver not Found.", eventName, webUrl), true);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[DeleteMissingEventReceivers: DeleteWebEventReceiver] failed for Event Receiver [{0}] on web {1}; Error={2}", eventName, webUrl, ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "EventReceiver", ex.Message, ex.ToString(), "DeleteWebEventReceiver",
                    ex.GetType().ToString(), String.Format("DeleteWebEventReceiver() failed for Event Receiver [{0}] on web {1}", eventName, webUrl));
            }
            return false;
        }
        private static bool DeleteListEventReceiver(string webUrl, string hostId, string eventName, string assemblyName)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Deleting LIST Event Receiver [{0}] from list [{1}] on web {2} ...", eventName, hostId, webUrl), true);

                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    Web web = userContext.Web;
                    userContext.Load(web);
                    userContext.ExecuteQuery();

                    //ListCollection lists = web.Lists;
                    //userContext.Load(lists);

                    Guid listId = new Guid(hostId);
                    List list = web.Lists.GetById(listId);
                    userContext.Load(list);
                    userContext.ExecuteQuery();
                    if (list == null)
                    {
                        Logger.LogErrorMessage(String.Format("[DeleteMissingEventReceivers: DeleteListEventReceiver] failed for Event Receiver [{0}] on list [{1}] of web {2}; Error=List not Found.", eventName, hostId, webUrl), true);
                        return false;
                    }

                    EventReceiverDefinitionCollection receivers = list.EventReceivers;
                    userContext.Load(receivers);
                    userContext.ExecuteQuery();

                    foreach (EventReceiverDefinition receiver in receivers)
                    {
                        if (receiver.ReceiverName.Equals(eventName, StringComparison.InvariantCultureIgnoreCase) &&
                            receiver.ReceiverAssembly.Equals(assemblyName, StringComparison.InvariantCultureIgnoreCase)
                            )
                        {
                            receiver.DeleteObject();
                            userContext.ExecuteQuery();
                            Logger.LogSuccessMessage(String.Format("[DeleteMissingEventReceivers: DeleteListEventReceiver] Deleted LIST Event Receiver [{0}] from list [{1}] on web {2} and output file is present in the path: {3}", eventName, list.Title, webUrl, Environment.CurrentDirectory), true);
                            return true;
                        }
                    }
                    Logger.LogErrorMessage(String.Format("[DeleteMissingEventReceivers: DeleteListEventReceiver] failed for Event Receiver [{0}] on list [{1}] of web {2}; Error=Event Receiver not Found.", eventName, list.Title, webUrl), true);
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.ToLower().Contains("list does not exist"))
                {
                    Logger.LogErrorMessage(String.Format("[DeleteMissingEventReceivers: DeleteListEventReceiver] failed for Event Receiver [{0}] on list [{1}] of web {2}; Error=List not Found.", eventName, hostId, webUrl), true);
                    ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "EventReceiver", ex.Message, ex.ToString(), "DeleteListEventReceiver",
                    ex.GetType().ToString(), String.Format("DeleteListEventReceiver() failed for Event Receiver [{0}] on list [{1}] of web {2}; Error=List not Found.", eventName, hostId, webUrl));
                    return false;
                }
                Logger.LogErrorMessage(String.Format("[DeleteMissingEventReceivers: DeleteListEventReceiver] failed for Event Receiver [{0}] on list [{1}] of web {2}; Error={3}", eventName, hostId, webUrl, ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, webUrl, "EventReceiver", ex.Message, ex.ToString(), "DeleteListEventReceiver",
                    ex.GetType().ToString(), String.Format("DeleteListEventReceiver() failed for Event Receiver [{0}] on list [{1}] of web {2}; Error={3}", eventName, hostId, webUrl, ex.Message));
            }
            return false;
        }
        private static bool ShowInformation()
        {
            bool doContinue = false;
            string option = string.Empty;
            System.Console.WriteLine("Event Receivers Input File (Pre-Scan OR Discovery Report) file needs to be present in current working directory (where JDP.Remediation.Console.exe is present) for EventReceivers cleanup ");
            System.Console.WriteLine("Please make sure you verify the data before executing Clean-up option as cleaned EventReceivers can't be rollback.");
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Press 'y' to proceed further. Press any key to go for Clean-Up Menu.");
            System.Console.ResetColor();
            option = System.Console.ReadLine().ToLower();
            if (option.Equals("y", StringComparison.OrdinalIgnoreCase))
                doContinue = true;
            return doContinue;
        }

        private static bool ReadInputFile(ref string eventReceiverInputFile)
        {
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            Logger.LogMessage("Enter Complete Input File Path of Event Receivers Report Either Pre-Scan OR Discovery Report:");
            System.Console.ForegroundColor = System.ConsoleColor.Yellow;
            System.Console.WriteLine("Please make sure you verify the data before executing Clean-up option as cleaned EventReceivers can't be rollback.");
            System.Console.ResetColor();
            eventReceiverInputFile = System.Console.ReadLine();
            Logger.LogMessage("[DeleteMissingEventReceivers.csv : ReadInputFile] Entered Input File of Event Receiver Data " + eventReceiverInputFile, false);
            if (string.IsNullOrEmpty(eventReceiverInputFile) || !System.IO.File.Exists(eventReceiverInputFile))
                return false;
            return true;
        }
    }
}
