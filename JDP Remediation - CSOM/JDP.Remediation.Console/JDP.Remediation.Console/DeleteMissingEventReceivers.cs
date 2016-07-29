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

namespace JDP.Remediation.Console
{
    public class DeleteMissingEventReceivers
    {
        public static void DoWork()
        {
            Logger.OpenLog("DeleteMissingEventReceivers");
            if (!ShowInformation())
                return;
            Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);

            string inputFileSpec = Environment.CurrentDirectory + "\\" + Constants.MissingEventReceiversInputFileName;
            //Read Input file
            IEnumerable<MissingEventReceiversInput> objInputMissingEventReceivers = ImportCSV.ReadMatchingColumns<MissingEventReceiversInput>(inputFileSpec, Constants.CsvDelimeter);
            if (objInputMissingEventReceivers != null)
            {
                try
                {
                    Logger.LogInfoMessage(String.Format("Preparing to delete a total of {0} event receivers ...", objInputMissingEventReceivers.Cast<Object>().Count()), true);

                    foreach (MissingEventReceiversInput MissingEventReceiver in objInputMissingEventReceivers)
                    {
                        DeleteMissingEventReceiver(MissingEventReceiver);
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogErrorMessage(String.Format("DeleteMissingEventReceivers() failed: Error={0}", ex.Message), true);
                }
            }
            else
            {
                Logger.LogInfoMessage("There is nothing to delete from the '" + inputFileSpec + "' File ", true);

            }
            Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            Logger.CloseLog();
        }

        private static void DeleteMissingEventReceiver(MissingEventReceiversInput MissingEventReceiver)
        {
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

            if (webUrl.IndexOf("http", StringComparison.InvariantCultureIgnoreCase) < 0)
            {
                // ignore the header row in case it is still present
                return;
            }

            try
            {
                Logger.LogInfoMessage(String.Format("Processing Event Receiver [{0}] of Assembly [{1}] ...", eventName, assemblyName), true);

                //Logger.LogInfoMessage(String.Format("-assemblyName= {0}", assemblyName), false);
                //Logger.LogInfoMessage(String.Format("-eventName= {0}", eventName), false);
                //Logger.LogInfoMessage(String.Format("-hostId= {0}", hostId), false);
                //Logger.LogInfoMessage(String.Format("-hostTypeInfo= {0}", hostTypeInfo), false);
                //Logger.LogInfoMessage(String.Format("-siteUrl= {0}", siteUrl), false);
                //Logger.LogInfoMessage(String.Format("-webUrl= {0}", webUrl), false);

                int hostType = -1;
                if (Int32.TryParse(hostTypeInfo, out hostType) == false)
                {
                    Logger.LogErrorMessage(String.Format("DeleteMissingEventReceiver() failed for Event Receiver [{0}]: Error= Unknown HostType value [{1}] ", eventName, hostTypeInfo), false);
                    return;
                }

                switch (hostType)
                {
                    case 0:
                        DeleteSiteEventReceiver(siteUrl, eventName, assemblyName);
                        break;
                    case 1:
                        DeleteWebEventReceiver(webUrl, eventName, assemblyName);
                        break;
                    case 2:
                        DeleteListEventReceiver(webUrl, hostId, eventName, assemblyName);
                        break;

                    default:
                        Logger.LogErrorMessage(String.Format("DeleteMissingEventReceiver() failed for Event Receiver [{0}] of Assembly [{1}]: Error= Unknown HostType value [{2}] ", eventName, assemblyName, hostTypeInfo), false);
                        return;
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("DeleteMissingEventReceiver() failed for Event Receiver [{0}] of Assembly [{1}]: Error={2}", eventName, assemblyName, ex.Message), false);
            }
        }

        private static void DeleteSiteEventReceiver(string siteUrl, string eventName, string assemblyName)
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
                            Logger.LogSuccessMessage(String.Format("Deleted SITE Event Receiver [{0}] from site {1}", eventName, siteUrl), false);
                            return;
                        }
                    }
                    Logger.LogErrorMessage(String.Format("DeleteSiteEventReceiver() failed for Event Receiver [{0}] on site {1}; Error=Event Receiver not Found.", eventName, siteUrl), false);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("DeleteSiteEventReceiver() failed for Event Receiver [{0}] on site {1}; Error={2}", eventName, siteUrl, ex.Message), false);
            }
        }
        private static void DeleteWebEventReceiver(string webUrl, string eventName, string assemblyName)
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
                            Logger.LogSuccessMessage(String.Format("Deleted WEB Event Receiver [{0}] from web {1}", eventName, webUrl), false);
                            return;
                        }
                    }
                    Logger.LogErrorMessage(String.Format("DeleteWebEventReceiver() failed for Event Receiver [{0}] on web {1}; Error=Event Receiver not Found.", eventName, webUrl), false);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("DeleteWebEventReceiver() failed for Event Receiver [{0}] on web {1}; Error={2}", eventName, webUrl, ex.Message), false);
            }
        }
        private static void DeleteListEventReceiver(string webUrl, string hostId, string eventName, string assemblyName)
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
                        Logger.LogErrorMessage(String.Format("DeleteListEventReceiver() failed for Event Receiver [{0}] on list [{1}] of web {2}; Error=List not Found.", eventName, hostId, webUrl), false);
                        return;
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
                            Logger.LogSuccessMessage(String.Format("Deleted LIST Event Receiver [{0}] from list [{1}] on web {2}", eventName, list.Title, webUrl), false);
                            return;
                        }
                    }
                    Logger.LogErrorMessage(String.Format("DeleteListEventReceiver() failed for Event Receiver [{0}] on list [{1}] of web {2}; Error=Event Receiver not Found.", eventName, list.Title, webUrl), false);
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.ToLower().Contains("list does not exist"))
                {
                    Logger.LogErrorMessage(String.Format("DeleteListEventReceiver() failed for Event Receiver [{0}] on list [{1}] of web {2}; Error=List not Found.", eventName, hostId, webUrl), false);
                    return;
                }
                Logger.LogErrorMessage(String.Format("DeleteListEventReceiver() failed for Event Receiver [{0}] on list [{1}] of web {2}; Error={3}", eventName, hostId, webUrl, ex.Message), false);
            }
        }
        private static bool ShowInformation()
        {
            bool doContinue = false;
            string option = string.Empty;
            System.Console.WriteLine(Constants.EventReceiversInput + " file needs to be present in current working directory (where JDP.Remediation.Console.exe is present) for EventReceivers cleanup ");
            System.Console.WriteLine("Please make sure you verify the data before executing Clean-up option as cleaned EventReceivers can't be rollback.");
            System.Console.WriteLine("Press 'y' to proceed further. Press any key to go for Clean-Up Menu.");
            option = System.Console.ReadLine().ToLower();
            if (option.Equals("y", StringComparison.OrdinalIgnoreCase))
                doContinue = true;
            return doContinue;
        }
    }
}
