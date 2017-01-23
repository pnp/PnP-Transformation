using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;
using PeoplePickerRemediation.Console.Common.Base;
using PeoplePickerRemediation.Console.Common.CSV;
using System.Web;
using System.Security;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using Microsoft.SharePoint.Client.EventReceivers;
using System.Xml.Linq;
using System.IO;
using PeoplePickerRemediation.Console.Common.Utilities;

namespace PeoplePickerRemediation.Console
{
    public class PeoplePickerRemediation
    {
        static int totalCount = 0;
        public static string listName = string.Empty;
        public static string webUrl = string.Empty;
        //public static ClientContext _context;
        //public static List list;
        public static Dictionary<string, string> dictUserUpns = new Dictionary<string, string>();
        public static void DoWork(string inputFileSpec)
        {
            Logger.OpenLog("PeoplePickerRemediation");
            Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);
            Logger.LogInfoMessage(inputFileSpec, true);
            List<PeoplePickerListOutput> _WriteUDCList = null;

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
            Logger.LogInfoMessage(String.Format("- AppSettings[LocalAdLdapQuery] = {0}", ConfigurationManager.AppSettings["LocalAdLdapQuery"].ToString()), true);
            Logger.LogInfoMessage(String.Format("- AppSettings[UpnPrefix] = {0}", ConfigurationManager.AppSettings["UpnPrefix"].ToString()), true);
            Logger.LogInfoMessage(String.Format("- AppSettings[UpdateUserInfoEvenIfEventReceiversEnabled] = {0}", ConfigurationManager.AppSettings["UpdateUserInfoEvenIfEventReceiversEnabled"].ToString()), true);
            Logger.LogInfoMessage(String.Format("- AppSettings[UpdateUserInfoEvenIfWorkflowsEnabled] = {0}", ConfigurationManager.AppSettings["UpdateUserInfoEvenIfWorkflowsEnabled"].ToString()), true);
            Logger.LogInfoMessage(String.Format("- AppSettings[CamlQueryRowLimit] = {0}", ConfigurationManager.AppSettings["CamlQueryRowLimit"].ToString()), true);

            IEnumerable<PeoplePickerListsInput> pprCSVRows = ImportCSV.ReadMatchingColumns<PeoplePickerListsInput>(inputFileSpec, Constants.CsvDelimeter);
            if (pprCSVRows != null)
            {
                try
                {
                    var lists = pprCSVRows.Where(x => x.ListName != null && x.WebUrl != null);
                    if (lists != null && lists.Count() > 0)
                    {
                        _WriteUDCList = new List<PeoplePickerListOutput>();
                        Logger.LogInfoMessage(String.Format("Preparing to process a total of {0} PeoplePicker InfoPath Form Libraries ...", lists.Count()), true);
                        PeoplePickerListOutput ps = new PeoplePickerListOutput();
                        foreach (PeoplePickerListsInput pprFileInput in lists)
                        {
                            ReadFormLibUsingAppOnlyAndCredentials(pprFileInput, ref _WriteUDCList);
                        }

                        if (_WriteUDCList != null && _WriteUDCList.Any())
                        {
                            GenerateStatusReport(_WriteUDCList);
                            _WriteUDCList = null;
                        }
                    }
                    else
                    {
                        Logger.LogInfoMessage("No valid PeoplePicker InfoPath Form Library records found in '" + inputFileSpec + "' File ", true);
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogErrorMessage(String.Format("PeoplePickerRemediation:DoWork() failed: Error={0}", ex.Message), true);
                }

                Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            }
            else
            {
                Logger.LogInfoMessage("No PeoplePicker InfoPath Form Library records found in '" + inputFileSpec + "' File ", true);
            }

            Logger.CloseLog();
        }
        private static void GenerateStatusReport(List<PeoplePickerListOutput> _WriteUDCList)
        {
            try
            {
                bool headerOfCsv = false;
                string csvFileName = Environment.CurrentDirectory + "\\" + Constants.PeopplePickerReportOutput;
                if (!System.IO.File.Exists(csvFileName))
                {
                    headerOfCsv = true;
                }
               
                FileUtility.DoPeriodicFlushOfListObject(ref _WriteUDCList, csvFileName, ref headerOfCsv);
                //Export the result in CSV file
                if (_WriteUDCList != null && _WriteUDCList.Any())
                {
                    FileUtility.WriteCsVintoFile(csvFileName, ref _WriteUDCList, ref headerOfCsv);
                    _WriteUDCList = null;
                };
                //headerOfCsv = false;
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("Error while writing data to OutPut File--Message:" + ex.Message, false);
            }
        }

        public static void ReadFormLibUsingAppOnlyAndCredentials(PeoplePickerListsInput formLib, ref List<PeoplePickerListOutput> lstPeoplepickeroutput)
        {
            if (formLib == null)
                return;

            webUrl = formLib.WebUrl;
            listName = formLib.ListName;


            Logger.LogInfoMessage(String.Format("Processing PeoplePicker InfoPath Form Library [{0}] of Web [{1}] ...", listName, webUrl), true);

            CamlQuery camlQuery = new CamlQuery();
            //Set View Scope for the Query
            camlQuery.SetViewAttribute(QueryScope.RecursiveAll);
            //Or Set the ViewFields xml
            //camlQuery.SetViewFields(@"<FieldRef Name='ID'/><FieldRef Name='Title'/>");
            //Override the QueryThrottle Mode for avoiding ListViewThreshold exception
            camlQuery.SetQueryThrottleMode(QueryThrottleMode.Override);
            //Use OrderBy ID field if Query doesn't have filter with indexed column
            camlQuery.SetOrderByIDField();
            //Set RowLimit
            camlQuery.SetQueryRowlimit(Convert.ToUInt32(ConfigurationManager.AppSettings["CamlQueryRowLimit"]));
            PeoplePickerListOutput peoplePickerOutput = new PeoplePickerListOutput();
            peoplePickerOutput.WebUrl = webUrl;
            peoplePickerOutput.ListName = formLib.ListName;
            try
            {
                using (ClientContext userContext = Helper.CreateClientContextBasedOnAuthMode(Program.UseAppModel, Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, webUrl))
                {
                    List list = userContext.Web.Lists.GetByTitle(listName);
                    userContext.Load(list);
                    userContext.Load(list.EventReceivers);
                    userContext.Load(list.WorkflowAssociations);
                    userContext.ExecuteQueryRetry();

                    // By default, we will process the specified library unless we find that it has event receiver registrations or workflow associations. 
                    // We are about to update the files of this library and we do not want to cause side effects that might arise from triggering a event receiver or workflow.
                    bool processTheLibrary = true;

                    if (list.EventReceivers.Count > 0)
                    {
                        // We found some event receivers on the library.  Consult the app setting to determine if we should still process the library.
                        if (String.Equals(ConfigurationManager.AppSettings["UpdateUserInfoEvenIfEventReceiversEnabled"].ToString(), "Yes", StringComparison.InvariantCultureIgnoreCase) == true)
                        {
                            Logger.LogWarningMessage(String.Format("[{0}] EventReceivers are associated with List [{1}] of Web [{2}]; event receivers might be executed", list.EventReceivers.Count, listName, webUrl), true);
                        }
                        else
                        {
                            // The admin has not configured this utility to process libraries that have event receivers
                            processTheLibrary = false;
                            Logger.LogWarningMessage(String.Format("[{0}] EventReceivers are associated with List [{1}] of Web [{2}]; skipping the list per the AppSetting", list.EventReceivers.Count, listName, webUrl), true);
                        }
                    }

                    if (list.WorkflowAssociations.Count > 0)
                    {
                        // We found some workflow associations on the library.  Consult the app setting to determine if we should still process the library.
                        if (String.Equals(ConfigurationManager.AppSettings["UpdateUserInfoEvenIfWorkflowsEnabled"].ToString(), "Yes", StringComparison.InvariantCultureIgnoreCase) == true)
                        {
                            Logger.LogWarningMessage(String.Format("[{0}] Workflows are associated with List [{1}] of Web [{2}]; workflows might be started", list.WorkflowAssociations.Count, listName, webUrl), true);
                        }
                        else
                        {
                            // The admin has not configured this utility to process libraries that have workflow associations
                            processTheLibrary = false;
                            Logger.LogWarningMessage(String.Format("[{0}] Workflows are associated with List [{1}] of Web [{2}]; skipping the list per the AppSetting", list.WorkflowAssociations.Count, listName, webUrl), true);
                        }
                    }

                    // Process the library if it is still OK to do so.
                    if (processTheLibrary)
                    {
                        // TODO: uncomment below lines to switch read-only property
                        //PeoplePickerEdiorModified modified = new PeoplePickerEdiorModified();
                        //UpdateEditorAndModifiedFieldsProperty(userContext, list, listName, ref modified, false);

                        ContentIterator contentIterator = new ContentIterator(userContext);
                        try
                        {

                            contentIterator.ProcessListItem(listName, camlQuery, ProcessItem, ref lstPeoplepickeroutput,
                                delegate(ListItem item, System.Exception ex)
                                {
                                    return true;
                                });

                            Logger.LogInfoMessage(String.Format("[{0}] InfoPath Form files processed for List [{1}] of Web [{2}]", totalCount, listName, webUrl), true);
                        }
                        catch (Exception ex)
                        {
                            Logger.LogErrorMessage(ex.Message);
                        }

                        // TODO: uncomment below lines to switch read-only property
                        //if (modified.isEditorfieldModified || modified.isModifiedFieldModified)
                        //{
                        //    UpdateEditorAndModifiedFieldsProperty(userContext, list, listName, ref modified, true);
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ReadFormLibUsingAppOnlyAndCredentials() failed for PeroplePicker InfoPath Library [{0}] of Web [{1}]: Error={2}", listName, webUrl, ex.Message), true);
                peoplePickerOutput.Status = Constants.ErrorStatus;
                peoplePickerOutput.ErrorDetails = ex.Message;
                lstPeoplepickeroutput.Add(peoplePickerOutput);
            }
        }

        private static void UpdateEditorAndModifiedFieldsProperty(ClientContext userContext, List list, string listName, ref PeoplePickerEditorModified isEditorModified, bool reset)
        {
            try
            {
                if (list == null)
                {
                    list = userContext.Web.Lists.GetByTitle(listName);
                    Web oweb = userContext.Web;
                    userContext.Load(oweb);
                    userContext.Load(list);
                    userContext.ExecuteQueryRetry();
                }
                FieldCollection fields = list.Fields;
                userContext.Load(fields, flds => flds.Where(field => field.InternalName == "Editor" || field.InternalName == "Modified"));
                userContext.ExecuteQueryRetry();

                foreach (Field oField in fields)
                {
                    if (!reset)
                    {
                        if (oField.InternalName == "Editor" && !oField.ReadOnlyField)
                        {
                            oField.ReadOnlyField = true;
                            isEditorModified.isEditorfieldModified = true;
                            oField.Update();
                        }

                        if (oField.InternalName == "Modified" && !oField.ReadOnlyField)
                        {
                            oField.ReadOnlyField = true;
                            isEditorModified.isModifiedFieldModified = true;
                            oField.Update();
                        }
                    }
                    else
                    {
                        //reset to original status

                        if (oField.InternalName == "Editor" && isEditorModified.isEditorfieldModified)
                        {
                            oField.ReadOnlyField = !oField.ReadOnlyField;
                            oField.Update();
                        }
                        if (oField.InternalName == "Modified" && isEditorModified.isModifiedFieldModified)
                        {
                            oField.ReadOnlyField = !oField.ReadOnlyField;
                            oField.Update();
                        }
                    }

                }
                userContext.ExecuteQueryRetry();


            }
            catch (Exception Ex)
            {
                Logger.LogErrorMessage("Failed Updating readonly fileds in Method  UpdateEditorAndModifiedFieldsProperty() Exception Message--"+Ex.Message, false);
            }
        }

        public static void ProcessItem(ListItem item, ClientContext _context, ref List<PeoplePickerListOutput> lstPeoplepickeroutput)
        {
            PeoplePickerListOutput Peoplepickeroutput = new PeoplePickerListOutput();

            try
            {
                totalCount++;
                Logger.LogInfoMessage("item id : " + item.Id);
                //Get Web
                Web web = _context.Web;
                _context.Load(web);
                _context.Load(item);
                _context.Load(item.ParentList);
                _context.Load(item.File);
                _context.Load(item.Folder);
                _context.ExecuteQueryRetry();

                Peoplepickeroutput.ListName = item.ParentList.Title;
                Peoplepickeroutput.WebUrl = web.Url;
                Peoplepickeroutput.ItemID =Convert.ToString(item.Id);
                StringBuilder usersString = new StringBuilder();
                StringBuilder groupsString = new StringBuilder();
                int itemID = item.Id;
                FieldUserValue fuLastModifiedUser = (FieldUserValue)item["Editor"];
                string lastModifiedUser = fuLastModifiedUser.LookupValue;
                DateTime lastModifiedTimeStamp = Convert.ToDateTime(item["Modified"].ToString());

                XDocument xmlDoc = new XDocument();
                string fileName = item.File.Name;
                string fileServerRelativeUrl = item.File.ServerRelativeUrl;
                string folderServerRelativeUrl = fileServerRelativeUrl.Substring(0, fileServerRelativeUrl.LastIndexOf("/"));
                string upnPrefix = ConfigurationManager.AppSettings["UpnPrefix"].ToString();
                // Approach to read File contents depends on Auth Model chosen
                if (Program.UseAppModel == true)
                {
                    string originalFileContents = SafeGetFileAsString(web, item.File.ServerRelativeUrl);
                    if (String.IsNullOrEmpty(originalFileContents))
                    {
                        Logger.LogErrorMessage(String.Format("Could not get contents of InfoPath Form File [{0}]", item.File.ServerRelativeUrl), false);
                        return;
                    }

                    xmlDoc = XDocument.Load(new StringReader(originalFileContents));
                }
                else
                {
                    FileInformation fileInformation = Microsoft.SharePoint.Client.File.OpenBinaryDirect(_context, item.File.ServerRelativeUrl);
                    xmlDoc = XDocument.Load(XmlReader.Create(fileInformation.Stream));
                }
               
                bool fileupdated = false;

                Logger.LogInfoMessage(String.Format("Got InfoPath Form File [{0}] from [{1}] Library", fileName, listName), true);

                // convert XDocument into string
                string xmlData = xmlDoc.ToString();
                //Logger.LogInfoMessage(String.Format("Xml Data {0}", xmlData), true);
                string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>";
                xmlData = xml + xmlData;
                //Get top nodes
                string xmlTopData = xmlData.Substring(0, xmlData.IndexOf("<my"));
                if (xmlData.ToLower().Contains("<?mso-infopathsolution"))
                {
                    XmlDocument xDoc = new XmlDocument();
                    xDoc.LoadXml(xmlData);
                    //XPathNavigator class provides a set of methods used to modify nodes and values in an XML document
                    XPathNavigator navItem = xDoc.CreateNavigator();
                    //Redefine NameSpaceManager (Generate the NameSpace manager for this item)
                    navItem.MoveToFollowing(XPathNodeType.Element);
                    XmlNamespaceManager nsManager = new XmlNamespaceManager(new NameTable());
                    foreach (var ns in navItem.GetNamespacesInScope(XmlNamespaceScope.All))
                    {
                        if (ns.Key == String.Empty)
                        {
                            nsManager.AddNamespace("def", ns.Value);
                        }
                        else
                        {
                            nsManager.AddNamespace(ns.Key, ns.Value);
                        }
                    }

                    XPathNodeIterator nodes = navItem.Select("//pc:Person", nsManager);
                    foreach (XPathNavigator node in nodes)
                    {
                        try
                        {
                            XPathNavigator nav = node.SelectSingleNode("pc:DisplayName", nsManager);
                            string displayName = nav == null ? string.Empty : nav.Value;
                            nav = node.SelectSingleNode("pc:AccountId", nsManager);
                            string accountId = nav == null ? string.Empty : nav.Value;
                            nav = node.SelectSingleNode("pc:AccountType", nsManager);
                            string accountType = nav == null ? string.Empty : nav.Value;

                            string samaccountname = string.Empty;

                            if (accountType.ToLower().Contains("user"))
                            {
                                if (accountId.Contains(@"\"))
                                {
                                    samaccountname = accountId.Substring(accountId.LastIndexOf(@"\") + 1);
                                    usersString.Append(samaccountname);
                                }
                            }
                            else if (accountType.ToLower().Contains("group"))
                            {
                                Logger.LogInfoMessage(String.Format("Found group: {0}", accountId), true);
                                //Peoplepickeroutput.AccountType = "Group";

                                // TO DO: Convert group id into MT compatiable 
                                if (accountId.Contains("|"))
                                {
                                    samaccountname = accountId.Substring(accountId.LastIndexOf(@"|") + 1);
                                    groupsString.Append(samaccountname).Append(Constants.OutPutreportSeparator);
                                }
                            }

                            if (!accountId.ToLower().Contains(upnPrefix) && !string.IsNullOrEmpty(samaccountname))
                            {
                                if (dictUserUpns.ContainsKey(samaccountname))
                                {
                                    string upn = dictUserUpns[samaccountname];
                                    node.SelectSingleNode("pc:AccountId", nsManager).SetValue(upn);
                                    fileupdated = true;
                                    //Peoplepickeroutput.NewUPN = upn;
                                    usersString.Append(Constants.OutPutreportSeparator).Append(upn).Append(Constants.OutPutreportSeparator);
                                }
                                else
                                {
                                    try
                                    {
                                        string upn = GetUPN(accountType, samaccountname);
                                        string newUPN = string.Empty;

                                        if (!string.IsNullOrEmpty(upn))
                                        {
                                            newUPN = upnPrefix + upn;
                                            dictUserUpns.Add(samaccountname, newUPN);
                                            node.SelectSingleNode("pc:AccountId", nsManager).SetValue(newUPN);
                                            fileupdated = true;
                                            //Peoplepickeroutput.NewUPN = newUPN;
                                            usersString.Append(Constants.OutPutreportSeparator).Append(newUPN).Append(Constants.OutPutreportSeparator);
                                        }
                                        else
                                        {
                                            Logger.LogInfoMessage(String.Format("UPN not found for [{0}]", samaccountname), true);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Logger.LogErrorMessage(String.Format("ProcessItem() failed for InfoPath Form Library [{0}] of Web [{1}]: Reason={2}; Error={3}", listName, web.Url,
                                    "LDAP Query Processing not successful.",
                                    "[" + ex.Message + "] | [" + ex.HResult + "] | [" + ex.Source + "] | [" + ex.StackTrace + "] | [" + ex.TargetSite + "]"), false);
                                        lstPeoplepickeroutput.Add(LogExceptionMessages(Peoplepickeroutput, ex));
                                        return;
                                    }

                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.LogErrorMessage(String.Format("ProcessItem() failed for InfoPath Form Library [{0}] of Web [{1}]: Reason={2}; Error={3}", listName, webUrl,
                        "failed while iterating nodes.",
                        "[" + ex.Message + "] | [" + ex.HResult + "] | [" + ex.Source + "] | [" + ex.StackTrace + "] | [" + ex.TargetSite + "]"), false);
                            lstPeoplepickeroutput.Add(LogExceptionMessages(Peoplepickeroutput, ex));
                            return;
                        }
                    }
                    if (usersString != null && !string.IsNullOrEmpty(Convert.ToString(usersString)))
                    {
                        Peoplepickeroutput.Users = Convert.ToString(usersString).TrimEnd(';');
                    }
                    if (groupsString != null && !string.IsNullOrEmpty(Convert.ToString(groupsString)))
                    {
                        Peoplepickeroutput.Groups = Convert.ToString(groupsString).TrimEnd(';');
                    }
                    if (fileupdated)
                    {
                        string updateXml = xmlTopData + navItem.OuterXml;
                        string fileUrl = string.Empty;

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            StreamWriter writer = new StreamWriter(memoryStream);
                            writer.Write(updateXml);
                            writer.Flush();
                            memoryStream.Position = 0;
                            Logger.LogInfoMessage(String.Format("Saving contents of InfoPath Form File [{0}] of Web [{1}] ...", fileName, webUrl), false);
                            // Approach to save File contents depends on Auth Model chosen
                            if (Program.UseAppModel == true)
                            {
                                Folder targetFolder = null;

                                Logger.LogInfoMessage(String.Format("Getting folder of InfoPath Form File [{0}] of Web [{1}] ...", fileName, webUrl), false);
                                try
                                {
                                    targetFolder = web.GetFolderByServerRelativeUrl(folderServerRelativeUrl);
                                    web.Context.Load(targetFolder);
                                    web.Context.ExecuteQueryRetry();
                                    Logger.LogInfoMessage(String.Format("Got folder"), false);
                                }
                                catch (Exception ex)
                                {
                                    Logger.LogErrorMessage(String.Format("ProcessItem() failed for InfoPath Form File [{0}] in the Form Library {1}] of Web [{2}]: Reason={3}; Error={4}", fileName, listName, webUrl,
                                        "Upload Folder was not Found.",
                                        "[" + ex.Message + "] | [" + ex.HResult + "] | [" + ex.Source + "] | [" + ex.StackTrace + "] | [" + ex.TargetSite + "]"), false);
                                    lstPeoplepickeroutput.Add(LogExceptionMessages(Peoplepickeroutput, ex));
                                    return;
                                }

                                Logger.LogInfoMessage(String.Format("Uploading InfoPath Form File [{0}] ...", fileName), false);
                                try
                                {
                                    targetFolder.UploadFile(fileName, memoryStream, true);
                                    Logger.LogInfoMessage(String.Format("Uploaded file"), false);
                                }
                                catch (Exception ex)
                                {
                                    Logger.LogErrorMessage(String.Format("ProcessItem() failed for InfoPath Form File [{0}] in the Form Library {1}] of Web [{2}]: Reason={3}; Error={4}", fileName, listName, webUrl,
                                        "File Upload Failed.",
                                        "[" + ex.Message + "] | [" + ex.HResult + "] | [" + ex.Source + "] | [" + ex.StackTrace + "] | [" + ex.TargetSite + "]"), false);
                                    lstPeoplepickeroutput.Add(LogExceptionMessages(Peoplepickeroutput, ex));
                                    return;
                                }
                            }
                            else
                            {
                                try
                                {
                                    fileUrl = String.Format("{0}/{1}", folderServerRelativeUrl, string.Format(fileName, lastModifiedTimeStamp.ToString()));
                                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(_context, fileUrl, memoryStream, true);
                                }
                                catch (Exception ex)
                                {
                                    Logger.LogErrorMessage(String.Format("ProcessItem() failed for InfoPath Form File [{0}] in the Form Library {1}] of Web [{2}]: Reason={3}; Error={4}", fileName, listName, webUrl,
                                        "File Upload Failed.",
                                        "[" + ex.Message + "] | [" + ex.HResult + "] | [" + ex.Source + "] | [" + ex.StackTrace + "] | [" + ex.TargetSite + "]"), false);
                                    lstPeoplepickeroutput.Add(LogExceptionMessages(Peoplepickeroutput, ex));
                                    return;
                                }
                            }
                            Logger.LogInfoMessage(String.Format("Saved contents of InfoPath Form File [{0}] in FormLibrary [{1}] of Web [{2}]", fileUrl, listName, webUrl), true);

                            try
                            {
                                _context.Load(item);
                                _context.ExecuteQueryRetry();

                                item["Editor"] = fuLastModifiedUser;
                                item["Modified"] = lastModifiedTimeStamp.ToString();
                                item.Update();
                                _context.ExecuteQueryRetry();
                                Peoplepickeroutput.Status = Constants.SuccessStatus;
                                Peoplepickeroutput.ErrorDetails = Constants.NotApplicable;
                            }
                            catch (Exception ex)
                            {
                                Logger.LogErrorMessage(String.Format("ProcessItem() failed for InfoPath Form File [{0}] in the Form Library {1}] of Web [{2}]: Reason={3}; Error={4}", fileName, listName, webUrl,
                                "Editor & Modified field update failed.",
                                "[" + ex.Message + "] | [" + ex.HResult + "] | [" + ex.Source + "] | [" + ex.StackTrace + "] | [" + ex.TargetSite + "]"), false);
                                lstPeoplepickeroutput.Add(LogExceptionMessages(Peoplepickeroutput, ex));
                                return;
                            }
                            //Logger.LogInfoMessage(String.Format("Updated Item with previous Editor [{0}] and Modified Timestamp [{1}] in FormLibrary [{2}] of Web [{3}]",
                            //    lastModifiedUser, lastModifiedTimeStamp.ToString(), listName, webUrl), true);
                        }
                        
                    }
                    else
                    {
                        Peoplepickeroutput.Status = Constants.NoUpdateRequired;
                        Peoplepickeroutput.ErrorDetails = Constants.NotApplicable;
                    }
                    lstPeoplepickeroutput.Add(Peoplepickeroutput);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ProcessItem() failed for InfoPath Form Library [{0}] of Web [{1}]: Reason={2}; Error={3}", listName, webUrl,
            "failed while processing item.",
            "[" + ex.Message + "] | [" + ex.HResult + "] | [" + ex.Source + "] | [" + ex.StackTrace + "] | [" + ex.TargetSite + "]"), false);
                return;
            }


        }
        public static PeoplePickerListOutput LogExceptionMessages(PeoplePickerListOutput peoplepickerOutputException, Exception ex)
        {
            peoplepickerOutputException.Status = Constants.ErrorStatus;
            peoplepickerOutputException.ErrorDetails = ex.Message;

            return peoplepickerOutputException;
        }
        public static bool ValidateLDAPVariable()
        {
            string ldapQuery = ConfigurationManager.AppSettings["LocalAdLdapQuery"].ToString();

            if (string.IsNullOrEmpty(ldapQuery))
            {
                System.Console.WriteLine("The required AppSettings[LocalAdLdapQuery] element is empty or null");
                return false;
            }

            return true;
        }

        public static string GetUserPrinicpalNameFromDirectorySearcher(string accountId)
        {
            string samaccountname = string.Empty;

            if (accountId.Contains(@"\"))
            {
                samaccountname = accountId.Substring(accountId.LastIndexOf(@"\") + 1);
            }

            return GetUPN("user", samaccountname);
        }

        private static string GetUPN(string accountType, string samaccountname)
        {
            string ldapQuery = ConfigurationManager.AppSettings["LocalAdLdapQuery"].ToString();

            // Bind to the users container.
            DirectoryEntry entry = new DirectoryEntry(ldapQuery);
            // Create a DirectorySearcher object.
            DirectorySearcher mySearcher = new DirectorySearcher(entry);
            // Create a SearchResultCollection object to hold a collection of SearchResults
            // returned by the FindAll method.
            mySearcher.PageSize = 500;  // ADD THIS LINE HERE !
            string strFilter = string.Empty;
            if (accountType.ToLower().Equals("user"))
                strFilter = string.Format("(&(objectCategory=User)(SAMAccountName={0}))", samaccountname);
            else if (accountType.ToLower().Contains("group"))
                strFilter = string.Format("(&(objectCategory=Group)(sid={0}))", samaccountname);
            var propertiesToLoad = new[] { "SAMAccountName", "userprincipalname", "sid" };
            mySearcher.PropertiesToLoad.AddRange(propertiesToLoad);
            mySearcher.Filter = strFilter;
            mySearcher.CacheResults = false;
            SearchResultCollection result = mySearcher.FindAll();

            if (result != null && result.Count > 0)
            {
                return GetProperty(result[0], "userprincipalname");
            }

            return string.Empty;
        }

        private static string GetProperty(SearchResult searchResult, string PropertyName)
        {
            if (searchResult.Properties.Contains(PropertyName))
            {
                return searchResult.Properties[PropertyName][0].ToString();
            }
            else
            {
                return string.Empty;
            }
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
    public class PeoplePickerEditorModified
    {
        public bool isEditorfieldModified { get; set; }
        public bool isModifiedFieldModified { get; set; }
    }
}
