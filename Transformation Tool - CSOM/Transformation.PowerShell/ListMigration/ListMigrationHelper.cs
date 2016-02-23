using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using Transformation.PowerShell.Common;
using Transformation.PowerShell.Common.Utilities;
using Transformation.PowerShell.Common.CSV;
using System.IO;

namespace Transformation.PowerShell.ListMigration
{
    public class ListMigrationHelper
    {
        private void ListMigration_Initialization(string DiscoveryUsage_OutPutFolder)
        {
            //Excception CSV Creation Command
            ExceptionCsv objException = ExceptionCsv.CurrentInstance;
            objException.CreateLogFile(DiscoveryUsage_OutPutFolder);

            ExceptionCsv.WebApplication = Constants.NotApplicable;
            ExceptionCsv.SiteCollection = Constants.NotApplicable;
            ExceptionCsv.WebUrl = Constants.NotApplicable;

            //Trace Log TXT File Creation Command
            Logger objTraceLogs = Logger.CurrentInstance;
            objTraceLogs.CreateLogFile(DiscoveryUsage_OutPutFolder);
            //Trace Log TXT File Creation Command

            FileUtility.DeleteFiles(DiscoveryUsage_OutPutFolder + @"\" + Constants.ListMigration_Output);
        }
        private static List GetListByTitle(ClientContext clientContext, string ListTitle, string ActionType = "OLDLIST")
        {
            try
            {
                var list = clientContext.Web.Lists.GetByTitle(ListTitle);
                clientContext.Load(list);
                clientContext.ExecuteQuery();

                return list;
            }
            catch (Exception ex)
            {
                if (ActionType != "NEWLIST")
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] GetListByTitle. Exception Message: " + ex.Message);
                    ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "ListMigration", ex.Message, ex.ToString(), "GetListByTitle", ex.GetType().ToString(), String.Empty);
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("[Exception] GetListByTitle. Exception Message: " + ex.Message);
                    Console.ForegroundColor = ConsoleColor.Gray;
                }

                return null;
            }
        }
        private static List GetListByID(ClientContext clientContext, string ListID, string ActionType = "OLDLIST")
        {
            try
            {
                Guid List_GUID = new Guid(ListID);
                var list = clientContext.Web.Lists.GetById(List_GUID);
                clientContext.Load(list);
                clientContext.ExecuteQuery();

                return list;
            }
            catch (Exception ex)
            {
                if (ActionType != "NEWLIST")
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] GetListByID. Exception Message: " + ex.Message);
                    ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "ListMigration", ex.Message, ex.ToString(), "GetListByID", ex.GetType().ToString(), String.Empty);
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("[Exception] GetListByID. Exception Message: " + ex.Message);
                    Console.ForegroundColor = ConsoleColor.Gray;
                }

                return null;
            }
        }
        public void ListMigration_UsingCSV(string old_ListTitle, string old_ListID, string new_ListTitle, string ListUsageFilePath, string OutPutFolder, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "NA", string Password = "NA", string Domain = "NA")
        {
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                ListMigration_Initialization(OutPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## List Migration - Trasnformation Utility Execution Started - Using List Template Usage CSV ##############");
                Console.WriteLine("############## List Migration - Trasnformation Utility Execution Started - Using List Template Usage CSV ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: ListMigration_UsingCSV");
                Console.WriteLine("[START] ::: ListMigration_UsingCSV");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ListMigration_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + OutPutFolder);
                Console.WriteLine("[ListMigration_UsingCSV] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + OutPutFolder);

                //Reading Input File
                IEnumerable<ListMigrationInput> objLMInput;
                ListMigration_ReadInputCSV(old_ListTitle, old_ListID, new_ListTitle, ListUsageFilePath, OutPutFolder, out objLMInput, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
                //Reading Input File

                Console.WriteLine("[END] ::: ListMigration_UsingCSV");
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: ListMigration_UsingCSV");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## List Migration - Trasnformation Utility Execution Completed - Using List Template Usage CSV ##############");
                Console.WriteLine("############## List Migration - Trasnformation Utility Execution Completed - Using List Template Usage CSV ##############");
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] [ListMigration_UsingCSV]. Exception Message: " + ex.Message);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "ListMigrationUsingCSV", ex.Message, ex.ToString(), "ListMigration_UsingCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception] [ListMigration_UsingCSV]. Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        private void ListMigration_ReadInputCSV(string old_ListTitle, string old_ListID, string new_ListTitle, string ListUsageFilePath, string outPutFolder, out IEnumerable<ListMigrationInput> objLMInput, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "NA", string Password = "NA", string Domain = "NA")
        {
            string exceptionCommentsInfo1 = string.Empty;

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: ListMigration_ReadInputCSV");
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ListMigration_ReadInputCSV] [START] Calling function ImportCsv.ReadMatchingColumns<ListMigrationInput. Input CSV file is available at " + ListUsageFilePath);
            Console.WriteLine("[ListMigration_ReadInputCSV] [START] Calling function ImportCsv.ReadMatchingColumns<ListMigrationInput. Input CSV file is available at " + ListUsageFilePath);

            objLMInput = null;
            //objLMInput = ImportCsv.Read<ListMigrationInput>(ListUsageFilePath, Transformation.PowerShell.Common.Constants.CsvDelimeter);
            objLMInput = ImportCsv.ReadMatchingColumns<ListMigrationInput>(ListUsageFilePath, Transformation.PowerShell.Common.Constants.CsvDelimeter);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ListMigration_ReadInputCSV] [END] Read all the List records from Input and saved in List - out IEnumerable<ListMigrationInput> objLMInput, for processing.");
            Console.WriteLine("[ListMigration_ReadInputCSV] [END] Read all the List records from Input and saved in List - out IEnumerable<ListMigrationInput> objLMInput, for processing.");

            try
            {
                if (objLMInput.Any())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ListMigration_ReadInputCSV - After Loading InputCSV ");
                    Console.WriteLine("[START] ListMigration_ReadInputCSV - After Loading InputCSV");
                    bool headerCSVColumns = false;

                    //This is for Exception Comments:
                    exceptionCommentsInfo1 = "old_ListID: " + old_ListID + ", old_ListTitle: " + old_ListTitle + " new_ListTitle: " + new_ListTitle;
                    //This is for Exception Comments:

                    //Filter - List Using ListTitle
                    if (old_ListTitle != "")
                    {
                        objLMInput = from p in objLMInput
                                     where p.ListTitle == old_ListTitle
                                     select p;
                    }
                    //Filter - List Using ListId
                    else if (old_ListID != "")
                    {
                        objLMInput = from p in objLMInput
                                     where p.ListId == old_ListID
                                     select p;
                    }

                    foreach (ListMigrationInput objInput in objLMInput)
                    {
                        ExceptionCsv.WebUrl = objInput.WebUrl;
                        ExceptionCsv.SiteCollection = objInput.SiteCollectionUrl;
                        ExceptionCsv.WebApplication = objInput.WebApplicationUrl;

                        List<ListMigrationBase> objLM_CSOMBase = ListMigration_ForWEB(outPutFolder, objInput.WebUrl, objInput.ListTitle, objInput.ListId, new_ListTitle, Constants.ActionType_CSV, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                        if (objLM_CSOMBase != null)
                        {
                            if (objLM_CSOMBase.Count > 0)
                            {
                                FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.ListMigration_Output, ref objLM_CSOMBase, ref headerCSVColumns);
                            }
                        }
                    }

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ListMigration_ReadInputCSV - After Loading InputCSV");
                    Console.WriteLine("[END] ListMigration_ReadInputCSV - After Loading InputCSV");
                }
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] [ListMigration_ReadInputCSV]. Exception Message: " + ex.Message + " ExceptionComments: " + exceptionCommentsInfo1);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "ListMigrationUsingCSV", ex.Message, ex.ToString(), "ListMigration_ReadInputCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [ListMigration_ReadInputCSV] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: ListMigration_ReadInputCSV");
        }
        public List<ListMigrationBase> ListMigration_ForWEB(string outPutFolder, string WebUrl, string old_ListTitle, string old_ListID, string new_ListTitle, string ActionType = "", string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "NA", string Password = "NA", string Domain = "NA")
        {
            bool headerCSVColumns = false;
            string exceptionCommentsInfo1 = string.Empty;
            List<ListMigrationBase> objListBase = new List<ListMigrationBase>();

            ExceptionCsv.WebUrl = WebUrl;

            if (ActionType.ToLower().Trim() == Constants.ActionType_Web.ToLower())
            {
                ListMigration_Initialization(outPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## List Migration - Trasnformation Utility Execution Started - For Web ##############");
                Console.WriteLine("############## List Migration - Trasnformation Utility Execution Started - For Web ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: ListMigration_ForWEB");
                Console.WriteLine("[START] ::: ListMigration_ForWEB");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ListMigration_ForWEB] Initiated Logger and Exception Class. Logger and Exception file will be available in path " + outPutFolder);
                Console.WriteLine("[ListMigration_ForWEB] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ListMigration_ForWEB] WebUrl is " + WebUrl);
                Console.WriteLine("[ListMigration_ForWEB] WebUrl is " + WebUrl);
            }

            try
            {
                AuthenticationHelper ObjAuth = new AuthenticationHelper();
                ClientContext clientContext = null;
                exceptionCommentsInfo1 = "old_ListTitle: " + old_ListTitle + ", old_ListID: " + old_ListID + ", new_ListTitle: " + new_ListTitle;

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.OnPremise)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ListMigration_ForWEB] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(WebUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ListMigration_ForWEB] GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.Online)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ListMigration_ForWEB] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(WebUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ListMigration_ForWEB] GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                }

                if (clientContext != null)
                {
                    Web web = clientContext.Web;

                    List objOldList = null;

                    // Old List is present or not in this web ??
                    if (old_ListTitle.ToString() != "")
                    {
                        objOldList = GetListByTitle(clientContext, old_ListTitle);
                    }
                    else if (old_ListID.ToString() != "")
                    {
                        objOldList = GetListByID(clientContext, old_ListID);
                    }

                    if (objOldList != null)
                    {
                        //We found the old List in this Context
                        clientContext.Load(objOldList);
                        clientContext.ExecuteQuery();

                        //[START] Check if the New List does not exist yet
                        var objNewList = GetListByTitle(clientContext, new_ListTitle, "NEWLIST");
                        if (objNewList != null)
                        {
                            // New List exists already, no further action required
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ListMigration_ForWEB] New List " + new_ListTitle + " is already exists in this Web: " + WebUrl);
                            Console.WriteLine("[ListMigration_ForWEB] New List " + new_ListTitle + " is already exists in this Web: " + WebUrl);
                            return null;
                        }
                        //[END] Check if the New List does not exist yet

                        //New List Creation
                        List newList = null;

                        //Create List or Library
                        int intListTemplateType = -1;
                        intListTemplateType = Get_ListTemplateID(objOldList.BaseType.ToString(), (int)objOldList.BaseTemplate);
                        
                        //If List Template ID is Valid
                        if(intListTemplateType >=0)
                        {
                            ListCreationInformation creationInformation = new ListCreationInformation();
                            creationInformation.Title = new_ListTitle;
                            creationInformation.TemplateType = intListTemplateType;
                            
                            newList = clientContext.Web.Lists.Add(creationInformation);
                            clientContext.ExecuteQuery();

                            // Add Columns in List/Library
                            AddField_Using_FieldInternalDetails(clientContext, objOldList, newList);

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ListMigration_ForWEB] - List/Library " + new_ListTitle + " is created. BaseTemplate is " + objOldList.BaseType.ToString() + " and TemplateType(OLD) = " + (int)objOldList.BaseTemplate + ", TemplateType(NEW) = " + intListTemplateType);
                            Console.WriteLine("[ListMigration_ForWEB] - List/Library " + new_ListTitle + " is created. BaseTemplate is " + objOldList.BaseType.ToString() + " and TemplateType(OLD) = " + (int)objOldList.BaseTemplate + ", TemplateType(NEW) = " + intListTemplateType);

                            Replace_List_and_Library(clientContext, objOldList, newList);

                            newList.Description = objOldList.Description;
                            newList.LastItemModifiedDate = objOldList.LastItemModifiedDate;
                            
                            newList.Update();
                            clientContext.ExecuteQuery();
                        }
                        else
                        {
                             //Invalid List Template
                        }
                        
                        // Write Output in CSV = After Creation and Migration of New List
                        if (newList != null)
                        {
                            clientContext.Load(newList);
                            clientContext.ExecuteQuery();

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ListMigration_ForWEB] New List Title :" + newList.Title + ", ListID: " + newList.Id.ToString() + " created successfully in Web: " + WebUrl);
                            Console.WriteLine("[ListMigration_ForWEB] New List Title:" + newList.Title + ", ListID: " + newList.Id.ToString() + " created successfully in Web: " + WebUrl);

                            ListMigrationBase objLMBase = new ListMigrationBase();
                            objLMBase.Old_ListTitle = objOldList.Title;
                            objLMBase.Old_ListID = objOldList.Id.ToString();
                            objLMBase.Old_ListBaseTemplate = objOldList.BaseTemplate.ToString();

                            objLMBase.New_ListTitle = newList.Title.ToString();
                            objLMBase.New_ListID = newList.Id.ToString();
                            objLMBase.New_ListBaseTemplate = newList.BaseTemplate.ToString();

                            objLMBase.WebUrl = WebUrl;
                            objLMBase.SiteCollection = Constants.NotApplicable;
                            objLMBase.WebApplication = Constants.NotApplicable;

                            objListBase.Add(objLMBase);
                        }
                        else
                        {
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ListMigration_ForWEB] The new list " + new_ListTitle + " is not created.");
                            Console.WriteLine("[ListMigration_ForWEB] The new list " + new_ListTitle + " is not created.");
                        }

                        //If ==> This is for WEB
                        if (ActionType.ToString().ToLower() == Constants.ActionType_Web.ToLower())
                        {
                            if (objListBase != null)
                            {
                                FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.ListMigration_Output, ref objListBase,
                                    ref headerCSVColumns);

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ListMigration_ForWEB] Writing the List Migration Report(Output CSV) file after creating/migrating a new list. Output CSV Path: " + outPutFolder + @"\" + Constants.ListMigration_Output);
                                Console.WriteLine("[ListMigration_ForWEB] Writing the List Migration Report(Output CSV) file after creating/migrating a new list. Output CSV Path: " + outPutFolder + @"\" + Constants.ListMigration_Output);

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: ListMigration_ForWEB for WebUrl: " + WebUrl);
                                Console.WriteLine("[END] ::: ListMigration_ForWEB for WebUrl: " + WebUrl);

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## List Migration - Trasnformation Utility Execution Completed for Web ##############");
                                Console.WriteLine("############## List Migration - Trasnformation Utility Execution Completed for Web ##############");
                            }
                        }
                        //If ==> This is for WEB
                    }
                    else
                    {
                        //Old List does not present in this Context
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[ListMigration_ForWEB] Old List: " + old_ListTitle + " does not exists in this Web: " + WebUrl);
                        Console.WriteLine("[ListMigration_ForWEB] Old List: " + old_ListTitle + " does not exists in this Web: " + WebUrl);
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "ListMigration", ex.Message, ex.ToString(), "ListMigration_ForWEB", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ListMigration_ForWEB] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [ListMigration_ForWEB] Exception Message: " + ex.Message + ", ExceptionComments: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            return objListBase;
        }
        public void Replace_List_and_Library(ClientContext clientContext, List oldList, List newList)
        {
            SetListSettings(clientContext, oldList, newList);

            SetContentTypes(clientContext, oldList, newList);

            AddViews(clientContext, oldList, newList);

            //This function is not required now....
            /*RemoveViews(clientContext, oldList, newList);*/

            if (oldList.BaseType.ToString().Equals("GenericList"))
            {
                //Discussion Forum
                if(oldList.BaseTemplate.ToString().Equals("108"))
                {
                    MigreateContent_DiscussionBoard(clientContext, oldList, newList);
                }
                // All Other List
                else
                {
                    MigrateContent_List(clientContext, oldList, newList);
                }
            }
            else if (oldList.BaseType.ToString().Equals("DocumentLibrary"))
            {
                MigrateContent_Library(clientContext, oldList, newList);
            }
        }
        private static void MigrateContent_List(ClientContext clientContext, List listToBeReplaced, List newList)
        {
            try
            {
                ListItemCollection sourceListItems = listToBeReplaced.GetItems(CamlQuery.CreateAllItemsQuery());
                FieldCollection sourceListFields = listToBeReplaced.Fields;

                clientContext.Load(sourceListItems, sListItems => sListItems.IncludeWithDefaultProperties(li => li.AttachmentFiles));
                clientContext.Load(sourceListFields);
                clientContext.ExecuteQuery();
                
                var sourceItemEnumerator = sourceListItems.GetEnumerator();
                while (sourceItemEnumerator.MoveNext())
                {
                    var sourceItem = sourceItemEnumerator.Current;
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem targetItem = newList.AddItem(itemCreateInfo);
                    object sourceModifiedDate = null;
                    object sourceModifiledBy = null;
                    
                    foreach (Field sourceListField in sourceListFields)
                    {
                        try
                        {
                            //[START]Copy all except Attachments,ReadOnlyField,ContentType
                            if (!sourceListField.ReadOnlyField && sourceListField.InternalName != "Attachments" && sourceListField.InternalName != "ContentType" && null != sourceItem[sourceListField.InternalName])
                            {
                                //[START] Calendar and Event List
                                if (listToBeReplaced.BaseTemplate.ToString().Equals("106"))
                                {
                                    if (sourceListField.InternalName.Equals("EndDate"))
                                    {
                                        continue;
                                    }
                                    else if (sourceListField.InternalName.Equals("EventDate"))
                                    {
                                        targetItem[sourceListField.InternalName] = sourceItem[sourceListField.InternalName];
                                        targetItem["EndDate"] = sourceItem["EndDate"];
                                        targetItem.Update();
                                        clientContext.ExecuteQuery();
                                        //[START] [Load "Target Items" After Update, to avoid Version Conflict]
                                        targetItem = newList.GetItemById(targetItem.Id);
                                        clientContext.Load(targetItem);
                                        clientContext.ExecuteQuery();
                                        //[END] [Load "Target Items" After Update, to avoid Version Conflict]
                                    }
                                    else if (ContainsField(newList, sourceListField.InternalName))
                                    {
                                        targetItem[sourceListField.InternalName] = sourceItem[sourceListField.InternalName];
                                        targetItem.Update();
                                        clientContext.ExecuteQuery();
                                        //[START] [Load "Target Items" After Update, to avoid Version Conflict]
                                        targetItem = newList.GetItemById(targetItem.Id);
                                        clientContext.Load(targetItem);
                                        clientContext.ExecuteQuery();
                                        //[END] [Load "Target Items" After Update, to avoid Version Conflict]
                                    }
                                }
                                //[END] Calendar and Event List
                                else
                                {
                                    if (ContainsField(newList, sourceListField.InternalName))
                                    {
                                        targetItem[sourceListField.InternalName] = sourceItem[sourceListField.InternalName];
                                        targetItem.Update();
                                        clientContext.ExecuteQuery();
                                        //[START] [Load "Target Items" After Update, to avoid Version Conflict]
                                        targetItem = newList.GetItemById(targetItem.Id);
                                        clientContext.Load(targetItem);
                                        clientContext.ExecuteQuery();
                                        //[END] [Load "Target Items" After Update, to avoid Version Conflict]
                                    }
                                }
                            }
                            //[END]Copy all except Attachments, ReadOnlyField, ContentType

                            //Created, Author Field
                            if (sourceItem.FieldValues.Keys.Contains(sourceListField.InternalName))
                            {
                                //Created By
                                if (sourceListField.InternalName.Equals("Author"))
                                {
                                    //newList.Fields.GetByInternalNameOrTitle("Author").ReadOnlyField = false;
                                    if (ContainsField(newList, sourceListField.InternalName))
                                    {
                                        targetItem[sourceListField.InternalName] = sourceItem[sourceListField.InternalName];
                                        targetItem.Update();
                                        clientContext.ExecuteQuery();
                                    }
                                }
                                //Created Date
                                if (sourceListField.InternalName.Equals("Created"))
                                {
                                    //newList.Fields.GetByInternalNameOrTitle("Created").ReadOnlyField = false;
                                    if (ContainsField(newList, sourceListField.InternalName))
                                    {
                                        targetItem[sourceListField.InternalName] = sourceItem[sourceListField.InternalName];
                                        targetItem.Update();
                                        clientContext.ExecuteQuery();
                                    }
                                }
                                
                                //Modified Date
                                if (sourceListField.InternalName.Equals("Modified"))
                                {
                                    if (ContainsField(newList, sourceListField.InternalName))
                                    {
                                        sourceModifiedDate = sourceItem["Modified"];
                                    }
                                }
                                //Modified By
                                if (sourceListField.InternalName.Equals("Editor"))
                                {
                                    if (ContainsField(newList, sourceListField.InternalName))
                                    {
                                        sourceModifiledBy = sourceItem["Editor"];
                                    }
                                }
                            }
                            //Created, Author Field
                        }
                        catch (Exception ex)
                        {
                            ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "ListMigration", ex.Message, ex.ToString(), "MigrateContent_List - Copy Items", ex.GetType().ToString(), "Not initialized: " + sourceListField.Title);
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [MigrateContent_List]  Copy Items, Exception Message: " + ex.Message + ", Exception Comment: Not initialized: " + sourceListField.Title);

                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("[EXCEPTION] [MigrateContent_List]  Copy Items, Exception Message: " + ex.Message + ", ExceptionComments: Not initialized: " + sourceListField.Title);
                            Console.ForegroundColor = ConsoleColor.Gray;
                        }
                    }

                    #region Copy Attachments

                    //Copy attachments
                   foreach (Attachment fileName in sourceItem.AttachmentFiles)
                    {
                        try
                        {
                            Microsoft.SharePoint.Client.File oAttachment = clientContext.Web.GetFileByServerRelativeUrl(fileName.ServerRelativeUrl);
                            clientContext.Load(oAttachment);
                            clientContext.ExecuteQuery();

                            FileInformation fInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, oAttachment.ServerRelativeUrl);
                            AttachmentCreationInformation attachFileInfo = new AttachmentCreationInformation();

                            Byte[] buffer = new Byte[oAttachment.Length];
                            int bytesRead = fInfo.Stream.Read(buffer, 0, buffer.Length);

                            MemoryStream stream = new MemoryStream(buffer);
                            attachFileInfo.ContentStream = stream;
                            attachFileInfo.FileName = oAttachment.Name;
                            targetItem.AttachmentFiles.Add(attachFileInfo);
                            targetItem.Update();
                            clientContext.ExecuteQuery();
                            stream.Dispose();

                            //[START] [Load "Target Items" After Update, to avoid Version Conflict]
                            targetItem = newList.GetItemById(targetItem.Id);
                            clientContext.Load(targetItem);
                            clientContext.ExecuteQuery();
                            //[END] [Load "Target Items" After Update, to avoid Version Conflict]
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Equals("Version conflict.", StringComparison.CurrentCultureIgnoreCase))
                            {                              
                                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "ListMigration", ex.Message, ex.ToString(), "MigrateContent_List - Copy Attachments", ex.GetType().ToString(), "");
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [MigrateContent_List] Copy Attachments, Exception Message: " + ex.Message);

                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("[EXCEPTION] [MigrateContent_List] Copy Attachments, Exception Message: " + ex.Message);
                                Console.ForegroundColor = ConsoleColor.Gray;
                            }
                        }
                    }
                    #endregion                 
                   
                   //[START] [Load "Target Items" After Update, to avoid Version Conflict]
                   targetItem = newList.GetItemById(targetItem.Id);
                   clientContext.Load(targetItem);
                   clientContext.ExecuteQuery();
                   //[END] [Load "Target Items" After Update, to avoid Version Conflict]

                   targetItem["Modified"] = sourceModifiedDate;
                   targetItem["Editor"] = sourceModifiledBy;
                   targetItem.Update();
                   clientContext.ExecuteQuery();

                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "ListMigration", ex.Message, ex.ToString(), "MigrateContent_List", ex.GetType().ToString(), "");
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [MigrateContent_List] Exception Message: " + ex.Message);

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [MigrateContent_List] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        private static void MigrateContent_Library(ClientContext clientContext, List listToBeReplaced, List newList)
        {
            try
            {
                ListItemCollection items = listToBeReplaced.GetItems(CamlQuery.CreateAllItemsQuery());
                Folder destination = newList.RootFolder;
                Folder source = listToBeReplaced.RootFolder;

                FieldCollection sourceListFields = listToBeReplaced.Fields;
                clientContext.Load(sourceListFields);

                clientContext.Load(destination,
                                    d => d.ServerRelativeUrl);
                clientContext.Load(source,
                                    s => s.Files,
                                    s => s.ServerRelativeUrl);
                clientContext.Load(items,
                                    i => i.IncludeWithDefaultProperties(item => item.File));
                clientContext.ExecuteQuery();


                foreach (Microsoft.SharePoint.Client.File file in source.Files)
                {
                    string newUrl = file.ServerRelativeUrl.Replace(source.ServerRelativeUrl, destination.ServerRelativeUrl);
                    file.CopyTo(newUrl, true);

                    ListItemCollection newListItems = newList.GetItems(CamlQuery.CreateAllItemsQuery());
                    clientContext.Load(destination,
                                    d => d.Files,
                                    d => d.ServerRelativeUrl);
                    clientContext.Load(newListItems,
                                        i => i.IncludeWithDefaultProperties(item => item.File));
                    clientContext.ExecuteQuery();

                    object sourceModifiedDate = null;

                    foreach (ListItem newListItem in newListItems)
                    {
                        if (newListItem.File.Name.Equals(file.Name, StringComparison.CurrentCultureIgnoreCase))
                        {
                            foreach (Field sourceListField in sourceListFields)
                            {                                
                                if (sourceListField.InternalName.Equals("Modified"))
                                {
                                    if (ContainsField(newList, sourceListField.InternalName))
                                    {
                                        sourceModifiedDate = file.TimeLastModified;
                                    }
                                }
                                
                                if (sourceListField.InternalName.Equals("Editor"))
                                {
                                    if (ContainsField(newList, sourceListField.InternalName))
                                    {
                                        newListItem[sourceListField.InternalName] = file.ModifiedBy;
                                        newListItem.Update();
                                        clientContext.Load(newListItem);
                                        clientContext.ExecuteQuery();
                                    }
                                }
                            }
                            newListItem["Modified"] = sourceModifiedDate;
                            newListItem.Update();
                            clientContext.Load(newListItem);
                            clientContext.ExecuteQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "ListMigration", ex.Message, ex.ToString(), "MigrateContent_Library", ex.GetType().ToString(), "");
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [MigrateContent_Library] Exception Message: " + ex.Message);

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION] [MigrateContent_Library] Exception Message: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        private static void MigreateContent_DiscussionBoard(ClientContext clientContext, List listToBeReplaced, List newList)
        {
            // Source Discussion Board List
            CamlQuery Query_SourceTopics = CamlQuery.CreateAllFoldersQuery();
            ListItemCollection Source_Topics = listToBeReplaced.GetItems(Query_SourceTopics);
            clientContext.Load(Source_Topics);
            clientContext.ExecuteQuery();

            foreach (ListItem Source_Topic in Source_Topics)
            {
                ListItem Target_Topic = Microsoft.SharePoint.Client.Utilities.Utility.CreateNewDiscussion(clientContext, newList, Source_Topic["Title"].ToString());

                Target_Topic["Body"] = Source_Topic["Body"];
                Target_Topic["Author"] = Source_Topic["Author"];
                Target_Topic["Created"] = Source_Topic["Created"];
                Target_Topic["Modified"] = Source_Topic["Modified"];

                Target_Topic.Update();
                clientContext.ExecuteQuery();

                // Target Discussion Board List
                CamlQuery Query_TargetTopics = CamlQuery.CreateAllFoldersQuery();
                ListItemCollection Target_Topics = newList.GetItems(Query_TargetTopics);
                clientContext.Load(Target_Topics);
                clientContext.ExecuteQuery();

                foreach (ListItem Target_Topics_Item in Target_Topics)
                {
                    if ((Source_Topic["Title"].ToString()).Equals(Target_Topics_Item["Title"].ToString()))
                    {
                        Query_SourceTopics = CamlQuery.CreateAllItemsQuery();
                        Query_SourceTopics.FolderServerRelativeUrl = Source_Topic["FileRef"].ToString();
                        Query_SourceTopics.ViewXml = "<View Scope='RecursiveAll'></View>";

                        //Updating/Loading the target list
                        ListItemCollection SourceTopics_Replies = listToBeReplaced.GetItems(Query_SourceTopics);
                        clientContext.Load(SourceTopics_Replies);
                        clientContext.ExecuteQuery();

                        //Copying the responses....
                        foreach (ListItem SourceTopics_reply in SourceTopics_Replies)
                        {
                            ListItem TargetTopics_reply = Microsoft.SharePoint.Client.Utilities.Utility.CreateNewDiscussionReply(clientContext, Target_Topics_Item);
                            TargetTopics_reply["Body"] = SourceTopics_reply["Body"];
                            TargetTopics_reply["Created"] = SourceTopics_reply["Created"];
                            TargetTopics_reply["Modified"] = SourceTopics_reply["Modified"];
                            TargetTopics_reply["Author"] = SourceTopics_reply["Author"];
                            TargetTopics_reply["ParentFolderId"] = SourceTopics_reply["ParentFolderId"];

                            TargetTopics_reply.Update();
                            clientContext.ExecuteQuery();
                        }
                    }
                }
            }
        }
        private static void AddField_Using_FieldInternalDetails(ClientContext clientContext, List listToBeReplaced, List NewList)
        {
            FieldCollection fields = listToBeReplaced.Fields;
            clientContext.Load(fields);
            clientContext.ExecuteQuery();

            foreach (Field _Fld in fields)
            {
                string _IName = _Fld.InternalName;
                if (!Field_ISAlreadyExists(NewList, _IName, clientContext))
                {
                    Field _Updated_SiteColumnField = NewList.Fields.Add(_Fld);

                    clientContext.Load(_Updated_SiteColumnField);
                    clientContext.ExecuteQuery();
                }
            }
        }

        #region <<<Not In Use>>> - AddField_Using_FieldSchemaXml and Field_UpdateSchemaXml
        //Not In Use - AddField_Using_FieldSchemaXml
        public static void AddField_Using_FieldSchemaXml(ClientContext clientContext, List listToBeReplaced, List NewList)
        {
            FieldCollection fields = listToBeReplaced.Fields;
            clientContext.Load(fields);
            clientContext.ExecuteQuery();

            foreach (Field _Fld in fields)
            {
                string _IName = _Fld.InternalName;

                /*Console.WriteLine("InternalName: " + _Fld.InternalName);
                Console.WriteLine("Title: " + _Fld.Title);
                Console.WriteLine("StaticName: " + _Fld.StaticName);
                Console.WriteLine("SchemaXml: " + _Fld.SchemaXml);
                Console.WriteLine("**************************");*/

                if (!Field_ISAlreadyExists(NewList, _IName, clientContext))
                {
                    string _ListColumnNewSchemaXml = Field_UpdateSchemaXml(_Fld, _Fld.InternalName, _Fld.Title);

                    if (_ListColumnNewSchemaXml != null)
                    {
                        Field _Updated_SiteColumnField = NewList.Fields.AddFieldAsXml(_ListColumnNewSchemaXml, true, AddFieldOptions.DefaultValue);
                        clientContext.Load(_Updated_SiteColumnField);
                        clientContext.ExecuteQuery();
                    }
                }
            }
        }
        //Not In Use - Field_UpdateSchemaXml
        #endregion

        public static string Field_UpdateSchemaXml(Field _ListColumn, string newListColumn_InternalName, string newListColumn_DisplayName)
        {
            string FieldAsXML = _ListColumn.SchemaXml.ToString();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(FieldAsXML);
            XmlNode node = xmlDoc.SelectSingleNode("Field");
            XmlElement Xe = (XmlElement)node;

            Xe.SetAttribute("DisplayName", newListColumn_DisplayName);
            Xe.SetAttribute("StaticName", newListColumn_InternalName);
            Xe.SetAttribute("Name", newListColumn_InternalName);

            FieldAsXML = xmlDoc.InnerXml;

            return FieldAsXML;
        }
        private static bool Field_ISAlreadyExists(List newList, string Field_InternalName, ClientContext clientContext)
        {
            FieldCollection collFields = newList.Fields;
            clientContext.Load(collFields);
            clientContext.ExecuteQuery();

            // Check existing fields
            bool _IsAlreadyExists = false;

            if (Field_InternalName != "")
            {
                var fieldExists = collFields.Any(f => f.InternalName.Trim().ToLower() == Field_InternalName.Trim().ToLower());

                if (fieldExists)
                {
                    _IsAlreadyExists = true;
                    //Console.WriteLine("This Column " + Field_InternalName + " is already exist");
                }
            }
            else
            {
                _IsAlreadyExists = true;
                Console.WriteLine("Please enter correct list column detail.");
            }

            return _IsAlreadyExists;
        }
        private static bool ContainsField(List list, string fieldName)
        {
            var ctx = list.Context;
            var result = ctx.LoadQuery(list.Fields.Where(f => f.InternalName == fieldName));
            ctx.ExecuteQuery();
            return result.Any();
        }

        #region Migrate Views

        private static void RemoveDefaultViews_From_NewlyCreatedList(ClientContext clientContext, List listToBeReplaced, List newList)
        {
            //Update the list of views
            
            //Source
            clientContext.Load(listToBeReplaced, l => l.Views);
            
            //Target
            clientContext.Load(newList, l => l.Views);
            clientContext.ExecuteQuery();

            var viewsToRemove = new List<View>();
            foreach (View view in newList.Views)
            {
                if (listToBeReplaced.Views.Any(v => v.Title == view.Title))
                {
                    //new list contains a view which is not on the source list, remove it
                    viewsToRemove.Add(view);
                }
            }
            foreach (View view in viewsToRemove)
            {
                view.DeleteObject();
            }
            newList.Update();
            clientContext.ExecuteQuery();
        }
        private static void AddViews(ClientContext clientContext, List listToBeReplaced, List newList)
        {           
            //Delete All Views from New List
            RemoveDefaultViews_From_NewlyCreatedList(clientContext, listToBeReplaced, newList);
            //Delete All Views from New List

            ViewCollection views = listToBeReplaced.Views;
            clientContext.Load(views,
                                v => v.Include(view => view.Paged,
                                    view => view.PersonalView,
                                    view => view.ViewQuery,
                                    view => view.Title,
                                    view => view.RowLimit,
                                    view => view.DefaultView,
                                    view => view.ViewFields,
                                    view => view.Hidden,
                                    view => view.ViewType,
                                    view => view.DefaultView,
                                    view => view.MobileDefaultView));
            clientContext.ExecuteQuery();

            //Build a list of views which only exist on the source list
            var viewsToCreate = new List<ViewCreationInformation>();
            foreach (View view in listToBeReplaced.Views)
            {
                var createInfo = new ViewCreationInformation
                {
                    Paged = view.Paged,
                    PersonalView = view.PersonalView,
                    Query = view.ViewQuery,
                    Title = view.Title,
                    RowLimit = view.RowLimit,
                    SetAsDefaultView = view.DefaultView,
                    ViewFields = view.ViewFields.ToArray(),
                    ViewTypeKind = GetViewType(view.ViewType)

                };
                //ViewCreationInformation createInfo = new ViewCreationInformation();
                //createInfo.Title = view.Title;
                //createInfo.Paged = view.Paged;
                //createInfo.PersonalView = view.PersonalView;
                //createInfo.RowLimit = view.RowLimit;
                //createInfo.ViewFields = view.ViewFields.ToArray();
                //createInfo.ViewTypeKind = GetViewType(view.ViewType);
                //createInfo.SetAsDefaultView = view.DefaultView;
                //createInfo.Query = view.ViewQuery;
                
               if (!view.Hidden)
                    viewsToCreate.Add(createInfo);
            }

            //Create the list that we need to
            foreach (ViewCreationInformation newView in viewsToCreate)
            {
                newList.Views.Add(newView);
            }

            newList.Update();

            UpdateViewProperties(clientContext, listToBeReplaced, newList);
        }
        private static void RemoveViews(ClientContext clientContext, List listToBeReplaced, List newList)
        {
            //Update the list of views
            clientContext.Load(newList, l => l.Views);
            clientContext.ExecuteQuery();

            var viewsToRemove = new List<View>();
            foreach (View view in newList.Views)
            {
                if (!listToBeReplaced.Views.Any(v => v.Title == view.Title))
                {
                    //new list contains a view which is not on the source list, remove it
                    viewsToRemove.Add(view);
                }
            }
            foreach (View view in viewsToRemove)
            {
                view.DeleteObject();
            }
            newList.Update();
            clientContext.ExecuteQuery();
        }
        private static ViewType GetViewType(string viewType)
        {
            switch (viewType)
            {
                case "HTML":
                    return ViewType.Html;
                case "GRID":
                    return ViewType.Grid;
                case "CALENDAR":
                    return ViewType.Calendar;
                case "RECURRENCE":
                    return ViewType.Recurrence;
                case "CHART":
                    return ViewType.Chart;
                case "GANTT":
                    return ViewType.Gantt;
                default:
                    return ViewType.None;
            }
        }
        private static void UpdateViewProperties(ClientContext clientContext, List listToBeReplaced, List newList)
        {
            ViewCollection views_listToBeReplaced = listToBeReplaced.Views;
            clientContext.Load(views_listToBeReplaced);
            clientContext.ExecuteQuery();

            ViewCollection views_newList = newList.Views;
            clientContext.Load(views_newList);
            clientContext.ExecuteQuery();

            foreach (View oldView in views_listToBeReplaced)
            {
                foreach (View newView in views_newList)
                {
                    if (newView.Title.Equals(oldView.Title))
                    {
                        newView.Scope = oldView.Scope;
                        newView.Hidden = oldView.Hidden;
                        newView.ViewData = oldView.ViewData;

                        //This Will Not Work
                        //newView.ListViewXml = oldView.ListViewXml;
                        //newView.Aggregations = oldView.Aggregations;
                        //newView.AggregationsStatus = oldView.AggregationsStatus;
                        
                        /*
                        Console.WriteLine("ServerRelativeUrl" + oldView.ServerRelativeUrl);
                        Console.WriteLine("StyleId" + oldView.StyleId);
                        Console.WriteLine("Threaded" + oldView.Threaded);
                        Console.WriteLine("Toolbar" + oldView.Toolbar);
                        Console.WriteLine("ViewFields" + oldView.ViewFields);
                        Console.WriteLine("Scope" + oldView.Scope);
                        Console.WriteLine("ModerationType" + oldView.ModerationType);
                        Console.WriteLine("ListViewXml" + oldView.ListViewXml);
                        Console.WriteLine("HtmlSchemaXml" + oldView.HtmlSchemaXml);
                        Console.WriteLine("Formats" + oldView.Formats);
                        Console.WriteLine("Aggregations" + oldView.Aggregations);
                        string _OP = "ServerRelativeUrl: " + oldView.ServerRelativeUrl + " StyleId: " + oldView.StyleId + " Threaded: " + oldView.Threaded + " Toolbar: " + oldView.Toolbar
                            + " ViewFields: " + oldView.ViewFields + " Scope: " + oldView.Scope + " ModerationType: " + oldView.ModerationType + " ListViewXml: " + oldView.ListViewXml
                            + " HtmlSchemaXml: " + oldView.HtmlSchemaXml + " Formats: " + oldView.Formats + " Aggregations" + oldView.Aggregations + " NEW ViEW ListViewXml" + newView.ListViewXml;
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[VIEW]" + oldView.Title +" Details: " + _OP);*/

                        newView.Update();
                        clientContext.ExecuteQuery();
                    }
                }
            }
        }

        #endregion

        private static void SetContentTypes(ClientContext clientContext, List listToBeReplaced, List newList)
        {
            clientContext.Load(listToBeReplaced,
                        l => l.ContentTypesEnabled,
                        l => l.ContentTypes);
            clientContext.Load(newList,
                                l => l.ContentTypesEnabled,
                                l => l.ContentTypes);
            clientContext.ExecuteQuery();

            //If the originat list doesn't use ContentTypes there's nothing to do here.
            if (!listToBeReplaced.ContentTypesEnabled) return;

            newList.ContentTypesEnabled = true;
            newList.Update();
            clientContext.ExecuteQuery();

            foreach (var contentType in listToBeReplaced.ContentTypes)
            {
                if (!newList.ContentTypes.Any(ct => ct.Name == contentType.Name))
                {
                    //current Content Type needs to be added to new list
                    //Note that the Parent is used as contentType is the list instance not the site instance.
                    newList.ContentTypes.AddExistingContentType(contentType.Parent);
                    newList.Update();
                    clientContext.ExecuteQuery();


                }
            }
            //We need to re-load the ContentTypes for newList as they may have changed due to an add call above
            clientContext.Load(newList, l => l.ContentTypes);
            clientContext.ExecuteQuery();

            //Remove any content type that are not needed
            var contentTypesToDelete = new List<ContentType>();
            foreach (var contentType in newList.ContentTypes)
            {
                if (!listToBeReplaced.ContentTypes.Any(ct => ct.Name == contentType.Name))
                {
                    //current Content Type needs to be removed from new list
                    contentTypesToDelete.Add(contentType);
                }
            }
            foreach (var contentType in contentTypesToDelete)
            {
                contentType.DeleteObject();
            }
            newList.Update();
            clientContext.ExecuteQuery();
        }
        private static void SetListSettings(ClientContext clientContext, List listToBeReplaced, List newList)
        {
            clientContext.Load(listToBeReplaced,
                        l => l.EnableVersioning,
                        l => l.EnableModeration,
                        l => l.EnableMinorVersions,
                        l => l.DraftVersionVisibility);
            clientContext.ExecuteQuery();
            newList.EnableVersioning = listToBeReplaced.EnableVersioning;
            newList.EnableModeration = listToBeReplaced.EnableModeration;
            newList.EnableMinorVersions = listToBeReplaced.EnableMinorVersions;
            newList.DraftVersionVisibility = listToBeReplaced.DraftVersionVisibility;
            newList.Update();
            clientContext.ExecuteQuery();
        }
        private static int Get_ListTemplateID(string strOldListBaseType, int OldListBaseTemplate)
        {
            int ListTemplateID = 0;

            //Generic list, Promoted Links
            if (OldListBaseTemplate == 100 || OldListBaseTemplate == 170)
            {
                ListTemplateID = 100;
            }
            //Document library
            else if (OldListBaseTemplate == 101)
            {
                ListTemplateID = 101;
            }
            //Survey
            else if (OldListBaseTemplate == 102)
            {
                ListTemplateID = 102;
            }
            //Links List
            else if (OldListBaseTemplate == 103)
            {
                ListTemplateID = 103;
            }
            //Announcements list
            else if (OldListBaseTemplate == 104)
            {
                ListTemplateID = 104;
            }
            //Contacts list
            else if (OldListBaseTemplate == 105)
            {
                ListTemplateID = 105;
            }
            //Events list
            else if (OldListBaseTemplate == 106)
            {
                ListTemplateID = 106;
            }
            //Tasks list
            else if (OldListBaseTemplate == 107)
            {
                ListTemplateID = 107;
            }
            //Discussion board
            else if (OldListBaseTemplate == 108)
            {
                ListTemplateID = 108;
            }
            //Picture library
            else if (OldListBaseTemplate == 109)
            {
                ListTemplateID = 109;
            }
            //Data sources
            else if (OldListBaseTemplate == 110)
            {
                ListTemplateID = 110;
            }
            else
            {
                if (strOldListBaseType == "GenericList")
                {
                    ListTemplateID = 100;
                }
                else if (strOldListBaseType == "DocumentLibrary")
                {
                    ListTemplateID = 101;
                }
                else
                {
                    //Invalid Template
                    ListTemplateID = -1;
                }
            }

            return ListTemplateID;
        }
    }
}
