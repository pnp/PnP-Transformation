using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Transformation.PowerShell.Common;
using Transformation.PowerShell.Common.CSV;
using Transformation.PowerShell.Common.Utilities;

namespace Transformation.PowerShell.MasterPage
{
    public class MasterPageHelper
    {
       /// <summary>
       /// Initialized of Exception and Logger Class. Deleted the Master Page Replace Usage File
       /// </summary>
        public void MasterPage_Initialization(string DiscoveryUsage_OutPutFolder)
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

            //Delete MasterPage Replace OUTPUT File
            DeleteMasterPage_ReplaceOutPutFiles(DiscoveryUsage_OutPutFolder);

        }

        public void ChangeMasterPageForDiscoveryOutPut(string DiscoveryUsage_OutPutFolder, string MasterPageUsagePath, string New_MasterPageDetails = "N/A", string Old_MasterPageDetails = "N/A", string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;

            try
            {
                //Initialized Exception and Logger. Deleted the Master Page Replace Usage File
                MasterPage_Initialization(DiscoveryUsage_OutPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Master Page Trasnformation Utility Execution Started : Using InputCSV ##############");
                Console.WriteLine("############## Master Page Trasnformation Utility Execution Started : Using InputCSV ##############");
               
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());
                
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] FUNCTION ::: ChangeMasterPageForDiscoveryOutPut");
                Console.WriteLine("[START] FUNCTION ::: ChangeMasterPageForDiscoveryOutPut");
                
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForDiscoveryOutPut] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + DiscoveryUsage_OutPutFolder);
                Console.WriteLine("[ChangeMasterPageForDiscoveryOutPut] Initiated Logger and Exception Class. Logger and Exception file will be available at path: " + DiscoveryUsage_OutPutFolder);
                
                //Reading Master Page Input File
                IEnumerable<MasterPageInput> objMPInput;
                ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV(DiscoveryUsage_OutPutFolder, MasterPageUsagePath, out objMPInput, New_MasterPageDetails, Old_MasterPageDetails, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
                
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] FUNCTION ::: ChangeMasterPageForDiscoveryOutPut");
                Console.WriteLine("[END] FUNCTION ::: ChangeMasterPageForDiscoveryOutPut");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Master Page Trasnformation Utility Execution Completed : Using InputCSV ##############");
                Console.WriteLine("############## Master Page Trasnformation Utility Execution Completed : Using InputCSV ##############");
            }
            catch (Exception ex)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] ChangeMasterPageForDiscoveryOutPut. Exception Message: " + ex.Message);
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "MasterPage", ex.Message, ex.ToString(), "ChangeMasterPageForDiscoveryOutPut", ex.GetType().ToString());

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception] ChangeMasterPageForDiscoveryOutPut. Exception Message:" + ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        public void ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV(string outPutFolder, string MasterPageUsagePath, out IEnumerable<MasterPageInput> objMPInput, string New_MasterPageDetails = "N/A", string Old_MasterPageDetails = "N/A", string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;
            List<MasterPageBase> _WriteMasterList = new List<MasterPageBase>();

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ::: ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV");
            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV] [START] Calling function ImportCsv.ReadMatchingColumns<MasterPageInput>. Master Page Input CSV file is available at " + MasterPageUsagePath);
            Console.WriteLine("[ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV] [START] Calling function ImportCsv.ReadMatchingColumns<MasterPageInput>. Master Page Input CSV file is available at " + MasterPageUsagePath);

            objMPInput = null;
            //objMPInput = ImportCsv.Read<MasterPageInput>(outPutFolder + @"\" + Transformation.PowerShell.Common.Constants.MasterPageInput, Transformation.PowerShell.Common.Constants.CsvDelimeter);
            objMPInput = ImportCsv.ReadMatchingColumns<MasterPageInput>(MasterPageUsagePath, Transformation.PowerShell.Common.Constants.CsvDelimeter);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV] [END] Read all the INPUT from Master Page and saved in List - out IEnumerable<MasterPageInput> objMpInput, for processing.");
            Console.WriteLine("[ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV] [END] Read all the INPUT from Master Page and saved in List - out IEnumerable<MasterPageInput> objMpInput, for processing.");

            try
            {
                if (objMPInput.Any())
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV - After Loading InputCSV ");
                    Console.WriteLine("[START] ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV - After Loading InputCSV");
                    
                    bool headerMasterPage = false;

                    foreach (MasterPageInput objInput in objMPInput)
                    {
                        //This is for Exception Comments:
                        ExceptionCsv.WebUrl = objInput.WebUrl;
                        ExceptionCsv.SiteCollection = objInput.SiteCollectionUrl;
                        ExceptionCsv.WebApplication = objInput.WebApplicationUrl;
                        exceptionCommentsInfo1 = "<Input>New MasterPage Url = " + New_MasterPageDetails + ", <Input> OLD MasterUrl: " + Old_MasterPageDetails + ", WebUrl: " + objInput.WebUrl + ", CustomMasterUrlStatus" + objInput.CustomMasterUrlStatus + "MasterUrlStatus" + objInput.MasterUrlStatus;
                        //This is for Exception Comments:

                        MasterPageBase objMPBase = new MasterPageBase();
                        //if (Old_MasterPageDetails.Trim().ToLower() != Constants.Input_Blank && Old_MasterPageDetails.Trim().ToLower() != Constants.Input_All)
                        if (Old_MasterPageDetails.Trim().ToLower() != Constants.Input_All)
                        {
                            objMPBase = ChangeMasterPageForWeb(outPutFolder, objInput.WebUrl, New_MasterPageDetails, Old_MasterPageDetails, Convert.ToBoolean(objInput.CustomMasterUrlStatus), Convert.ToBoolean(objInput.MasterUrlStatus), Constants.ActionType_CSV, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
                        }
                        else
                        {
                            objMPBase = ChangeMasterPageForWeb(outPutFolder, objInput.WebUrl, New_MasterPageDetails, Old_MasterPageDetails, true, Convert.ToBoolean(objInput.MasterUrlStatus), Constants.ActionType_CSV, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
                        }

                        if (objMPBase != null)
                        { _WriteMasterList.Add(objMPBase); }
                    }

                    if (_WriteMasterList != null)
                    {
                        if (_WriteMasterList.Count > 0)
                        {
                            FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.MasterPageUsage, ref _WriteMasterList, ref headerMasterPage);
                        }
                    }

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV - After Loading InputCSV. Updated the Master Pages.");
                    Console.WriteLine("[END] ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV - After Loading InputCSV. Updated the Master Pages.");
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "MasterPage", ex.Message, ex.ToString(), "ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[Exception] ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV. Exception Message:" + ex.Message);

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception] ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV. Exception Message:" + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] ::: ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV");
            Console.WriteLine("[END] ::: ChangeMasterPageForDiscoveryOutPut_ReadMasterPagesCSV");
        }

        public void ChangeMasterPageForSiteCollection(string outPutFolder, string SiteCollectionUrl, string NewMasterPageURL, string OldMasterPageURL = "N/A", bool CustomMasterUrlStatus = true, bool MasterUrlStatus = true, string SharePointOnline_OR_OnPremise = "N/A", string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            string exceptionCommentsInfo1 = string.Empty;
            List<MasterPageBase> _WriteMasterList = new List<MasterPageBase>();
            //Initialized Exception and Logger. Deleted the Master Page Replace Usage File

            MasterPage_Initialization(outPutFolder);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Master Page Trasnformation Utility Execution Started - For Site Collection ##############");
            Console.WriteLine("############## Master Page Trasnformation Utility Execution Started - For Site Collection ##############");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
            Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ::: ChangeMasterPageForSiteCollection");
            Console.WriteLine("[START] ENTERING IN FUNCTION ::: ChangeMasterPageForSiteCollection");

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForSiteCollection] Initiated Logger and Exception Class. Logger and Exception file will be available at path " + outPutFolder);
            Console.WriteLine("[ChangeMasterPageForSiteCollection] Initiated Logger and Exception Class. Logger and Exception file will be available at path " + outPutFolder);

            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForSiteCollection] SiteCollectionUrl is " + SiteCollectionUrl);
            Console.WriteLine("[ChangeMasterPageForSiteCollection] SiteCollectionUrl is " + SiteCollectionUrl);
                
            try
            {
                AuthenticationHelper ObjAuth = new AuthenticationHelper();
                ClientContext clientContext = null;
                
                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.OnPremise)
                {
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(SiteCollectionUrl, UserName, Password, Domain);
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.Online)
                {
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(SiteCollectionUrl, UserName, Password);
                }

                if (clientContext != null)
                {
                    bool headerMasterPage = false;
                    MasterPageBase objMPBase = new MasterPageBase();
                    
                    Web rootWeb = clientContext.Web;
                    clientContext.Load(rootWeb);
                    clientContext.ExecuteQuery();

                    //This is for Exception Comments:
                    ExceptionCsv.SiteCollection = SiteCollectionUrl;
                    ExceptionCsv.WebUrl = rootWeb.Url.ToString();
                    exceptionCommentsInfo1 = "<Input>New MasterPage Url = " + NewMasterPageURL + ", <Input> OLD MasterUrl: " + OldMasterPageURL + ", WebUrl: " + rootWeb.Url.ToString() + ", CustomMasterUrlStatus" + CustomMasterUrlStatus + "MasterUrlStatus" + MasterUrlStatus;
                    //This is for Exception Comments:

                    //Root Web
                    objMPBase = ChangeMasterPageForWeb(outPutFolder, rootWeb.Url.ToString(), NewMasterPageURL, OldMasterPageURL, CustomMasterUrlStatus, MasterUrlStatus, Constants.ActionType_SiteCollection, SharePointOnline_OR_OnPremise, UserName, Password, Domain);

                    if (objMPBase != null)
                    {
                        _WriteMasterList.Add(objMPBase);
                    }
                    WebCollection webCollection = rootWeb.Webs;      
                    clientContext.Load(webCollection);
                    clientContext.ExecuteQuery();

                    foreach (Web webSite in webCollection)
                    {
                        try
                        {
                            //This is for Exception Comments:
                            ExceptionCsv.SiteCollection = SiteCollectionUrl;
                            ExceptionCsv.WebUrl = webSite.Url.ToString();
                            exceptionCommentsInfo1 = "<Input>New MasterPage Url = " + NewMasterPageURL + ", <Input> OLD MasterUrl: " + OldMasterPageURL + ", WebUrl: " + webSite.Url.ToString() + ", CustomMasterUrlStatus" + CustomMasterUrlStatus + "MasterUrlStatus" + MasterUrlStatus;
                            //This is for Exception Comments:

                            //Web
                            objMPBase = ChangeMasterPageForWeb(outPutFolder, webSite.Url, NewMasterPageURL, OldMasterPageURL, CustomMasterUrlStatus, MasterUrlStatus, Constants.ActionType_SiteCollection, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
                            
                            if (objMPBase != null)
                            { _WriteMasterList.Add(objMPBase); }
                        }
                        catch (Exception ex)
                        {
                            ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "MasterPage", ex.Message, ex.ToString(), "ChangeMasterPageForSiteCollection", ex.GetType().ToString(), exceptionCommentsInfo1);
                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ChangeMasterPageForSiteCollection] ChangeMasterPageForSiteCollection. Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                            
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("[Exception] [ChangeMasterPageForSiteCollection]. Exception Message:" + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                            Console.ForegroundColor = ConsoleColor.Gray;
                        }
                    }

                    if (_WriteMasterList != null)
                    {
                        if (_WriteMasterList.Count > 0)
                        {
                            FileUtility.WriteCsVintoFile(outPutFolder + @"\" + Constants.MasterPageUsage, ref _WriteMasterList, ref headerMasterPage);
                        }
                    }
                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] [ChangeMasterPageForSiteCollection] EXIT FROM FUNCTION ChangeMasterPageForSiteCollection for SiteCollectionUrl: " + SiteCollectionUrl);
                Console.WriteLine("[END] [ChangeMasterPageForSiteCollection] EXIT FROM FUNCTION ChangeMasterPageForSiteCollection for SiteCollectionUrl: " + SiteCollectionUrl);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Master Page Trasnformation Utility Execution Completed - For Site Collection ##############");
                Console.WriteLine("############## Master Page Trasnformation Utility Execution Completed - For Site Collection ##############");
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "MasterPage", ex.Message, ex.ToString(), "ChangeMasterPageForSiteCollection", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION] [ChangeMasterPageForSiteCollection] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[Exception] [ChangeMasterPageForSiteCollection]. Exception Message:" + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        public MasterPageBase ChangeMasterPageForWeb(string outPutFolder, string WebUrl, string NewMasterPageURL, string OldMasterPageURL = "N/A", bool CustomMasterUrlStatus = true, bool MasterUrlStatus = true, string ActionType = Constants.ActionType_Web, string SharePointOnline_OR_OnPremise = Constants.OnPremise, string UserName = "N/A", string Password = "N/A", string Domain = "N/A")
        {
            bool headerMasterPage = false;
            List<MasterPageBase> _WriteMasterList = null;
            ExceptionCsv.WebUrl = WebUrl;

            ///<ActionType=="Web"> That means this function running only for a web. We have to write the output in this function only
            ///<Action Type=="SiteCollection"> The function will return object MasterPageBase, and consolidated output will be written in SiteCollection function - ChangeMasterPageForSiteCollection

            //If ==> This is for WEB
            if (ActionType.ToString().ToLower() == Constants.ActionType_Web.ToLower())
            {
                MasterPage_Initialization(outPutFolder);
                _WriteMasterList = new List<MasterPageBase>();

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Master Page Trasnformation Utility Execution Started - For Web ##############");
                Console.WriteLine("############## Master Page Trasnformation Utility Execution Started - For Web ##############");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + Logger.CurrentDateTime());
                Console.WriteLine("[DATE TIME] " + Logger.CurrentDateTime());

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[START] ENTERING IN FUNCTION ::: ChangeMasterPageForWeb");
                Console.WriteLine("[START] ENTERING IN FUNCTION ::: ChangeMasterPageForWeb");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb] Initiated Logger and Exception Class. Logger and Exception file will be available in path " + outPutFolder);
                Console.WriteLine("[ChangeMasterPageForWeb] Initiated Logger and Exception Class. Logger and Exception file will be available in path" + outPutFolder);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb] WebUrl is " + WebUrl);
                Console.WriteLine("[ChangeMasterPageForWeb] WebUrl is " + WebUrl);
            }
            
            string exceptionCommentsInfo1 = string.Empty;
            MasterPageBase objMaster = new MasterPageBase();
            
            try
            {
                AuthenticationHelper ObjAuth = new AuthenticationHelper();
                ClientContext clientContext = null;

                //SharePoint on-premises / SharePoint Online Dedicated => OP (On-Premises)
                if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.OnPremise)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ChangeMasterPageForWeb] ENTERING IN FUNCTION GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetNetworkCredentialAuthenticatedContext(WebUrl, UserName, Password, Domain);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ChangeMasterPageForWeb] EXIT FROM FUNCTION GetNetworkCredentialAuthenticatedContext for WebUrl: " + WebUrl);
                    
                }
                //SharePointOnline  => OL (Online)
                else if (SharePointOnline_OR_OnPremise.ToUpper() == Constants.Online)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ChangeMasterPageForWeb] ENTERING IN FUNCTION GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                    clientContext = ObjAuth.GetSharePointOnlineAuthenticatedContextTenant(WebUrl, UserName, Password);
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ChangeMasterPageForWeb] EXIT FROM FUNCTION GetSharePointOnlineAuthenticatedContextTenant for WebUrl: " + WebUrl);
                }

                if (clientContext != null)
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[START][ChangeMasterPageForWeb] ChangeMasterPageForWeb for WebUrl: " + WebUrl);
                    Console.WriteLine("[START][ChangeMasterPageForWeb] ChangeMasterPageForWeb for WebUrl: " + WebUrl);
                    Web web = clientContext.Web;
                    
                    //Load Web to get old Master Page details
                    clientContext.Load(web);
                    clientContext.ExecuteQuery();
                    //Load Web to get old Master Page details

                    //Create New Master Page Relative URL
                    string masterPageUrl = string.Empty;
                    masterPageUrl = GetMasterPageRelativeURL(clientContext, NewMasterPageURL);

                    //Create OldMasterPageURL Relative URL
                    string _strOldMasterPageURL = string.Empty;
                    if (OldMasterPageURL.Trim().ToLower() != Constants.Input_Blank && OldMasterPageURL.Trim().ToLower() != Constants.Input_All)
                    {
                        _strOldMasterPageURL = GetMasterPageRelativeURL(clientContext, OldMasterPageURL);
                    }

                    //Prepare Exception Comments
                    exceptionCommentsInfo1 = "New Master URL: " + masterPageUrl + ", OldMasterPageURL="+_strOldMasterPageURL+", CustomMasterUrlStatus: " + CustomMasterUrlStatus + ", MasterUrlStatus: " + MasterUrlStatus;

                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb]: Input Master Page URL(New) was " + NewMasterPageURL + ". After processing Master Page URL(New) is " + masterPageUrl);
                    Console.WriteLine("[ChangeMasterPageForWeb]: Input Master Page URL(New) was " + NewMasterPageURL + ". After processing Master Page URL(New) is " + masterPageUrl);
                    
                    //Check if new master page is available in Gallery
                    if (Check_MasterPageExistsINGallery(clientContext, masterPageUrl))
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb] Check_MasterPageExistsINGallery: This New Master Page is present in Gallery: " + masterPageUrl);
                        Console.WriteLine("[ChangeMasterPageForWeb] Check_MasterPageExistsINGallery: This New Master Page is present in Gallery: " + masterPageUrl);

                        //Added in Output Object <objMaster> - To Write old Master Page details
                        objMaster.OLD_CustomMasterUrl = web.CustomMasterUrl;
                        objMaster.OLD_MasterUrl = web.MasterUrl;
                        //Added in Output Object <objMaster> - To Write old Master Page details

                        //if (OldMasterPageURL.Trim().ToLower() != Constants.Input_Blank && OldMasterPageURL.Trim().ToLower() != Constants.Input_All)
                        if (OldMasterPageURL.Trim().ToLower() != Constants.Input_All)
                        {
                            bool _UpdateMasterPage = false;

                            if (CustomMasterUrlStatus && _strOldMasterPageURL.ToLower().Trim() == web.CustomMasterUrl.ToString().Trim().ToLower())
                            {
                                web.CustomMasterUrl = masterPageUrl;
                                _UpdateMasterPage = true;

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: Updated Custom Master Page " + _strOldMasterPageURL + " with new Master Page URL " + masterPageUrl);
                                Console.WriteLine("[ChangeMasterPageForWeb]:[OldMasterPageURL !=\"\"]: Updated Custom Master Page " + _strOldMasterPageURL + " with new Master Page URL " + masterPageUrl);
                            }
                            else
                            {
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: [NO Update in CustomMasterUrl] <INPUT> OLD Custom Master Page " + _strOldMasterPageURL.Trim().ToLower() + ", <WEB> OLD Master Page URL " + web.CustomMasterUrl.ToString().Trim().ToLower());
                                Console.WriteLine("[ChangeMasterPageForWeb]:[OldMasterPageURL !=\"\"]: [NO Update in CustomMasterUrl] <INPUT> OLD Custom Master Page " + _strOldMasterPageURL.Trim().ToLower() + ", <WEB> OLD Master Page URL " + web.CustomMasterUrl.ToString().Trim().ToLower());
                            }

                            if (MasterUrlStatus && _strOldMasterPageURL.ToLower().Trim() == web.MasterUrl.ToString().Trim().ToLower())
                            {
                                web.MasterUrl = masterPageUrl;
                                _UpdateMasterPage = true;

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: Updated Master Page " + _strOldMasterPageURL + " with new Master Page URL " + masterPageUrl);
                                Console.WriteLine("[ChangeMasterPageForWeb]:[OldMasterPageURL !=\"\"]: Updated Master Page " + _strOldMasterPageURL + " with new Master Page URL " + masterPageUrl);
                            }
                            else
                            {
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: [NO Update in MasterUrl] <INPUT> OLD Master Page " + _strOldMasterPageURL.Trim().ToLower() + ", <WEB> OLD Master Page URL " + web.MasterUrl.ToString().Trim().ToLower());
                                Console.WriteLine("[ChangeMasterPageForWeb]:[OldMasterPageURL !=\"\"]: [NO Update in MasterUrl] <INPUT> OLD Master Page " + _strOldMasterPageURL.Trim().ToLower() + ", <WEB> OLD Master Page URL " + web.MasterUrl.ToString().Trim().ToLower());
                            }

                            if (_UpdateMasterPage)
                            {
                                web.Update();

                                clientContext.Load(web);
                                clientContext.ExecuteQuery();

                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"] Changed Master Page for - " + WebUrl + ", New Master Page is " + masterPageUrl);
                                Console.WriteLine("[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: Changed Master Page for - " + WebUrl + ", New Master Page is " + masterPageUrl);
                            }
                            else
                            {
                                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: The <Input> OLD MasterPage does not match with this site's old <WEB> master page for WEB: " + WebUrl);
                                Console.WriteLine("[ChangeMasterPageForWeb][OldMasterPageURL !=\"\"]: The <Input> OLD MasterPage does not match with this site's old <WEB> master page for WEB: " + WebUrl);
                            }
                        }
                        else
                        {
                            if (CustomMasterUrlStatus)
                            { web.CustomMasterUrl = masterPageUrl; }

                            if (MasterUrlStatus)
                            { web.MasterUrl = masterPageUrl; }

                            //Update Web
                            web.Update();

                            //Load Web to get Updated Details
                            clientContext.Load(web);
                            clientContext.ExecuteQuery();

                            Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb][OldMasterPageURL ==\"\"] Changed Master Page for - " + WebUrl + ", New Master Page is " + masterPageUrl);
                            Console.WriteLine("[ChangeMasterPageForWeb][OldMasterPageURL ==\"\"] Changed Master Page for - " + WebUrl + ", New Master Page is " + masterPageUrl);
                        }

                        //Added in Output Object <objMaster> 
                        objMaster.CustomMasterUrl = web.CustomMasterUrl;
                        objMaster.MasterUrl = web.MasterUrl;
                        objMaster.WebApplication = Constants.NotApplicable;
                        objMaster.SiteCollection = Constants.NotApplicable;
                        objMaster.WebUrl = web.Url;
                        //Added in Output Object <objMaster> 
                    }
                    else 
                    {
                        Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb] We have not changed the master page because this new Master Page " + masterPageUrl + " is not present in Gallary, for Web " + WebUrl);
                        Console.WriteLine("[ChangeMasterPageForWeb] We have not changed the master page because this new Master Page " + masterPageUrl + " is not present in Gallary, for Web " + WebUrl);
                    }
                }
                else
                {
                    Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb] Please check if the site exists and the user has required access permissions on this site: " + WebUrl);
                    Console.WriteLine("[ChangeMasterPageForWeb] Please check if the site exists and the user has required access permissions on this site: " + WebUrl);
                }

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END] [ChangeMasterPageForWeb] EXIT FROM FUNCTION ChangeMasterPageForWeb for WebUrl: " + WebUrl);
                Console.WriteLine("[END] [ChangeMasterPageForWeb] EXIT FROM FUNCTION ChangeMasterPageForWeb for WebUrl: " + WebUrl);
                
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "MasterPage", ex.Message, ex.ToString(), "ChangeMasterPageForWeb", ex.GetType().ToString(), exceptionCommentsInfo1);
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[EXCEPTION][ChangeMasterPageForWeb] Exception Message: " + ex.Message + ", Exception Comment: " + exceptionCommentsInfo1);
                
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[EXCEPTION][ChangeMasterPageForWeb] Exception Message: " + ex.Message + " for Web:  " + WebUrl);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            //If ==> This is for WEB
            if (ActionType.ToString().ToLower() == Constants.ActionType_Web.ToLower())
            {
                if (objMaster != null)
                {
                    _WriteMasterList.Add(objMaster);
                }

                FileUtility.WriteCsVintoFile(outPutFolder +@"\" + Constants.MasterPageUsage, ref _WriteMasterList,
                        ref headerMasterPage);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[ChangeMasterPageForWeb] Writing the Replace Output CSV file after replacing the master page - FileUtility.WriteCsVintoFile");
                Console.WriteLine("[ChangeMasterPageForWeb] Writing the Replace Output CSV file after replacing the master page - FileUtility.WriteCsVintoFile");

                Logger.AddMessageToTraceLogFile(Constants.Logging, "[END][ChangeMasterPageForWeb] EXIT FROM FUNCTION ChangeMasterPageForWeb for WebUrl: " + WebUrl);
                Console.WriteLine("[END][ChangeMasterPageForWeb] EXIT FROM FUNCTION ChangeMasterPageForWeb for WebUrl: " + WebUrl);

                Logger.AddMessageToTraceLogFile(Constants.Logging, "############## Master Page Trasnformation Utility Execution Completed for Web ##############");
                Console.WriteLine("############## Master Page Trasnformation Utility Execution Completed  for Web ##############");
            }

            return objMaster;
        }

        public string GetMasterPageRelativeURL(ClientContext clientContext, string MasterPageURL)
        {
            string _masterPageUrl = string.Empty;

            //User has Input Master Page URL from Root Gallery
            if (MasterPageURL.ToLower().StartsWith("/_catalogs/masterpage/"))
            {
                Web rootWeb = clientContext.Site.RootWeb;
                clientContext.Load(rootWeb);
                clientContext.ExecuteQuery();

                _masterPageUrl = rootWeb.ServerRelativeUrl.ToString() + MasterPageURL;
            }
            else if (MasterPageURL.ToLower().Contains("/_catalogs/masterpage/"))
            {
                _masterPageUrl = MasterPageURL;
            }
            else
            {
                Web rootWeb = clientContext.Site.RootWeb;

                clientContext.Load(rootWeb);
                clientContext.ExecuteQuery();

                if (rootWeb.ServerRelativeUrl.ToString().EndsWith("/"))
                {
                    _masterPageUrl = rootWeb.ServerRelativeUrl.ToString() + "_catalogs/masterpage/" + MasterPageURL;
                }
                else
                {
                    _masterPageUrl = rootWeb.ServerRelativeUrl.ToString() + "/_catalogs/masterpage/" + MasterPageURL;
                }
            }

            return _masterPageUrl;

        }

        /// <summary>
        /// How to determine if a file exists in a SharePoint SPFolder
        /// CSOM: File Check in SP Gallary. It would actually throw an exception if the file doesn't exist
        /// </summary>
        public bool Check_MasterPageExistsINGallery(string WebUrl, string MasterPageURL)
        {
            using (var clientContext = new ClientContext(WebUrl))
            {
                Web web = clientContext.Web;
                Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(MasterPageURL);
                bool bExists = false;

                try
                {
                    clientContext.Load(file);
                    clientContext.ExecuteQuery(); //Raises exception if the file doesn't exist
                    bExists = file.Exists;  //may not be needed - here for good measure
                }
                catch { }

                return bExists;
            }
        }

       /// <summary>
       /// Master Pages and Page Layouts Always be Saved in Root Site inside the "_catalogs" folder
       /// Function Used For: How to determine if a file exists in a SharePoint SPFolder - True => Exists, False => Not Exists
       /// CSOM: File Check in SP Gallary. It would actually throw an exception if the file doesn't exist 
       /// </summary>
        public bool Check_MasterPageExistsINGallery(ClientContext clientContext, string MasterPageURL)
        {
            //Checking The File in Root Web Gallery
            Microsoft.SharePoint.Client.File file = clientContext.Site.RootWeb.GetFileByServerRelativeUrl(MasterPageURL);

            bool bExists = false;
            try
            {
                clientContext.Load(file);
                clientContext.ExecuteQuery(); //Raises exception if the file doesn't exist
                bExists = file.Exists; 
            }
            catch { }

            return bExists;
        }
        
        /// <summary>
        /// This function delete all the existing files from <outPutFolder> folder
        /// </summary>
        private void DeleteMasterPage_ReplaceOutPutFiles(string outPutFolder)
        {
            FileUtility.DeleteFiles(outPutFolder + @"\" + Constants.MasterPageUsage);
        }
        private string GetPageNameFromURL(string URL)
        {
            string FileName = string.Empty;

            if (URL != null)
            { FileName = System.IO.Path.GetFileName(URL); }

            return FileName;
        }
        private string GetPageNameWithSuffix(string PageNameWithExtension, string Suffix)
        {
            string PageNameWithSuffix = string.Empty;

            string Name = System.IO.Path.GetFileNameWithoutExtension(PageNameWithSuffix);
            string Extension = System.IO.Path.GetExtension(PageNameWithSuffix);

            return PageNameWithSuffix = Name + Suffix + Extension;
        }
        
    }
}
