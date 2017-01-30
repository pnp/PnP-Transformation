using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

using OfficeDevPnP.Core;

using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using JDP.Remediation.Console.Common.Utilities;

namespace JDP.Remediation.Console
{
    public class Helper
    {
        public static int contextCount = 0;
        public static bool alreadyAuthorized = false;

        public class MasterPageInfo
        {
            public string MasterPageUrl;
            public bool InheritMaster;
            public string CustomMasterPageUrl;
            public bool InheritCustomMaster;
        }

        public static ClientContext CreateAuthenticatedUserContextOld(string domain, string username, SecureString password, string siteUrl)
        {
            ClientContext userContext = new ClientContext(siteUrl);
            if (String.IsNullOrEmpty(domain))
            {
                // use o365 authentication (SPO-MT or vNext)
                userContext.Credentials = new SharePointOnlineCredentials(username, password);
            }
            else
            {
                // use Windows authentication (SPO-D or On-Prem) 
                userContext.Credentials = new NetworkCredential(username, password, domain);
            }

            return userContext;
        }
        public static ClientContext CreateAuthenticatedUserContext(string domain, string username, SecureString password, string siteUrl)
        {
            ClientContext userContext = new ClientContext(siteUrl);
            try
            {
                if (String.IsNullOrEmpty(domain))
                {
                    // use o365 authentication (SPO-MT or vNext)
                    userContext.Credentials = new SharePointOnlineCredentials(username, password);
                }
                else
                {
                    // use Windows authentication (SPO-D or On-Prem) 
                    userContext.Credentials = new NetworkCredential(username, password, domain);
                }

                Web web = userContext.Web;
                userContext.Load(web);
                userContext.ExecuteQuery();
                contextCount = 0;
                alreadyAuthorized = true;
                return userContext;
            }
            catch (System.Net.WebException exc)
            {
                if (exc.Message.ToLower().Contains("unauthorized") && alreadyAuthorized == false)
                {
                    contextCount++;
                    if (contextCount == 1)
                    {
                        Logger.LogMessage(String.Format("\n"), true);
                        Logger.LogErrorMessage(String.Format("Attempt [{0}]: You have entered an invalid username or password. The maximum retry attempts allowed for login are 3. You have 2 more attempts.", contextCount, 3 - contextCount), true);
                        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, siteUrl, "Authentication", exc.Message, exc.ToString(), "CreateAuthenticatedUserContext()", exc.GetType().ToString());
                    }
                    else if (contextCount == 2)
                    {
                        Logger.LogMessage(String.Format("\n"), true);
                        Logger.LogErrorMessage(String.Format("Attempt [{0}]: Incorrect login credentials twice. You have one more attempt. If you fail to enter correct credentials this time, application would be terminated.", contextCount, 3 - contextCount), true);
                        //Logger.LogErrorMessage(String.Format("\nWrong user credentials given for {0} time. {1} attemps remained", contextCount, 3 - contextCount), true);
                        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, siteUrl, "Authentication", exc.Message, exc.ToString(), "CreateAuthenticatedUserContext()", exc.GetType().ToString());
                    }
                    else if (contextCount == 3)
                    {
                        Logger.LogErrorMessage(String.Format("\n"), true);
                        Logger.LogErrorMessage(String.Format("Attempt [{0}]: You have entered an invalid username or password. Press any key to terminate the application!!", contextCount, 3 - contextCount), true);
                        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, siteUrl, "Authentication", exc.Message, exc.ToString(), "CreateAuthenticatedUserContext()", exc.GetType().ToString());

                        System.Console.ReadKey();
                        Environment.Exit(0);
                    }

                    Program.GetCredentials();
                    userContext = CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, siteUrl);
                }
            }
            catch (System.ArgumentNullException exc)
            {
                contextCount++;
                if (contextCount == 1)
                {
                    Logger.LogMessage(String.Format("\n"), true);
                    Logger.LogErrorMessage(String.Format("Attempt [{0}]: You have entered an invalid username or password. The maximum retry attempts allowed for login are 3. You have 2 more attempts.", contextCount, 3 - contextCount), true);
                    ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, siteUrl, "Authentication", exc.Message, exc.ToString(), "CreateAuthenticatedUserContext()", exc.GetType().ToString());
                }
                else if (contextCount == 2)
                {
                    Logger.LogMessage(String.Format("\n"), true);
                    Logger.LogErrorMessage(String.Format("Attempt [{0}]: Incorrect login credentials twice. You have one more attempt. If you fail to enter correct credentials this time, application would be terminated.", contextCount, 3 - contextCount), true);
                    //Logger.LogErrorMessage(String.Format("\nWrong user credentials given for {0} time. {1} attemps remained", contextCount, 3 - contextCount), true);
                    ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, siteUrl, "Authentication", exc.Message, exc.ToString(), "CreateAuthenticatedUserContext()", exc.GetType().ToString());
                }
                else if (contextCount == 3)
                {
                    Logger.LogErrorMessage(String.Format("\n"), true);
                    Logger.LogErrorMessage(String.Format("Attempt [{0}]: You have entered an invalid username or password. Press any key to terminate the application!!", contextCount, 3 - contextCount), true);
                    ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, siteUrl, "Authentication", exc.Message, exc.ToString(), "CreateAuthenticatedUserContext()", exc.GetType().ToString());

                    System.Console.ReadKey();
                    Environment.Exit(0);
                }

                Program.GetCredentials();
                userContext = CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, siteUrl);
            }

            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("\nCreateAuthenticatedUserContext() failed for {0}: Error={1}", siteUrl, ex.Message), false);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, siteUrl, "Authentication", ex.Message, ex.ToString(), "CreateAuthenticatedUserContext()", ex.GetType().ToString());
            }

            return userContext;
        }


        /// <summary>
        /// Creates a Secure String
        /// </summary>
        /// <param name="data">string to be converted</param>
        /// <returns>secure string instance</returns>
        public static SecureString CreateSecureString(string data)
        {
            if (data == null || string.IsNullOrEmpty(data))
            {
                return null;
            }

            System.Security.SecureString secureString = new System.Security.SecureString();

            char[] charArray = data.ToCharArray();

            foreach (char ch in charArray)
            {
                secureString.AppendChar(ch);
            }

            return secureString;
        }

        public static Field EnsureSiteColumn(Web root, Guid fieldID, string fieldAsXml)
        {
            Field existingField = root.GetFieldById<Field>(fieldID);
            if (existingField != null)
            {
                return existingField;
            }
            Field newField = root.CreateField(fieldAsXml, true);
            return newField;
        }

        public static File GetFileFromWeb(Web web, string filePath)
        {
            try
            {
                File webFile = web.GetFileByServerRelativeUrl(filePath);
                web.Context.Load(webFile);
                web.Context.Load(webFile.ListItemAllFields);
                web.Context.ExecuteQuery();

                return webFile;
            }
            catch {}

            return null;
        }

        public static string UploadMasterPage(Web web, string mpFileName, string localFilePath, string title, string description)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Uploading Master Page File: {0} ...", mpFileName), false);

                List mpGallery = web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                Folder mpGalleryRoot = mpGallery.RootFolder;
                web.Context.Load(mpGallery);
                web.Context.Load(mpGalleryRoot);
                web.Context.ExecuteQuery();

                string mpFilePath = mpGalleryRoot.ServerRelativeUrl + "/" + mpFileName;
                File mpFile = GetFileFromWeb(web, mpFilePath);
                if (mpFile == null)
                {
                    // Get the file name from the provided path
                    Byte[] fileBytes = System.IO.File.ReadAllBytes(localFilePath);

                    // Use CSOM to upload the file in
                    FileCreationInformation newFile = new FileCreationInformation();
                    newFile.Content = fileBytes;
                    newFile.Overwrite = true;
                    newFile.Url = mpFileName;

                    File uploadFile = mpGalleryRoot.Files.Add(newFile);
                    web.Context.Load(uploadFile);
                    web.Context.ExecuteQuery();
                }

                // Grab the file we just uploaded so we can edit its properties
                mpFile = GetFileFromWeb(web, mpFilePath);
                if (mpGallery.ForceCheckout || mpGallery.EnableVersioning)
                {
                    if (mpFile.CheckOutType == CheckOutType.None)
                    {
                        mpFile.CheckOut();
                    }
                }

                ListItem fileListItem = mpFile.ListItemAllFields;
                fileListItem["MasterPageDescription"] = description;
                fileListItem["UIVersion"] = "15";
                if (mpGallery.AllowContentTypes && mpGallery.ContentTypesEnabled)
                {
                    fileListItem["Title"] = title;
                    fileListItem["ContentTypeId"] = Constants.MASTERPAGE_CONTENT_TYPE;
                }
                fileListItem.Update();

                if (mpGallery.ForceCheckout || mpGallery.EnableVersioning)
                {
                    mpFile.CheckIn(string.Empty, CheckinType.MajorCheckIn);
                    if (mpGallery.EnableModeration)
                    {
                        mpFile.Approve("");
                    }
                }
                web.Context.ExecuteQuery();

                Logger.LogSuccessMessage(String.Format("Uploaded Master Page File: {0}", mpFile.ServerRelativeUrl), false);
                return mpFile.ServerRelativeUrl;
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("UploadMasterPage() failed for {0}: Error={1}", web.Url, ex.Message), false);
                return String.Empty;
            }
        }

        public static string PublishMasterPage(Web web, string mpFilePath, string title, string description)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Publishing Master Page File: {0} ...", mpFilePath), false);

                List mpGallery = web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                Folder mpGalleryRoot = mpGallery.RootFolder;
                web.Context.Load(mpGallery);
                web.Context.Load(mpGalleryRoot);
                web.Context.ExecuteQuery();

                File mpFile = GetFileFromWeb(web, mpFilePath);
                if (mpFile == null)
                {
                    Logger.LogErrorMessage(String.Format("UploadMasterPage() failed for {0}: Error=File Not Found", web.Url), false);
                    return String.Empty;
                }

                if (mpGallery.ForceCheckout || mpGallery.EnableVersioning)
                {
                    if (mpFile.CheckOutType == CheckOutType.None)
                    {
                        mpFile.CheckOut();
                    }
                }

                ListItem fileListItem = mpFile.ListItemAllFields;
                fileListItem["MasterPageDescription"] = description;
                fileListItem["UIVersion"] = "15";
                if (mpGallery.AllowContentTypes && mpGallery.ContentTypesEnabled)
                {
                    fileListItem["Title"] = title;
                    fileListItem["ContentTypeId"] = Constants.MASTERPAGE_CONTENT_TYPE;
                }
                fileListItem.Update();

                if (mpGallery.ForceCheckout || mpGallery.EnableVersioning)
                {
                    mpFile.CheckIn(string.Empty, CheckinType.MajorCheckIn);
                    if (mpGallery.EnableModeration)
                    {
                        mpFile.Approve("");
                    }
                }
                web.Context.ExecuteQuery();

                Logger.LogSuccessMessage(String.Format("Published Master Page File: {0}", mpFile.ServerRelativeUrl), false);
                return mpFile.ServerRelativeUrl;
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("UploadMasterPage() failed for {0}: Error={1}", web.Url, ex.Message), false);
                return String.Empty;
            }
        }

        public static MasterPageInfo GetMasterPageInfo(Web web)
        {
            MasterPageInfo mpi = new MasterPageInfo();
            try
            {
                Logger.LogInfoMessage(String.Format("Getting Master Page info for: {0} ...", web.Url), false);

                web.Context.Load(web.AllProperties);
                web.Context.ExecuteQuery();

                mpi.MasterPageUrl = web.MasterUrl;
                mpi.CustomMasterPageUrl = web.CustomMasterUrl;
                mpi.InheritMaster = web.GetPropertyBagValueString(Constants.PropertyBagInheritMaster, "False").ToBoolean();
                mpi.InheritCustomMaster = web.GetPropertyBagValueString(Constants.PropertyBagInheritCustomMaster, "False").ToBoolean();
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("GetMasterPageInfo() failed for {0}: Error={1}", web.Url, ex.Message), false);
            }

            return mpi;
        }

        public static void SetMasterPages(Web web, string mpFilePath, bool isRoot, bool inheritMaster)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Setting Master Pages to: {0} ...", mpFilePath), false);

                Logger.LogInfoMessage(String.Format("MasterUrl (before): {0}", web.MasterUrl), false);
                Logger.LogInfoMessage(String.Format("CustomMasterUrl (before): {0}", web.CustomMasterUrl), false);

                web.Context.Load(web.AllProperties);
                web.Context.ExecuteQuery();

                web.MasterUrl = mpFilePath;
                web.CustomMasterUrl = mpFilePath;
                web.SetPropertyBagValue(Constants.PropertyBagInheritMaster, ((!isRoot && inheritMaster) ? "True" : "False"));
                web.SetPropertyBagValue(Constants.PropertyBagInheritCustomMaster, ((!isRoot && inheritMaster) ? "True" : "False"));
                web.Update();
                web.Context.ExecuteQuery();

                Logger.LogSuccessMessage(String.Format("MasterUrl (after): {0}", web.MasterUrl), false);
                Logger.LogSuccessMessage(String.Format("CustomMasterUrl (after): {0}", web.CustomMasterUrl), false);
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("SetMasterPages() failed for {0}: Error={1}", web.Url, ex.Message), false);
            }
        }

        public static void SetMasterPages(Web web, MasterPageInfo mpi, bool isRoot)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Setting Master Pages to: {0} & {1}...", mpi.MasterPageUrl, mpi.CustomMasterPageUrl), false);

                Logger.LogInfoMessage(String.Format("MasterUrl (before): {0}", web.MasterUrl), false);
                Logger.LogInfoMessage(String.Format("CustomMasterUrl (before): {0}", web.CustomMasterUrl), false);

                web.Context.Load(web.AllProperties);
                web.Context.ExecuteQuery();

                web.MasterUrl = mpi.MasterPageUrl;
                web.CustomMasterUrl = mpi.CustomMasterPageUrl;
                web.SetPropertyBagValue(Constants.PropertyBagInheritMaster, ((!isRoot && mpi.InheritMaster) ? "True" : "False"));
                web.SetPropertyBagValue(Constants.PropertyBagInheritCustomMaster, ((!isRoot && mpi.InheritCustomMaster) ? "True" : "False"));
                web.Update();
                web.Context.ExecuteQuery();

                Logger.LogSuccessMessage(String.Format("MasterUrl (after): {0}", web.MasterUrl), false);
                Logger.LogSuccessMessage(String.Format("CustomMasterUrl (after): {0}", web.CustomMasterUrl), false);
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("SetMasterPages() failed for {0}: Error={1}", web.Url, ex.Message), false);
            }
        }

        public static void DeleteMasterPageCatalogFile(Web web, string mpFileName)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Deleting Master Page File: {0} ...", mpFileName), false);

                List mpGallery = web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                Folder mpGalleryRoot = mpGallery.RootFolder;
                web.Context.Load(mpGalleryRoot);
                web.Context.ExecuteQuery();

                string mpFilePath = mpGalleryRoot.ServerRelativeUrl + "/" + mpFileName;
                File mpFile = web.GetFileByServerRelativeUrl(mpFilePath);

                Logger.LogInfoMessage(String.Format("Deleting Master Page File: {0} ...", mpFilePath), false);

                web.Context.Load(mpFile);
                web.Context.ExecuteQuery();

                if (mpFile.ServerObjectIsNull == false)
                {
                    mpFile.DeleteObject();
                    web.Context.ExecuteQuery();

                    Logger.LogSuccessMessage(String.Format("Deleted Master Page File: {0}", mpFilePath), false);
                }
                else
                {
                    Logger.LogErrorMessage(String.Format("DeleteMasterPageCatalogFile() failed for {0}: Error=File {1} was not found", web.Url, mpFileName), false);
                }
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName.Equals("System.IO.FileNotFoundException", StringComparison.InvariantCultureIgnoreCase) == false)
                {
                    Logger.LogErrorMessage(String.Format("DeleteMasterPageCatalogFile() failed for {0}: Error={1}", web.Url, ex.Message), false);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("DeleteMasterPageCatalogFile() failed for {0}: Error={1}", web.Url, ex.Message), false);
            }
        }
        /// <summary>
        /// Deletes the specified file from the specified web.
        /// </summary>
        /// <param name="web">this MUST be the web that contains the file to delete</param>
        /// <param name="serverRelativeFilePath">the SERVER-relative path to the file ("/sites/site/web/lib/folder/file.ext")</param>
        public static bool DeleteFileByServerRelativeUrl(Web web, string serverRelativeFilePath)
        {
            bool result = false;
            try
            {
                Logger.LogInfoMessage(String.Format("Deleting File: {0} ...", serverRelativeFilePath), false);

                File targetFile = web.GetFileByServerRelativeUrl(serverRelativeFilePath);
                web.Context.Load(targetFile);
                web.Context.ExecuteQuery();

                if (targetFile.ServerObjectIsNull == false)
                {
                    targetFile.DeleteObject();
                    web.Context.ExecuteQuery();
                    result = true;
                    Logger.LogSuccessMessage(String.Format("Deleted File: {0}", serverRelativeFilePath), true);
                }
                else
                {
                    Logger.LogErrorMessage(String.Format("DeleteFileByServerRelativeUrl() failed for {0}: Error={1}", serverRelativeFilePath, "File was not Found."), true);
                }
            }
            catch (ServerException ex)
            {
                Logger.LogErrorMessage(String.Format("[Helper: DeleteFileByServerRelativeUrl] failed for {0}: Error={1}", serverRelativeFilePath, ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "N/A", ex.Message, ex.ToString(), "DeleteFileByServerRelativeUrl",
                    ex.GetType().ToString(), String.Format("DeleteFileByServerRelativeUrl() failed for {0}: Error={1}", serverRelativeFilePath, ex.Message));
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[Helper: DeleteFileByServerRelativeUrl] failed for {0}: Error={1}", serverRelativeFilePath, ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, web.Url, "N/A", ex.Message, ex.ToString(), "DeleteFileByServerRelativeUrl",
                    ex.GetType().ToString(), String.Format("DeleteFileByServerRelativeUrl() failed for {0}: Error={1}", serverRelativeFilePath, ex.Message));
            }
            return result;
        }

        public static void AddWebPartToPage(Web web, string webPartFile, Microsoft.SharePoint.Client.File page, string zoneId)
        {
            Folder rootFolder = web.RootFolder;
            web.Context.Load(rootFolder, root => root.WelcomePage, root => root.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();

            LimitedWebPartManager wpMgr = page.GetLimitedWebPartManager(PersonalizationScope.Shared);
            XDocument xDoc = XDocument.Load(webPartFile);
            WebPartDefinition def = wpMgr.ImportWebPart(xDoc.ToString());
            WebPartDefinition wpDef = wpMgr.AddWebPart(def.WebPart, zoneId, 0);

            web.Context.Load(wpDef, wp => wp.Id);
            web.Context.ExecuteQuery();
        }

        public static void RemoveWebPartFromPage(Web web, string wpTitle, File wpPage)
        {
            LimitedWebPartManager wpManager = wpPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            web.Context.Load(wpManager.WebParts, wps => wps.Include(wp => wp.WebPart.Title));
            web.Context.ExecuteQueryRetry();

            if (wpManager.WebParts.Count >= 0)
            {
                for (int i = wpManager.WebParts.Count-1; i >= 0; i--)
                {
                    WebPart oWebPart = wpManager.WebParts[i].WebPart;
                    if (oWebPart.Title.Equals(wpTitle, StringComparison.InvariantCultureIgnoreCase))
                    {
                        wpManager.WebParts[i].DeleteWebPart();
                        web.Context.ExecuteQuery();
                    }
                }
            }
        }

        public static bool WebPartIsPresentOnPage(Web web, string wpTitle, File wpPage)
        {
            LimitedWebPartManager wpManager = wpPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            web.Context.Load(wpManager.WebParts, wps => wps.Include(wp => wp.WebPart.Title));
            web.Context.ExecuteQueryRetry();

            if (wpManager.WebParts.Count >= 0)
            {
                for (int i = 0; i < wpManager.WebParts.Count; i++)
                {
                    WebPart oWebPart = wpManager.WebParts[i].WebPart;
                    if (oWebPart.Title.Equals(wpTitle, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public static void UploadWebPartFile(Web web, string wpFileName, string localFilePath, string title, string description)
        {
            try 
            {
                Logger.LogInfoMessage(String.Format("Uploading Web Part File: {0} ...", wpFileName), false);

                List wpGallery = web.GetCatalog((int)ListTemplateType.WebPartCatalog);
                Folder wpGalleryRoot = wpGallery.RootFolder;
                web.Context.Load(wpGallery);
                web.Context.Load(wpGalleryRoot);
                web.Context.ExecuteQuery();

                string wpFilePath = wpGalleryRoot.ServerRelativeUrl + "/" + wpFileName;
                File wpFile = GetFileFromWeb(web, wpFilePath);
                if (wpFile == null)
                {
                    // Get the file name from the provided path
                    Byte[] fileBytes = System.IO.File.ReadAllBytes(localFilePath);

                    // Use CSOM to upload the file in
                    FileCreationInformation newFile = new FileCreationInformation();
                    newFile.Content = fileBytes;
                    newFile.Overwrite = true;
                    newFile.Url = wpFileName;

                    File uploadFile = wpGalleryRoot.Files.Add(newFile);
                    web.Context.Load(uploadFile);
                    web.Context.ExecuteQuery();
                }

                // Grab the file we just uploaded so we can edit its properties
                wpFile = GetFileFromWeb(web, wpFilePath);
                if (wpGallery.ForceCheckout || wpGallery.EnableVersioning)
                {
                    if (wpFile.CheckOutType == CheckOutType.None)
                    {
                        wpFile.CheckOut();
                    }
                }

                ListItem fileListItem = wpFile.ListItemAllFields;
                fileListItem["Title"] = title;
                fileListItem["WebPartDescription"] = description;
                fileListItem["Group"] = "Contoso Custom Web Parts";
                fileListItem.Update();

                if (wpGallery.ForceCheckout || wpGallery.EnableVersioning)
                {
                    wpFile.CheckIn(string.Empty, CheckinType.MajorCheckIn);
                    if (wpGallery.EnableModeration)
                    {
                        wpFile.Approve("");
                    }
                }
                web.Context.ExecuteQuery();

                Logger.LogSuccessMessage(String.Format("Uploaded Web Part File: {0}", wpFile.ServerRelativeUrl), false);
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("UploadWebPartFile() failed for {0}: Error={1}", web.Url, ex.Message), false);
            }
        }
        public static void DeleteWebPartFile(Web web, string wpFileName)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Deleting Web Part File: {0} ...", wpFileName), false);

                List wpGallery = web.GetCatalog((int)ListTemplateType.WebPartCatalog);
                Folder wpGalleryRoot = wpGallery.RootFolder;
                web.Context.Load(wpGalleryRoot);
                web.Context.ExecuteQuery();

                string wpFilePath = wpGalleryRoot.ServerRelativeUrl + "/" + wpFileName;
                File wpFile = web.GetFileByServerRelativeUrl(wpFilePath);

                Logger.LogInfoMessage(String.Format("Deleting Web Part File: {0} ...", wpFilePath), false);

                web.Context.Load(wpFile);
                web.Context.ExecuteQuery();

                if (wpFile.ServerObjectIsNull == false)
                {
                    wpFile.DeleteObject();
                    web.Context.ExecuteQuery();

                    Logger.LogSuccessMessage(String.Format("Deleted Web Part File: {0}", wpFilePath), false);
                }
                else
                {
                    Logger.LogErrorMessage(String.Format("DeleteWebPartFile() failed for {0}: Error=File {1} was not found", web.Url, wpFileName), false);
                }
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName.Equals("System.IO.FileNotFoundException", StringComparison.InvariantCultureIgnoreCase) == false)
                {
                    Logger.LogErrorMessage(String.Format("DeleteWebPartFile() failed for {0}: Error={1}", web.Url, ex.Message), false);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("DeleteWebPartFile() failed for {0}: Error={1}", web.Url, ex.Message), false);
            }
        }

        public static void UploadDisplayTemplateFileJS(Web web, string dtFileName, string localFilePath, string displayTemplateFolderName, string title, string description)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Uploading Display Template JS File: {0} ...", dtFileName), false);

                List mpGallery = web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                Folder mpGalleryRoot = mpGallery.RootFolder;
                web.Context.Load(mpGallery);
                web.Context.Load(mpGalleryRoot);
                web.Context.ExecuteQuery();

                string serverRelativeFolderUrl = mpGalleryRoot.ServerRelativeUrl + "/Display Templates/" + displayTemplateFolderName;
                Folder dtFolder = web.GetFolderByServerRelativeUrl(serverRelativeFolderUrl);
                web.Context.Load(dtFolder);
                web.Context.ExecuteQuery();

                string dtFilePath = dtFolder.ServerRelativeUrl + "/" + dtFileName;
                File dtFile = GetFileFromWeb(web, dtFilePath);
                if (dtFile == null)
                {
                    // Get the file name from the provided path
                    Byte[] fileBytes = System.IO.File.ReadAllBytes(localFilePath);

                    // Use CSOM to upload the file in
                    FileCreationInformation newFile = new FileCreationInformation();
                    newFile.Content = fileBytes;
                    newFile.Overwrite = true;
                    newFile.Url = dtFilePath;

                    File uploadFile = mpGalleryRoot.Files.Add(newFile);
                    web.Context.Load(uploadFile);
                    web.Context.ExecuteQuery();
                }

                // Grab the file we just uploaded so we can edit its properties
                dtFile = GetFileFromWeb(web, dtFilePath);
                if (mpGallery.ForceCheckout || mpGallery.EnableVersioning)
                {
                    if (dtFile.CheckOutType == CheckOutType.None)
                    {
                        dtFile.CheckOut();
                    }
                }

                ListItem fileListItem = dtFile.ListItemAllFields;
                string controlContentTypeId = GetContentType(web, mpGallery, "Display Template Code");
                fileListItem["Title"] = title;
                fileListItem["MasterPageDescription"] = description;
                fileListItem["ContentTypeId"] = controlContentTypeId;
                fileListItem["TargetControlType"] = ";#Content Web Parts;#";
                fileListItem["DisplayTemplateLevel"] = "Item";
                fileListItem["TemplateHidden"] = "0";
                fileListItem["UIVersion"] = "15";
                fileListItem["ManagedPropertyMapping"] = "'Link URL'{Link URL}:'Path','Line 1'{Line 1}:'Title','Line 2'{Line 2}:'','FileExtension','SecondaryFileExtension'";
                fileListItem.Update();

                if (mpGallery.ForceCheckout || mpGallery.EnableVersioning)
                {
                    dtFile.CheckIn(string.Empty, CheckinType.MajorCheckIn);
                    if (mpGallery.EnableModeration)
                    {
                        dtFile.Approve("");
                    }
                }
                web.Context.ExecuteQuery();

                Logger.LogSuccessMessage(String.Format("Uploaded Display Template JS File: {0}", dtFile.ServerRelativeUrl), false);
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("UploadDisplayTemplateFileJS() failed for {0}: Error={1}", web.Url, ex.Message), false);
            }
        }

        private static string GetContentType(Web web, List list, string contentType)
        {
            ContentTypeCollection collection = list.ContentTypes;
            web.Context.Load(collection);
            web.Context.ExecuteQuery();
            var ct = collection.Where(c => c.Name == contentType).FirstOrDefault();
            string contentTypeID = "";
            if (ct != null)
            {
                contentTypeID = ct.StringId;
            }

            return contentTypeID;
        }

        public static void DeleteListByUrl(Web web, string webRelativeUrl)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Deleting List {0} from {1} ...", webRelativeUrl, web.Url), false);

                List targetList = web.GetListByUrl(webRelativeUrl);

                if (targetList != null && targetList.ServerObjectIsNull == false)
                {
                    targetList.DeleteObject();
                    web.Context.ExecuteQuery();

                    Logger.LogSuccessMessage(String.Format("Deleted List {0} from {1}", webRelativeUrl, web.Url), false);
                }
                else
                {
                    Logger.LogErrorMessage(String.Format("DeleteListByUrl() failed for {0}: Error=List {1} was not found", web.Url, webRelativeUrl), false);
                }
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName.Equals("System.IO.FileNotFoundException", StringComparison.InvariantCultureIgnoreCase) == false)
                {
                    Logger.LogErrorMessage(String.Format("DeleteListByUrl() failed for {0}: Error={1}", web.Url, ex.Message), false);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("DeleteListByUrl() failed for {0}: Error={1}", web.Url, ex.Message), false);
            }
        }

        public static string[] ReadInputFile(string inputFileSpec, bool hasHeader)
        {
            try
            {
                if (hasHeader == true)
                {
                    // remove the header row from the resulting string array...
                    List<string> temp = new List<string>(System.IO.File.ReadAllLines(inputFileSpec));
                    temp.RemoveAt(0);
                    return temp.ToArray();
                }
                else
                {
                    return System.IO.File.ReadAllLines(inputFileSpec);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ReadInputFile() failed for {0}: Error={1}", inputFileSpec, ex.Message), true);
                return new string[0]; 
            }
        }

        // Parses the CSV input line and returns a string array where each entry corresponds to a field parsed from the line
        // Strips double quotes from those fields that contain commas: e.g., 123,"abc,def",456
        public static string[] ParseInputLine(string inputLine)
        {
            try
            {
                if (inputLine.Contains('"'))
                {
                    int pos1 = inputLine.IndexOf('"');
                    int pos2 = inputLine.IndexOf('"', pos1 + 1);

                    string left = inputLine.Substring(0, pos1);
                    left = left.Trim(new char[] { '"', ',' });

                    string center = inputLine.Substring(pos1, pos2 - pos1);
                    center = center.TrimStart(new char[] { '"', ',' });

                    string right = inputLine.Substring(pos2);
                    right = right.TrimStart(new char[] { '"', ',' });

                    List<string> result = new List<string>();
                    string[] temp = null;

                    temp = ParseInputLine(left);
                    if (temp.Length > 0) result.AddRange(temp);

                    result.Add(center);

                    temp = ParseInputLine(right);
                    if (temp.Length > 0) result.AddRange(temp);

                    return result.ToArray();
                }
                else
                {
                    inputLine = inputLine.Trim(new char[] { ',' });
                    if (String.IsNullOrEmpty(inputLine))
                    {
                        return new string[0];
                    }
                    return inputLine.Split(',');
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("ParseInputLine() failed for [{0}]: Error={1}", inputLine, ex.Message), false);
                return new string[0];
            }
        }

        public static string SafeGetFileAsString(Web web, string serverRelativeMappingFilePath)
        {
            try
            {
                return web.GetFileAsString(serverRelativeMappingFilePath);
            }
            catch
            {
                return String.Empty;
            }
        }

        public static string UploadDeviceChannelMappingFile(Web web, string serverRelativeMappingFilePath, string fileContents, string description)
        {
            try
            {
                Logger.LogInfoMessage(String.Format("Uploading Device Channel Mapping File: {0} ...", serverRelativeMappingFilePath), false);

                // grab a reference to the MP Gallery where the device channel mapping file resides.
                List mpGallery = web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                Folder mpGalleryRoot = mpGallery.RootFolder;
                web.Context.Load(mpGallery);
                web.Context.Load(mpGalleryRoot);
                web.Context.ExecuteQuery();

                // get the file and check-out if necessary
                File dcmFile = GetFileFromWeb(web, serverRelativeMappingFilePath);
                if (dcmFile != null)
                {
                    web.CheckOutFile(serverRelativeMappingFilePath);
                }

                // prepare the file contents for upload
                Byte[] fileBytes = System.Text.Encoding.UTF8.GetBytes(fileContents);

                // Use CSOM to upload the file
                FileCreationInformation newFile = new FileCreationInformation();
                newFile.Content = fileBytes;
                newFile.Overwrite = true;
                newFile.Url = serverRelativeMappingFilePath;

                File uploadFile = mpGalleryRoot.Files.Add(newFile);
                web.Context.Load(uploadFile);
                web.Context.ExecuteQuery();

                // check-in and approve as necessary
                if (mpGallery.ForceCheckout || mpGallery.EnableVersioning)
                {
                    web.CheckInFile(uploadFile.ServerRelativeUrl, CheckinType.MajorCheckIn, description);
                    if (mpGallery.EnableModeration)
                    {
                        web.ApproveFile(uploadFile.ServerRelativeUrl, description);
                    }
                }

                Logger.LogSuccessMessage(String.Format("Uploaded Device Channel Mapping File: {0}", uploadFile.ServerRelativeUrl), false);
                return uploadFile.ServerRelativeUrl;
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("UploadDeviceChannelMappingFile() failed for {0}: Error={1}", serverRelativeMappingFilePath, ex.Message), false);
                return String.Empty;
            }
        }

    }
}
