using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

using EmployeeRegistration.MVCWeb.Models;
using System.IO;

namespace EmployeeRegistration.MVCWeb.Controllers
{
    public class EmployeeController : Controller
    {
        //
        // GET: /Employee/
        public ActionResult Index()
        {
            return View();
        }

        private string getStateIDFromCityList()
        {
            return "";
        }

        private string ConvertToEmptyString(string input)
        {
            if (input != null)
            {
                return input;
            }

            return string.Empty;
        }

        private string ConvertObjectToString(object input)
        {
            if (input != null)
            {
                return input.ToString();
            }

            return string.Empty;
        }

        private void LoadUserSiteGroups(ref Employee emp, ClientContext clientContext)
        {
            try
            {
                User user = clientContext.Web.EnsureUser(emp.UserID);
                GroupCollection userGroups = user.Groups;
                GroupCollection siteGroups = clientContext.Web.SiteGroups;
                clientContext.Load(siteGroups, groups => groups.Include(group => group.Title, group => group.Id));
                clientContext.Load(user);
                clientContext.Load(userGroups);
                clientContext.ExecuteQuery();

                List<int> memberOf = new List<int>();
                foreach (Group group in userGroups)
                {
                    memberOf.Add(group.Id);
                }

                List<SiteGroup> siteGropuList = new List<SiteGroup>();
                foreach (var siteGroup in siteGroups)
                {
                    siteGropuList.Add(new SiteGroup
                    {
                        Name = siteGroup.Title,
                        Id = siteGroup.Id,
                        Checked = memberOf.Contains(siteGroup.Id)
                    });
                }
                emp.SiteGroups = siteGropuList;
                emp.SiteGroupsCount = siteGropuList.Count;
                emp.PreviouslySelectedSiteGroups = memberOf.ToArray();
                emp.PreviouslySelectedSiteGroupsCount = memberOf.Count;
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.Message + "\n" + ex.StackTrace;
            }
        }

        [SharePointContextFilter]
        public ActionResult EmployeeForm()
        {
            Employee emp = new Employee();

            try
            {
                var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

                using (var clientContext = spContext.CreateUserClientContextForSPHost())
                {
                    if (clientContext != null)
                    {

                        SetupManager.Provision(clientContext);

                        var web = clientContext.Web;

                        List desgList = web.Lists.GetByTitle("EmpDesignation");
                        ListItemCollection desgItems = desgList.GetItems(CamlQuery.CreateAllItemsQuery());

                        List countryList = web.Lists.GetByTitle("EmpCountry");
                        ListItemCollection countryItems = countryList.GetItems(CamlQuery.CreateAllItemsQuery());

                        string itemID = Request.QueryString["itemId"];
                        ListItem emplistItem = null;

                        if (itemID != null)
                        {
                            List lstEmployee = web.Lists.GetByTitle("Employees");
                            emplistItem = lstEmployee.GetItemById(itemID);
                            clientContext.Load(emplistItem);
                            emp.Id = itemID;
                        }

                        clientContext.Load(desgItems);
                        clientContext.Load(countryItems);
                        clientContext.ExecuteQuery();

                        List<SelectListItem> empDesgList = new List<SelectListItem>();
                        foreach (var item in desgItems)
                        {
                            empDesgList.Add(new SelectListItem { Text = item["Title"].ToString() });
                        }
                        emp.Designations = new SelectList(empDesgList, "Text", "Text");

                        List<SelectListItem> cList = new List<SelectListItem>();
                        foreach (var item in countryItems)
                        {
                            cList.Add(new SelectListItem { Text = item["Title"].ToString(), Value = item["ID"].ToString() });
                        }
                        emp.Countries = new SelectList(cList, "Value", "Text");

                        string empDesignation = string.Empty;
                        int stateID = 0;
                        int countryId = 0;

                        if (emplistItem != null)
                        {
                            emp.EmpNumber = ConvertObjectToString(emplistItem["EmpNumber"]);
                            emp.Name = ConvertObjectToString(emplistItem["Title"]);
                            emp.UserID = ConvertObjectToString(emplistItem["UserID"]);
                            emp.EmpManager = ConvertObjectToString(emplistItem["EmpManager"]);
                            emp.Designation = ConvertObjectToString(emplistItem["Designation"]);

                            string cityVal = ConvertObjectToString(emplistItem["Location"]);
                            ViewBag.JsCity = "";
                            ViewBag.JsStateID = "";

                            if (cityVal != "")
                            {
                                ViewBag.JsCity = cityVal;
                                List lstCity = web.Lists.GetByTitle("EmpCity");
                                CamlQuery query = new CamlQuery();
                                query.ViewXml = string.Format("<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>{0}</Value></Eq></Where></Query></View>", cityVal);
                                ListItemCollection cityItems = lstCity.GetItems(query);
                                clientContext.Load(cityItems);
                                clientContext.ExecuteQuery();
                                if (cityItems.Count > 0)
                                {
                                    stateID = (cityItems[0]["State"] as FieldLookupValue).LookupId;
                                }
                                ViewBag.JsStateID = stateID;

                                List lstSate = web.Lists.GetByTitle("EmpState");
                                query.ViewXml = string.Format("<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Number'>{0}</Value></Eq></Where></Query></View>", stateID);
                                ListItemCollection stateItems = lstSate.GetItems(query);
                                clientContext.Load(stateItems);
                                clientContext.ExecuteQuery();
                                if (stateItems.Count > 0)
                                {
                                    countryId = (stateItems[0]["Country"] as FieldLookupValue).LookupId;
                                }
                                emp.CountryID = countryId.ToString();
                            }

                            string skillsData = ConvertObjectToString(emplistItem["Skills"]);
                            string[] skills = skillsData.Split(';');
                            List<Skill> lsSkills = new List<Skill>();
                            foreach (string skillData in skills)
                            {
                                if (skillData != "")
                                {
                                    string[] skill = skillData.Split(',');
                                    lsSkills.Add(new Skill { Technology = skill[0], Experience = skill[1] });

                                }
                            }
                            emp.Skills = lsSkills;
                            emp.SkillsCount = lsSkills.Count;

                            string attachementID = ConvertObjectToString(emplistItem["AttachmentID"]);
                            if (attachementID != "")
                            {
                                List lstAttachments = web.Lists.GetByTitle("EmpAttachments");
                                CamlQuery queryAttachments = new CamlQuery();
                                queryAttachments.ViewXml = string.Format("<View><Query><Where><Eq><FieldRef Name='AttachmentID' /><Value Type='Text'>{0}</Value></Eq></Where></Query></View>", attachementID);

                                ListItemCollection attachmentItems = lstAttachments.GetItems(queryAttachments);
                                clientContext.Load(attachmentItems);
                                clientContext.ExecuteQuery();

                                List<EmpAttachment> lsAttachments = new List<EmpAttachment>();
                                if (attachmentItems.Count > 0)
                                {
                                    foreach (ListItem item in attachmentItems)
                                    {
                                        lsAttachments.Add(new EmpAttachment
                                        {
                                            FileName = item["Title"].ToString(),
                                            FileUrl = Request.QueryString["SPHostUrl"] + "/Lists/EmpAttachments/" + item["FileLeafRef"].ToString(),
                                            FileRelativeUrl = item["FileRef"].ToString()
                                        });
                                    }
                                }
                                emp.AttachmentID = attachementID;
                                emp.Attachments = lsAttachments;
                                emp.AttachmentsCount = lsAttachments.Count;
                            }
                            else
                            {
                                emp.AttachmentID = Guid.NewGuid().ToString();
                            }

                            emp.ActionName = "UpdateEmployeeToSPList";
                            emp.SubmitButtonName = "Update Employee";
                        }
                        else
                        {
                            PeopleManager peopleManager = new PeopleManager(clientContext);
                            PersonProperties personProperties = peopleManager.GetMyProperties();
                            clientContext.Load(personProperties, p => p.AccountName);
                            clientContext.ExecuteQuery();

                            if (personProperties != null && personProperties.AccountName != null)
                            {
                                emp.UserID = personProperties.AccountName;
                            }

                            List<Skill> lsSkills = new List<Skill>();
                            lsSkills.Add(new Skill { Technology = "", Experience = "" });
                            emp.Skills = lsSkills;
                            emp.SkillsCount = lsSkills.Count;

                            emp.AttachmentID = Guid.NewGuid().ToString();
                            emp.isFileUploaded = false;

                            emp.ActionName = "AddEmployeeToSPList";
                            emp.SubmitButtonName = "Add Employee";
                        }

                        LoadUserSiteGroups(ref emp, clientContext);
                    } //  if (clientContext != null)
                } // using (var clientContext

                ViewBag.SPURL = Request.QueryString["SPHostUrl"];
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.Message + "\n" + ex.StackTrace;
            }

            return View(emp);
        }

        private void AddOrRemoveUserToGroup(IEnumerable<int> groupIDs, UserOpType operationType, ClientContext clientContext, string userID)
        {
            if (groupIDs.Count() > 0)
            {
                GroupCollection collGroup = clientContext.Web.SiteGroups;

                UserCreationInformation userCreationInfo = new UserCreationInformation();
                userCreationInfo.LoginName = userID;

                foreach (int groupID in groupIDs)
                {
                    Group oGroup = collGroup.GetById(groupID);
                    if (operationType.Equals(UserOpType.AddUser))
                    {
                        oGroup.Users.Add(userCreationInfo);
                        clientContext.ExecuteQuery();
                    }
                    else
                    {
                        User user = oGroup.Users.GetByLoginName(userID);
                        oGroup.Users.Remove(user);
                        clientContext.ExecuteQuery();
                    }
                } //  foreach (string group
            } // if (groups.Count() > 0)            
        }

        private void AddUserToSelectedSiteGroups(Employee model, ClientContext clientContext, string userID)
        {
            List<int> selectedGroups = new List<int>();
            int[] newGroups = new int[] { };
            int[] prevGroups = model.PreviouslySelectedSiteGroups;

            foreach (SiteGroup grp in model.SiteGroups)
            {
                if (grp.Checked)
                {
                    selectedGroups.Add(grp.Id);
                }
            }

            if (selectedGroups.Count > 0)
            {
                if (prevGroups != null && prevGroups.Length > 0)
                {
                    newGroups = selectedGroups.ToArray();

                    IEnumerable<int> deleteUserFromGroups = prevGroups.Except(newGroups);
                    IEnumerable<int> addUserToGroups = newGroups.Except(prevGroups);

                    AddOrRemoveUserToGroup(deleteUserFromGroups, UserOpType.RemoveUser, clientContext, userID);
                    AddOrRemoveUserToGroup(addUserToGroups, UserOpType.AddUser, clientContext, userID);
                }
                else // Add user to site groups if there is no previously selected groups
                {
                    AddOrRemoveUserToGroup(selectedGroups, UserOpType.AddUser, clientContext, userID);
                }
            } // if (groups.Count > 0)
            else if (prevGroups != null && prevGroups.Length > 0) // if enduser unselect all options, remove user from site group 
            {
                AddOrRemoveUserToGroup(prevGroups, UserOpType.RemoveUser, clientContext, userID);
            }
        }

        [HttpPost]
        public ActionResult AddEmployeeToSPList(EmployeeRegistration.MVCWeb.Models.Employee model)
        {
            string SPHostUrl = Request.QueryString["SPHostUrl"];

            StringBuilder sbSkills = new StringBuilder();
            StringBuilder sbAttachments = new StringBuilder();

            if (!model.isFileUploaded)
            {
                model.AttachmentID = "";
            }

            foreach (var skill in model.Skills)
            {
                sbSkills.Append(skill.Technology).Append(",").Append(skill.Experience).Append(";");
            }

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    var lstEmployee = clientContext.Web.Lists.GetByTitle("Employees");
                    var itemCreateInfo = new ListItemCreationInformation();
                    var listItem = lstEmployee.AddItem(itemCreateInfo);

                    listItem["EmpNumber"] = ConvertToEmptyString(model.EmpNumber);
                    listItem["Title"] = ConvertToEmptyString(model.Name);
                    listItem["UserID"] = ConvertToEmptyString(model.UserID);
                    listItem["EmpManager"] = ConvertToEmptyString(model.EmpManager);
                    listItem["Designation"] = ConvertToEmptyString(model.Designation);
                    listItem["Location"] = ConvertToEmptyString(model.Location);
                    listItem["Skills"] = ConvertToEmptyString(sbSkills.ToString());
                    listItem["AttachmentID"] = model.AttachmentID;

                    listItem.Update();
                    clientContext.ExecuteQuery();

                    AddUserToSelectedSiteGroups(model, clientContext, model.UserID);
                }
            } // using (var clientContext 

            return RedirectToAction("Thanks", new { SPHostUrl });
        }

        [HttpPost]
        public void UpdateEmployeeToSPList(EmployeeRegistration.MVCWeb.Models.Employee model)
        {
            string employeeListName = "Employees";
            string SPHostUrl = string.Format("{0}/Lists/{1}", Request.QueryString["SPHostUrl"], employeeListName);
            StringBuilder sbSkills = new StringBuilder();

            foreach (var skill in model.Skills)
            {
                sbSkills.Append(skill.Technology).Append(",").Append(skill.Experience).Append(";");
            }

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    List lstEmployee = clientContext.Web.Lists.GetByTitle(employeeListName);                   
                    var listItem = lstEmployee.GetItemById(model.Id);

                    listItem["EmpNumber"] = model.EmpNumber;
                    listItem["Title"] = model.Name;
                    listItem["UserID"] = model.UserID;
                    listItem["EmpManager"] = model.EmpManager;
                    listItem["Designation"] = model.Designation;
                    listItem["Location"] = model.Location;
                    listItem["Skills"] = sbSkills.ToString();

                    if (model.isFileUploaded)
                    {
                        listItem["AttachmentID"] = model.AttachmentID;
                    }

                    listItem.Update();
                    clientContext.ExecuteQuery();

                    AddUserToSelectedSiteGroups(model, clientContext, model.UserID);
                }
            } // using (var clientContext 

            Response.Redirect(SPHostUrl);
        }

        public ActionResult Thanks(string SPHostUrl)
        {
            ViewBag.Message = "Registered employee successfully";

            return View();
        }       

        [SharePointContextFilter]
        public JsonResult EmpStates(string selectedID)
        {
            List<SelectListItem> stateList = new List<SelectListItem>();

            try
            {
                SharePointContext spContext = Session["Context"] as SharePointContext;
                if (spContext == null)
                {
                    spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
                }

                using (var clientContext = spContext.CreateUserClientContextForSPHost())
                {
                    if (clientContext != null)
                    {
                        var lstState = clientContext.Web.Lists.GetByTitle("EmpState");
                        CamlQuery query = new CamlQuery();
                        query.ViewXml = string.Format("<View><Query><Where><Eq><FieldRef Name='Country' /><Value Type='String'>{0}</Value></Eq></Where></Query></View>", selectedID);

                        ListItemCollection stateItems = lstState.GetItems(query);
                        clientContext.Load(stateItems);
                        clientContext.ExecuteQuery();

                        foreach (var item in stateItems)
                        {
                            stateList.Add(new SelectListItem { Text = item["Title"].ToString(), Value = item["ID"].ToString() });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                stateList.Add(new SelectListItem { Text = "Error: " + ex.Message, Value = "-1" });
            }

            return Json(new SelectList(stateList, "Value", "Text"), JsonRequestBehavior.AllowGet);
        }

        [SharePointContextFilter]
        public JsonResult EmpCities(string selectedID)
        {
            List<SelectListItem> cityList = new List<SelectListItem>();

            try
            {
                SharePointContext spContext = Session["Context"] as SharePointContext;
                if (spContext == null)
                {
                    spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
                }

                using (var clientContext = spContext.CreateUserClientContextForSPHost())
                {
                    if (clientContext != null)
                    {
                        var lstCity = clientContext.Web.Lists.GetByTitle("EmpCity");
                        CamlQuery query = new CamlQuery();
                        query.ViewXml = string.Format("<View><Query><Where><Eq><FieldRef Name='State' /><Value Type='String'>{0}</Value></Eq></Where></Query></View>", selectedID);

                        ListItemCollection cityItems = lstCity.GetItems(query);
                        clientContext.Load(cityItems);
                        clientContext.ExecuteQuery();

                        foreach (var item in cityItems)
                        {
                            cityList.Add(new SelectListItem { Text = item["Title"].ToString(), Value = item["Title"].ToString() });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                cityList.Add(new SelectListItem { Text = "Error: " + ex.Message, Value = "-1" });
            }

            return Json(new SelectList(cityList, "Text", "Text"), JsonRequestBehavior.AllowGet);
        }


        [SharePointContextFilter]
        public JsonResult GetNameAndManagerFromProfile(string userID)
        {
            string empName = string.Empty;
            string empManager = string.Empty;

            try
            {
                SharePointContext spContext = Session["Context"] as SharePointContext;
                if (spContext == null)
                {
                    spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
                }

                using (var clientContext = spContext.CreateUserClientContextForSPHost())
                {
                    if (clientContext != null)
                    {
                        string[] propertyNames = { "FirstName", "LastName", "Manager" };

                        PeopleManager peopleManager = new PeopleManager(clientContext);
                        UserProfilePropertiesForUser prop = new UserProfilePropertiesForUser(clientContext, userID, propertyNames);
                        IEnumerable<string> profileProperty = peopleManager.GetUserProfilePropertiesFor(prop);
                        clientContext.Load(prop);
                        clientContext.ExecuteQuery();

                        if (profileProperty != null && profileProperty.Count() > 0)
                        {
                            empName = string.Format("{0} {1}", profileProperty.ElementAt(0), profileProperty.ElementAt(1));
                            empManager = profileProperty.ElementAt(2);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                empManager = ex.Message;
            }

            return Json(new { Name = empName, Manager = empManager }, JsonRequestBehavior.AllowGet);
        }

        [SharePointContextFilter]
        public JsonResult UploadAttachment(HttpPostedFileBase file, string attachmentID)
        {
            string fileName = file.FileName;
            string newFileName = string.Empty;
            string fileRelativeUrl = string.Empty;

            SharePointContext spContext = Session["Context"] as SharePointContext;
            if (spContext == null)
            {
                spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            }

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    using (var fs = file.InputStream)
                    {
                        FileInfo fileInfo = new FileInfo(fileName);
                        newFileName = string.Format("{0}{1}", Guid.NewGuid(), fileInfo.Extension);

                        List attachmentLib = clientContext.Web.Lists.GetByTitle("EmpAttachments");
                        Folder attachmentLibFolder = attachmentLib.RootFolder;
                        clientContext.Load(attachmentLibFolder);
                        clientContext.ExecuteQuery();
                        fileRelativeUrl = String.Format("{0}/{1}", attachmentLibFolder.ServerRelativeUrl, newFileName);

                        var fileCreationInformation = new FileCreationInformation();
                        fileCreationInformation.ContentStream = fs;
                        fileCreationInformation.Url = fileRelativeUrl;

                        Microsoft.SharePoint.Client.File uploadFile = attachmentLibFolder.Files.Add(fileCreationInformation);
                        uploadFile.ListItemAllFields["Title"] = fileName;
                        uploadFile.ListItemAllFields["AttachmentID"] = attachmentID;
                        uploadFile.ListItemAllFields.Update();
                        clientContext.ExecuteQuery();
                    }
                }
            }

            return Json(new { FileName = fileName, NewFileName = newFileName, FileRelativeURl = fileRelativeUrl });
        }

        [SharePointContextFilter]
        public JsonResult DeleteAttachment(string fileRelativeUrl)
        {
            try
            {
                SharePointContext spContext = Session["Context"] as SharePointContext;
                if (spContext == null)
                {
                    spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
                }

                using (var clientContext = spContext.CreateUserClientContextForSPHost())
                {
                    if (clientContext != null)
                    {
                        Microsoft.SharePoint.Client.File fileToDelete = clientContext.Web.GetFileByServerRelativeUrl(fileRelativeUrl);
                        clientContext.Load(fileToDelete);
                        fileToDelete.DeleteObject();
                        clientContext.ExecuteQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + "\n" + ex.StackTrace);
            }

            return Json("Attachment Deleted");
        } // public void DeleteAttachment
    }
}