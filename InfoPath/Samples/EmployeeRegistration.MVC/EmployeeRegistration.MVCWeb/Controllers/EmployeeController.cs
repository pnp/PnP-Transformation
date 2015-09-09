using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

using EmployeeRegistration.MVCWeb.Models;

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

        [SharePointContextFilter]
        public ActionResult EmployeeForm()
        {
            Employee emp = new Employee();
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

                    PeopleManager peopleManager = new PeopleManager(clientContext);
                    PersonProperties personProperties = peopleManager.GetMyProperties();

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
                    clientContext.Load(personProperties, p => p.AccountName);
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
                        emp.EmpNumber = emplistItem["EmpNumber"].ToString();
                        emp.Name = emplistItem["Title"].ToString();
                        emp.UserID = emplistItem["UserID"].ToString();
                        emp.EmpManager = emplistItem["EmpManager"].ToString();
                        emp.Designation = emplistItem["Designation"].ToString();

                        string cityVal = emplistItem["Location"].ToString();
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

                        string skillsData = emplistItem["Skills"].ToString();
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
                        emp.ActionName = "UpdateEmployeeToSPList";
                        emp.SubmitButtonName = "Update Employee";
                    }
                    else
                    {
                        if (personProperties != null && personProperties.AccountName != null)
                        {
                            //ViewBag.UserID = personProperties.AccountName;
                            emp.UserID = personProperties.AccountName;
                        }
                        List<Skill> lsSkills = new List<Skill>();
                        lsSkills.Add(new Skill { Technology = "", Experience = "" });
                        emp.Skills = lsSkills;
                        emp.SkillsCount = lsSkills.Count;
                        emp.ActionName = "AddEmployeeToSPList";
                        emp.SubmitButtonName = "Add Employee";
                    }



                } //  if (clientContext != null)
            } // using (var clientContext

            ViewBag.SPURL = Request.QueryString["SPHostUrl"];

            return View(emp);
        }


        [HttpPost]

        public ActionResult AddEmployeeToSPList(EmployeeRegistration.MVCWeb.Models.Employee model)
        {
            string number = model.EmpNumber;
            string name = model.Name;
            string designation = model.Designation;
            string SPHostUrl = Request.QueryString["SPHostUrl"];
            string userName = string.Empty;
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
                    var lstEmployee = clientContext.Web.Lists.GetByTitle("Employees");
                    var itemCreateInfo = new ListItemCreationInformation();
                    var listItem = lstEmployee.AddItem(itemCreateInfo);

                    listItem["EmpNumber"] = model.EmpNumber;
                    listItem["Title"] = model.Name;
                    listItem["UserID"] = model.UserID;
                    listItem["EmpManager"] = model.EmpManager;
                    listItem["Designation"] = model.Designation;
                    listItem["Location"] = model.Location;
                    listItem["Skills"] = sbSkills.ToString();

                    listItem.Update();
                    clientContext.ExecuteQuery();
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

                    listItem.Update();
                    clientContext.ExecuteQuery();
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
    }
}