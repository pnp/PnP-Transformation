# Pattern: Populating fields on form load - read list data #
This pattern shows how to programmatically read list item data on form load.

## Single Page Application using Knockout.js ##
When edit item is clicked in the list, it redirects to the asp.net provider hosted app page and loads the data of list item.

The loading code is in the `loadEditFormData` JavaScript function inside the `EmpViewModel` JavaScript function:

```JavaScript
self.loadEditFormData = function () {
	var listURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('" + employeeListname + "')/items(" + self.ID() + ")";
	$.ajax({
	    url: listURL,
	    type: "GET",
	    headers: { "accept": "application/json;odata=verbose" },
	    success: function (data) {
	        self.EmpNumber(data.d.EmpNumber);
	        self.CurrentUser(data.d.UserID);
	        self.UserID(data.d.UserID.replace('i:0#.f|membership|', ''));
	        self.Name(data.d.Title);
	        self.EmpManager(data.d.EmpManager);
	        self.Designation(data.d.Designation);
	        self.loadLocationDropDowns(data.d.Location);
	        self.loadSkills(data.d.Skills);
	        self.loadAttachments(data.d.AttachmentID);
	    },
	    error: function (error) {
	        alert(JSON.stringify(error));
	    }
	});
};  
```

As a result, App will read the list item and loads in a single app page.  
The below example shows the existing list item in a Sharepoint list.

[imgReadListItemData]: images/Common/P3_ReadListItemData.png
![][imgReadListItemData]

This data is then set into the App default page, as shown in the figure

![](images/KO/P3_SetListDataToForm.png)


## ASP.Net MVC approach ##
When edit item is clicked in the list, it redirects to the MVC provider hosted app page and loads the data of list item.

The submit code is in the `EmployeeController` inside method `EmployeeForm`:

```C#
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
clientContext.ExecuteQuery();List<SelectListItem> empDesgList = new List<SelectListItem>();
foreach (var item in desgItems)
{
	empDesgList.Add(new SelectListItem { Text = item["Title"].ToString() });
}
emp.Designations = new SelectList(empDesgList, "Text", "Text");List<SelectListItem> cList = new List<SelectListItem>();
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
```

As a result, App will read the list item and it gets loaded in the MVC app page.
The below example shows the existing list item in a Sharepoint list.

![][imgReadListItemData]

This data is then set into the App default page, as shown in the figure

![](images/MVC/P3_SetListDataToForm.png)


## ASP.Net Forms approach ##
When edit item is clicked in the list, it redirects to the asp.net provider hosted app page and loads the data of list item.

In `Default.aspx.cs` there the method `LoadListItems` that implements the save logic:

```C#
string itemID = Request.QueryString["itemId"];
if (itemID != null)
{
	List lstEmployee = web.Lists.GetByTitle("Employees");
	var emplistItem = lstEmployee.GetItemById(itemID);
	clientContext.Load(emplistItem);
	clientContext.ExecuteQuery();
	if (emplistItem != null)
	{
	    txtEmpNumber.Text = emplistItem["EmpNumber"].ToString();
	    txtName.Text = emplistItem["Title"].ToString();
	    txtUserID.Text = emplistItem["UserID"].ToString();
	    txtManager.Text = emplistItem["EmpManager"].ToString();
		ddlDesignation.SelectedValue = emplistItem["Designation"].ToString();
	    string cityVal = emplistItem["Location"].ToString();
		string stateVal = GetStateValFromCity(clientContext, web, cityVal);
		string countryVal = GetCountryValFromState(clientContext, web, stateVal);
		if (countryVal != "")
	     {
	        ddlCountry.SelectedValue = countryVal;
	        LoadStateItems();
	        ddlState.SelectedValue = stateVal;
	        LoadCityItems();
	        ddlCity.SelectedValue = cityVal;
	     }
	
	    LoadSkills(emplistItem["Skills"].ToString());
	    btnUpdate.Visible = true;
	} // if (emplistItem != null)
}
```

As a result, App will read the list item and load all the data in a asp.net forms app page.
As a result, App will read the list item and it gets loaded in the MVC app page.
The below example shows the existing list item in a Sharepoint list.

![][imgReadListItemData]

This data is then set into the App default page, as shown in the figure

![](images/Forms/P3_SetListDataToForm.png)

