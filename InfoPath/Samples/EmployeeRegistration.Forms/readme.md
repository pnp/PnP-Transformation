# InfoPath form alternative - ASP.Net Web Forms #

### Summary ###
This sample shows an application that's leveraging ASP.Net web forms to offer an alternative for an InfoPath form.

### Applies to ###
-  Office 365 Multi Tenant (MT)
-  Office 365 Dedicated (D)
-  SharePoint 2013 on-premises


### Solution ###
Solution | Author(s)
---------|----------
EmployeeRegistration.Forms | Raja Shekar Bhumireddy, Bert Jansen (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | September 8th 2015 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Introduction #
This sample is implementing the functionality of the reference InfoPath form using plain (old) ASP.Net web forms. 

![](http://i.imgur.com/96OyeMB.png)

# Application architecture #
The below diagram shows the application architecture used for this sample:
1. All data is stored in SharePoint lists inside the site. Employees is the list that will be populated by the application, whereas the other lists contain supporting. Notice that the Countries, States and Cities list are depending on each other via the use of lookup fields. This depending data model is also applied to the application itself.
2. The application is an ASP.Net provider hosted SharePoint Add-In. The main functionality is implemented in `default.aspx`
3. The `Style Library` holds a script that will be used to redirect from any SharePoint site toward the provider hosted application. In this redirect we optionally pass the item to be loaded, which will be honored by the provider hosted application.
4. Since the data is stored in a SharePoint list we want to also override the default list forms: the edit and new list form pages are updated by closing the default form web part on those pages and adding a content editor web part that loads the `applauncher.js` script which we explained in step 3.

![](http://i.imgur.com/D4Gb2P8.png)

# Application components #
## Chrome control ##
The application uses 2 aspx pages: `default.aspx` and `thanks.aspx`. Both of these pages have the SharePoint Chrome control implemented via the injection of following DIV:

```ASP
<div id="divSPChrome"></div>
```

When default.aspx is loaded we run the `renderSPChrome` JavaScript function like implemented in `app.js`:


```JavaScript
$(document).ready(function () {
    //Get the URI decoded SharePoint site url from the SPHostUrl parameter.
    var spHostUrl = decodeURIComponent(getQueryStringParameter('SPHostUrl'));

    //Build absolute path to the layouts root with the spHostUrl
    var layoutsRoot = spHostUrl + '/_layouts/15/';

    //load all appropriate scripts for the page to function
    $.getScript(layoutsRoot + 'SP.Runtime.js',
        function () {
            $.getScript(layoutsRoot + 'SP.js',
                function () {
                    //Execute the correct script based on the isDialog
                    //Load the SP.UI.Controls.js file to render the App Chrome
                    $.getScript(layoutsRoot + 'SP.UI.Controls.js', renderSPChrome);
                });
        });
});
```

The `renderSPChrome` JavaScript function itself is inserted on page load:

```C#
protected void Page_Load(object sender, EventArgs e)
{
    // define initial script, needed to render the chrome control
    string script = @"
    function chromeLoaded() {
        $('body').show();
    }

    //function callback to render chrome after SP.UI.Controls.js loads
    function renderSPChrome() {
        //Set the chrome options for launching Help, Account, and Contact pages
        var options = {
            'appTitle': document.title,
            'onCssLoaded': 'chromeLoaded()'
        };

        //Load the Chrome Control in the divSPChrome element of the page
        var chromeNavigation = new SP.UI.Controls.Navigation('divSPChrome', options);
        chromeNavigation.setVisible(true);
    }";

    //register script in page
    Page.ClientScript.RegisterClientScriptBlock(typeof(Default), "BasePageScript", script, true);
}
```

## Implementation of "Edit" ##
When the provider hosted application is launched with an URL indicating the item to load then this will trigger the application to load the provided item instead of showing an empty UI for a new item add. In the `LoadListItems` method this is realized via the following construct:

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
else
{
    PeopleManager peopleManager = new PeopleManager(clientContext);
    PersonProperties personProperties = peopleManager.GetMyProperties();
    clientContext.Load(personProperties, p => p.AccountName);

    clientContext.ExecuteQuery();
    if (personProperties != null && personProperties.AccountName != null)
    {
        txtUserID.Text = personProperties.AccountName;
    }

    AddEmptySkill();

    btnSave.Visible = true;
}
```

## Redirect to the application from the Employee list edit and new forms ##
Whenever a user goes the `Employee` list and clicks on `Add` or `Edit` he will navigate to the provider hosted application. This is realized by adding a content editor web part on the out of the box `newform.aspx` and `editform.aspx` pages for the `Employees` list. The added content editor web part will trigger the execution of `AppLaucher.js` which takes care of the redirect to the app:

```JavaScript
<script type="text/javascript">
JSRequest.EnsureSetup();
var displayMode = JSRequest.QueryString['DisplayMode'];
if (typeof displayMode == 'undefined') {
    var itemId = JSRequest.QueryString['ID']; 
    var itemData = ""; 
    if (typeof itemId != 'undefined') { 
        itemData = "%26itemId=" + itemId; 
    } 
    var redirectURL = "~sitecollection/_layouts/15/appredirect.aspx?client_id={%clientId%}&redirect_uri=%redirectURI%?{StandardTokens}" + itemData;
    
    window.location = redirectURL;
} 
</script>
```

The `~siteCollection`, `%client_id%` and `%redirect_uri%` variables will be replaced by actual values during the app installation.

## Application installation ##
When the application is first run it checks if all needed pre-requisites are present and if not it will create/implement these:
- Creating the lists with dummy data
- Deploying the `AppLauncher.js` script to the `Style Library`
- Updating the `newform.aspx` and `editform.aspx` pages for the `Employees` list

The code for doing this is ran from the `Page_Load` method:

```C#
if (!Page.IsPostBack)
{
    // Provision supporting artefacts in case that's still needed
    using (var clientContext = GetClientContext())
    {
        if (!SetupManager.Initialized)
        {
            // Provision lists
            SetupManager.ProvisionLists(clientContext);

            // upload assets and provision the application
            ProvisionEmployeeRegistrationApplication(clientContext);

            SetupManager.Initialized = true;
        }
    }
} 
```

<img src="https://telemetry.sharepointpnp.com/pnp-transformation/infopath/samples/EmployeeRegistration.Forms" /> 

