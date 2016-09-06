# InfoPath form alternative - ASP.Net MVC #

### Summary ###
This sample shows an application that's leveraging ASP.Net MVC to offer an alternative for an InfoPath form.

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
1.0  | September 9th 2015 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Introduction #
This sample is implementing the functionality of the reference InfoPath form using plain (old) ASP.Net web forms. 

![](http://i.imgur.com/xWZJwLH.png)

# Application architecture #
The below diagram shows the application architecture used for this sample:
1. All data is stored in SharePoint lists inside the site. Employees is the list that will be populated by the application, whereas the other lists contain supporting. Notice that the Countries, States and Cities list are depending on each other via the use of lookup fields. This depending data model is also applied to the application itself.
2. The application is an ASP.Net MVC provider hosted SharePoint Add-In. The main functionality is implemented in the employee controller (`EmployeeController.cs`), view (`EmployeeForm.cshtml`) and JavaScript (`employee.js`).
3. The `Style Library` holds a script that will be used to redirect from any SharePoint site toward the provider hosted application. In this redirect we optionally pass the item to be loaded, which will be honored by the provider hosted application.
4. Since the data is stored in a SharePoint list we want to also override the default list forms: the edit and new list form pages are updated by closing the default form web part on those pages and adding a content editor web part that loads the `applauncher.js` script which we explained in step 3.

![](http://i.imgur.com/vgJBdjo.png)

# Application components #

## Implementation of "Edit" ##
In the `EmployeeForm` action in `EmployeeController.cs` we check for the existence of the `itemId` query string parameter and based on that the item to be edited will be loaded or not. Below code snippet shows this:

```C#
[SharePointContextFilter]
public ActionResult EmployeeForm()
{
    Employee emp = new Employee();
    var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

    using (var clientContext = spContext.CreateUserClientContextForSPHost())
    {
        if (clientContext != null)
        {

            // code removed for clarity

            string itemID = Request.QueryString["itemId"];
            ListItem emplistItem = null;

            if (itemID != null)
            {
                List lstEmployee = web.Lists.GetByTitle("Employees");
                emplistItem = lstEmployee.GetItemById(itemID);
                clientContext.Load(emplistItem);
                emp.Id = itemID;
            }

            // code removed for clarity

            if (emplistItem != null)
            {
                // code removed for clarity
            }
            else
            {
                // code removed for clarity
            }
        } //  if (clientContext != null)
    } // using (var clientContext

    ViewBag.SPURL = Request.QueryString["SPHostUrl"];

    return View(emp);
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

The code for doing this is ran from the `EmployeeForm` action in `EmployeeController.cs`:

```C#
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

            // code removed for clarity

        } //  if (clientContext != null)
    } // using (var clientContext

    ViewBag.SPURL = Request.QueryString["SPHostUrl"];

    return View(emp);
}
```

<img src="https://telemetry.sharepointpnp.com/pnp-transformation/infopath/samples/EmployeeRegistration.MVC" /> 
