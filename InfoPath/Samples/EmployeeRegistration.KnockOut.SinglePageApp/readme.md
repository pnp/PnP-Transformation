# InfoPath form alternative - Single Page Application #

### Summary ###
This sample shows an application that's leveraging JavaScript and HTML to offer an alternative for an InfoPath form. The application is using the [knockoutjs](http://knockoutjs.com/) library to realize a dynamic JavaScript UI.

### Applies to ###
-  Office 365 Multi Tenant (MT)
-  Office 365 Dedicated (D)
-  SharePoint 2013 on-premises

### Solution ###
Solution | Author(s)
---------|----------
EmployeeRegistration.KnockOut.SinglePageApp | Raja Shekar Bhumireddy, Bert Jansen (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | September 7th 2015 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Introduction #
The core of this application are the JavaScript based view-model and view. This files are depending on [knockoutjs](http://knockoutjs.com/) to implement a Model-View-View Model (MVVM) pattern. The UI is defined using views and declarative bindings, its data and behavior using view-models and observables, and everything stays in sync automatically thanks to Knockout's dependency tracking. The model (=data) itself is stored in SharePoint lists.

To learn more about the Knockout JavaScript library check these resources:
- [Tutorials](http://learn.knockoutjs.com/)
- [Live examples](http://knockoutjs.com/examples/)
- [Documentation](http://knockoutjs.com/documentation/introduction.html)

The next chapters will explain the application architecture, the core components of the application and the installation process. 

![](http://i.imgur.com/LqTW3pH.png)

# Application architecture #
The below diagram shows the application architecture used for this sample:
1. All data is stored in SharePoint lists inside the site. Employees is the list that will be populated by the application, whereas the other lists contain supporting. Notice that the Countries, States and Cities list are depending on each other via the use of lookup fields. This depending data model is also applied to the application itself.
2. The application is nothing more then a bunch of JavaScript files which are deployed into the sites Style Library.
3. The main application runs on a page called empform.aspx. This wiki page contains a single content editor webpart that triggers the load of the `Emp-Registration-Form-Template.js` script.
4. Since the data is stored in a SharePoint list we want to also override the default list forms: the edit and new list form pages are updated by closing the default form web part on those pages and adding a content editor web part that loads the `emp-newform.js` or `emp-newform.js` script.

![](http://i.imgur.com/xrqAODM.png)

# Application components #
## Knockout views ##
The view is implemented in the `Emp-Registration-Form-Template.js` JavaScript file and uses specific Knockout elements. The file starts with referencing the JavaScript file that are needed to make the sample work.

```JavaScript
<script type="text/javascript" src="~sitecollection/Style%20Library/OfficeDevPnP/jquery-1.10.2.min.js"></script>
<script type="text/javascript" src="~sitecollection/Style%20Library/OfficeDevPnP/knockout-3.3.0.js"></script>
<script type="text/javascript" src="~sitecollection/Style%20Library/OfficeDevPnP/Emp-Registration-Form.js"></script>
```

Note that the `~sitecollection` token will be replaced during application installation with the actual site url. As one can see from the references the files will be loaded from the sites `Style Library`. 

Below code contains the element that's put on the page, which is a single div. Using Knockout this div is dynamically transformed into an application by loading a template with the name `EmpRegistrationForm-Template`.

```JavaScript
<div data-bind="template: {name: 'EmpRegistrationForm-Template'}" id="EmpRegistrationForm">Loading...</div>
```

The template that's being loaded is listed below. A template is implemented using a `script`tag. Inside that `script` tag you'll see a mix of Knockout specific syntax and HTML. If you want to understand the details of this specific Knockout syntax then please go checkout the Knockout resources listed earlier on this page.

```JavaScript
<script type="text/html" id="EmpRegistrationForm-Template">
    <!-- ko if:canSave -->
    <p>Saving employee information...</p>
    <!-- /ko -->
    <!-- ko if:canUpdate -->
    <p>Updating employee information...</p>
    <!-- /ko -->
    <!-- ko ifnot:canSave -->
        <!-- ko ifnot:isFormLoaded -->
            <p>Loading form...</p>
        <!-- /ko -->
        <!-- ko if:isFormLoaded -->
            <table border="0" cellspacing="2" cellpadding="2" class="ms-listviewtable">
                <tr>
                    <th style="text-align: right;">Emp Number</th>
                    <td><input data-bind="value: EmpNumber"/></td>
                </tr>
                <tr>
                    <th style="text-align: right;">User ID</th>
                    <td>
                        <input data-bind="value: UserID"/><button data-bind="click: getNameAndManagerFromProfile">Get name and manager from profile</button>
                    </td>
                </tr>
                <tr>
                    <th style="text-align: right;">Name</th>
                    <td><input data-bind="value: Name"/></td>
                </tr>
                <tr>
                    <th style="text-align: right;">Emp Manager</th>
                    <td><input data-bind="value: EmpManager"/></td>
                </tr>
                <tr>
                    <th style="text-align: right;">Designation</th>
                    <td><select data-bind="options: Designations, value: Designation, optionsCaption: 'Select Designation'"></select></td>
                </tr>
                <tr>
                    <th style="text-align: right;">Location</th>
                    <td>
                        <select data-bind="options: Countries, optionsValue:'CountryId', optionsText: 'CountryName', value: Country, optionsCaption: 'Select Country'"></select>
                        <select data-bind="options: States, optionsValue:'StateId', optionsText: 'StateName', value: State, optionsCaption: 'Select State'"></select>
                        <select id="selCityLocation" data-bind="options: Cities, optionsValue: 'CityName', optionsText: 'CityName', value: Location, optionsCaption: 'Select City'"></select>
                    </td>
                </tr>
                <tr>
                    <th style="text-align: right; vertical-align: top;">Skills</th>
                    <td>
                        <table>
                            <tr>
                                <th>Technology</th>
                                <th>experience</th>
                            </tr>
                            <tbody data-bind="foreach: Skills">
                                <tr>
                                    <td><input data-bind='value: technology' /></td>
                                    <td><input data-bind='value: experience' /></td>
                                </tr>
                            </tbody>
                        </table>
                        <a href='#' data-bind='click: $root.addSkill'>Add new skill</a>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <!-- ko if:isNewForm -->
                        <p><button data-bind="click: save">Save</button></p>
                        <!-- /ko -->
                        <!-- ko ifnot:isNewForm -->
                        <p><button data-bind="click: update">Update</button></p>
                        <!-- /ko -->
                    </td>
                </tr>
            </table>
        <!-- /ko -->
    <!-- /ko -->
</script>
```

## Knockout view-model ##
Just like the view the view-model is also implemented using JavaScript. In this sample the model is named `Emp-Registration-Form.js`. When this script is loaded it will load the `initEmpForm` function:

```JavaScript
$(document).ready(function () {
    initEmpForm();
});
```

in the `initEmpForm` function we load the model

```JavaScript
function initEmpForm() {
    JSRequest.EnsureSetup();
    var itemId = JSRequest.QueryString["itemId"];

    var empModel = new EmpViewModel();
    isDesignationsLoaded = $.Deferred();
    empModel.initNewFormData();
    
    if (typeof itemId != "undefined") {
        empModel.ID(itemId);
        $.when(isDesignationsLoaded).then(empModel.loadEditFormData());
        empModel.isNewForm(false);
    }
    else {
        empModel.getCurrentUser();
        empModel.addSkill();
        empModel.isFormLoaded(true);
    }

    ko.applyBindings(empModel, document.getElementById('EmpRegistrationForm'));
}
```

In `EmpViewModel` you'll see methods and observable properties which allow Knockout to track changes and update the UI accordingly. The Observable properties are defined like this:

```JavaScript
// generic view properties
self.ID = ko.observable();
self.isFormLoaded = ko.observable(false);
self.isNewForm = ko.observable(true);
self.canSave = ko.observable(false);
self.canUpdate = ko.observable(false);

// data view properties
self.Name = ko.observable();
self.EmpNumber = ko.observable();
self.Designation = ko.observable();
self.Designations = ko.observableArray([]),
self.Location = ko.observable();
self.Country = ko.observable();
self.Countries = ko.observableArray([]);
self.State = ko.observable();
self.States = ko.observableArray([]);
self.Cities = ko.observableArray([]);
self.Skills = ko.observableArray([]);
self.UserID = ko.observable();    
self.EmpManager = ko.observable();
self.CountryID = ko.observable();
```

There are multiple methods in the view-model, below code is only showing the `update` method:

```JavaScript
// Update an existing item
self.update = function () {
    self.canUpdate(true);
    var listURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('" + employeeListname + "')/items(" + self.ID() + ")";
            
    $.ajax({
        url: listURL,
        type: "POST",
        headers: {
            "accept": "application/json; odata=verbose", "content-type": "application/json; odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val(),
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*"
        },
        data: JSON.stringify(self.getEmployeeFormData()),
        success: function (data) {
            self.redirectToList();
        },
        error: function (error) {
            alert(JSON.stringify(error));
        }
    });
    
};
```

## Application installation ##
Since this application is simply a set of JavaScript files the solution contains C# code that's used to create the needed lists with sample data, deploy all the JavaScript files to the sites `Style Library`, add a wiki page where we use the content editor web part to run the application and finally update the existing `NewForm.aspx` and `EditForm.aspx` pages of the `employee` list. Below methods are called to realize the described installation steps:

```C#
ClientContext ctx = CreateContext();
// Provision supporting js files to the Style Library
ProvisionAssets(ctx);

// Provision lists and items
ProvisionLists(ctx);

// Provision the employee registration application
ProvisionEmployeeRegistrationApplication(ctx);
```

<img src="https://telemetry.sharepointpnp.com/pnp-transformation/infopath/samples/EmployeeRegistration.KnockOut.SinglePageApp" /> 