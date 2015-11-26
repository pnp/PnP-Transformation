var employeeListname = "Employees";
var designationListname = "EmpDesignation";
var countryListName = "EmpCountry";
var stateListName = "EmpState";
var cityListName = "EmpCity";

var isCountryIDLoaded;
var isStatesLoaded;
var isCitiesLoaded;
var isDesignationsLoaded;

function getListItemType(name) {
    return "SP.Data." + name[0].toUpperCase() + name.substring(1) + "ListItem";
}

function getFileBuffer(file) {
    var deferred = jQuery.Deferred();
    var reader = new FileReader();
    reader.onloadend = function (e) {
        deferred.resolve(e.target.result);
    }
    reader.onerror = function (e) {
        deferred.reject(e.target.error);
    }
    reader.readAsArrayBuffer(file);
    return deferred.promise();
}

function uploadDocumentInSharePointLibrary(arrayBuffer, libraryName, fileName) {
    var url = _spPageContextInfo.webAbsoluteUrl + "/_api/web/getfolderbyserverrelativeurl('Lists/" + libraryName + "')/files/add(url='" + fileName + "')";
    
    var uploadDoc = jQuery.ajax({
        url: url,
        method: "POST",
        data: arrayBuffer,
        processData: false,
        contentType: "application/json;odata=verbose",
        headers: {
            "accept": "application/json;odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val()
        }
    });

    return uploadDoc;
}

function getUploadedAttachmentItemProperties(file) {
    var attachmentItem = jQuery.ajax({
        url: file.ListItemAllFields.__deferred.uri,
        type: "GET",
        dataType: "json",
        headers: {
            Accept: "application/json;odata=verbose"
        }
    });

    return attachmentItem;
}

function updateUploadedDocItemFields(item, libraryName, fileName, attachmentID) {
    var url = _spPageContextInfo.webAbsoluteUrl + "/_api/web/Lists/getByTitle('" + libraryName + "')/Items(" + item.Id + ")";
    
   var updateFields = jQuery.ajax({
        url: url,
        type: "POST",
        data: JSON.stringify({
            "__metadata": { type: getListItemType(libraryName) },
            Title: fileName,
            AttachmentID: attachmentID
        }),
        headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
            "IF-MATCH": item.__metadata.etag,
            "X-Http-Method": "MERGE"
        }
    });

   return updateFields;
}

function failHandler(jqXHR, textStatus, errorThrown, btnForm) {
    btnForm.disabled = false;
    var response = JSON.parse(jqXHR.responseText);
    var message = response ? response.error.message.value : textStatus;
    alert("Error: " + message);
}

function EmpViewModel() {
    var self = this;
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
    self.CurrentUser = ko.observable();
    self.EmpManager = ko.observable();
    self.CountryID = ko.observable();
    self.AttachmentID = ko.observable();
    self.EmpAttachment = ko.observable();
    self.Attachments = ko.observableArray([]);
    self.SiteGroups = ko.observableArray([]);
    self.UserSiteGroups = ko.observableArray([]);
    self.PreviouslySeletedSiteGroups = ko.observableArray([]);

    // Get site groups of current user id
    self.getUserGroups = function () {
        var userID = encodeURIComponent(self.CurrentUser());
        var userGroupsURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/siteusers(@v)/groups?@v='" + userID + "'";

        $.ajax({
            url: userGroupsURL,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" },
            success: function (data) {
                var results = data.d.results;
                for (var i = 0; i < results.length; i++) {
                    self.UserSiteGroups.push(results[i].Id);
                    self.PreviouslySeletedSiteGroups.push(results[i].Id);
                }
            },
            error: function (error) {
                alert(JSON.stringify(error));
            }
        });        
    };

    // Get site groups of current web
    self.getSiteGroups = function () {
        var siteGroupURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/sitegroups?$select=Id,Title";

        $.ajax({
            url: siteGroupURL,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" },
            success: function (data) {
                var results = data.d.results;
                for (var i = 0; i < results.length; i++) {
                    self.SiteGroups.push({
                        name: results[i].Title,
                        id: results[i].Id
                    });
                }
            },
            error: function (error) {
                alert(JSON.stringify(error));
            }
        });
    };

    // Add attachment/file info to attachments collection
    self.addAttachment = function (libraryName, newFileName, fileName) {
        var fileRelativeUrl = "/Lists/" + libraryName + "/" + newFileName;
        var attachmentURL = _spPageContextInfo.webAbsoluteUrl + fileRelativeUrl;

        self.Attachments.push({
            attachmentURL: attachmentURL,
            fileName: fileName,
            fileRelativeUrl: newFileName
        });
    };

    // Delete uploaded attachment from library
    self.deleteAttachment = function (fileName, data, event) {
        var deleteFileUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/web/getfolderbyserverrelativeurl('Lists/EmpAttachments')/files('" + fileName + "')";
       
        $.ajax({
            url: deleteFileUrl,
            type: "POST",
            headers: {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                "IF-MATCH": "*",
                "X-Http-Method": "DELETE"
            },
            success: function (response) {
                self.Attachments.remove(data);
            },
            error: function (error) {
                alert(JSON.stringify(error));
            }
        });
    };

    // Upload attachment to library
    self.uploadAttachment = function () {
        var attachment = self.EmpAttachment();
        if (attachment === undefined) {
            alert("Please upload a file");
            return;
        }

        var btnForm;
        // Disable submit button until file gets uploaded to SharePoint library
        if (self.isNewForm()) {
            btnForm = document.getElementById('btnSave');
        }
        else {
            btnForm = document.getElementById('btnUpdate');
        }
        btnForm.disabled = true;

        var libraryName = "EmpAttachments";
        var file = attachment.files[0];
        var fileName = file.name;
        var file_name_array = fileName.split(".");
        var file_extension = file_name_array[file_name_array.length - 1];
        var newFileName = SP.Guid.newGuid().toString() + "." + file_extension;
        
        var getFile = getFileBuffer(file);
        getFile.done(function (arrayBuffer) {
            var uploadDocUsingRestApi = uploadDocumentInSharePointLibrary(arrayBuffer, libraryName, newFileName);
            uploadDocUsingRestApi.done(function (data, textStatus, jqXHR) {
                var uploadedDocItem = getUploadedAttachmentItemProperties(data.d);
                uploadedDocItem.done(function (data, textStatus, jqXHR) {
                    var updateFileProperties = updateUploadedDocItemFields(data.d, libraryName, fileName, self.AttachmentID());
                    updateFileProperties.done(function (data, textStatus, jqXHR) {
                        btnForm.disabled = false;
                        self.addAttachment(libraryName, newFileName, fileName);
                        var inputEmpAttachment = $("#empAttachment");
                        inputEmpAttachment.replaceWith(inputEmpAttachment.val('').clone(true));
                        self.EmpAttachment(undefined);
                    });
                    updateFileProperties.fail(function (jqXHR, textStatus, errorThrown) {
                        failHandler(jqXHR, textStatus, errorThrown, btnForm);
                    });
                });
                uploadedDocItem.fail(function (jqXHR, textStatus, errorThrown) {
                    failHandler(jqXHR, textStatus, errorThrown, btnForm);
                });
            });
            uploadDocUsingRestApi.fail(function (jqXHR, textStatus, errorThrown) {
                failHandler(jqXHR, textStatus, errorThrown, btnForm);
            });

        });
        getFile.fail(function (error) { alert("Error in reading file: " + error.responseText); });
    };

    self.changeEmpAttachmentFileInfo = function (data, event) {
        self.EmpAttachment(event.currentTarget);
    };

    self.addSkill = function () {
        self.Skills.push({
            technology: "",
            experience: ""
        });
    };

    // Get state ID based on the provided city value
    self.getStateIDFromCityList = function (cityValue) {
        var stateID = 0;
        var stateListURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('" + cityListName + "')/items/?$select=Id,StateId&$filter=Title eq '" + cityValue + "'";

        $.ajax({
            url: stateListURL,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" },
            success: function (data) {
                $.each(data.d.results, function (index, item) {
                    stateID = item.StateId;
                });
            },
            error: function (error) {
                alert(JSON.stringify(error));
            },
            async: false
      });

      return stateID;
    };

    // Get the country ID based on the provided state ID value
    self.getCountryIDFromStateList = function (stateID) {
        var countryListURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('" + stateListName + "')/items/?$select=CountryId&$filter=Id eq " + stateID;
        
        $.ajax({
            url: countryListURL,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" },
            success: function (data) {
                $.each(data.d.results, function (index, item) {
                    self.CountryID(item.CountryId);
                });
                isCountryIDLoaded.resolve();
            },
            error: function (error) {
                alert(JSON.stringify(error));
                isCountryIDLoaded.resolve();
            }
        });
    };

    // Load the cities for a given state
    self.loadCities = function (stateID) {
        var stateListURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('" + cityListName + "')/items/?$select=Id,Title,StateId&$filter=State/Id eq " + stateID;
        $.ajax({
            url: stateListURL,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" },
            success: function (data) {
                $.each(data.d.results, function (k, l) {
                    self.Cities.push({ CityName: l.Title });
                });
                isCitiesLoaded.resolve();
            },
            error: function (error) {
                alert(JSON.stringify(error));
                isCitiesLoaded.resolve();
            }
        });
    };

    // Load the states for a given country
    self.loadStates = function (countryID) {
        var stateListURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('" + stateListName + "')/items/?$select=Id,Title&$filter=Country/Id eq " + countryID;
        $.ajax({
            url: stateListURL,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" },
            success: function (data) {
                
                $.each(data.d.results, function (k, l) {
                    self.States.push({ StateId: l.Id, StateName: l.Title });
                });
                isStatesLoaded.resolve();
            },
            error: function (error) {
                alert(JSON.stringify(error));
                isStatesLoaded.resolve();
            }
        });
    };

    // Refresh cities when there's a state value change
    self.State.subscribe(function (newValue) {
        self.Cities.removeAll();

        if (newValue) {
             self.loadCities(newValue);
        }       
    });

    // Refresh the states when there's a country value change
    self.Country.subscribe(function (newValue) {
        self.States.removeAll();
        self.Cities.removeAll()
        
        if (typeof newValue != "undefined") {
            self.loadStates(newValue);
        }
    });

    // transform the skills list into a ; delimited string
    self.skillsToString = function () {
        var empSkills = "";
        $.each(self.Skills(), function (k, l) {
            empSkills += l.technology + "," + l.experience + ";";
        });

        return empSkills;
    }

    // Grab the currently entered data, needed to save it
    self.getEmployeeFormData = function () {
        return {
            "__metadata": { "type": getListItemType(employeeListname) },
            "Title": self.Name(),
            "EmpNumber": self.EmpNumber(),
            "Designation": self.Designation(),
            "Location": self.Location(),
            "Skills": self.skillsToString(),
            "UserID": self.CurrentUser(),
            "EmpManager": self.EmpManager(),
            "AttachmentID": self.AttachmentID()
        };
    };

    // Redirect to the employee list
    self.redirectToList = function () {
        window.location = _spPageContextInfo.webAbsoluteUrl + "/Lists/" + employeeListname;
    };

    // Add current user to site group based on selection made
    self.addUserToSiteGroup = function (siteGroupId) {
        var addUserURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/sitegroups/getbyid(" + siteGroupId + ")/users";
        var userItem = { "__metadata": { "type": 'SP.User' }, "LoginName": self.CurrentUser() };

        $.ajax({
            url: addUserURL,
            type: "POST",
            headers: {
                "accept": "application/json;odata=verbose",
                "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                "content-Type": "application/json;odata=verbose"
            },
            data: JSON.stringify(userItem),
            error: function (error) {
                alert(JSON.stringify(error));
            }
        });
    };

    // Remove current user from site group
    self.removeUserFromSiteGroup = function (siteGroupId) {
        var userID = encodeURIComponent(self.CurrentUser());
        var removeUserURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/sitegroups(" + siteGroupId + ")/users/removebyloginname(@v)?@v='" + userID + "'";

        $.ajax({
            url: removeUserURL,
            type: "POST",
            headers: {
                "Accept": "application/json;odata=verbose",
                "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                "content-Type": "application/json;odata=verbose"
            },
            error: function (error) {
                alert(JSON.stringify(error));
            }
        });
    };

    self.addOrRemoveUserToOrFromSiteGroups = function () {
        var addUsers = $(self.UserSiteGroups()).not(self.PreviouslySeletedSiteGroups());
        for (var i = 0; i < addUsers.length; i++) {
            self.addUserToSiteGroup(addUsers[i]);
        }

        var deleteUsers = $(self.PreviouslySeletedSiteGroups()).not(self.UserSiteGroups());
        for (var i = 0; i < deleteUsers.length; i++) {
            self.removeUserFromSiteGroup(deleteUsers[i]);
        }

        return true;
    };
   
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
                $.when(self.addOrRemoveUserToOrFromSiteGroups()).done(function () {
                    self.redirectToList();
                });
            },
            error: function (error) {
                alert(JSON.stringify(error));
            }
        });
        
    };

    // Add a new item
    self.save = function () {
        self.canSave(true);        
        var listURL =  _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('" + employeeListname + "')/items";

        $.ajax({
            url: listURL,
            type: "POST",
            headers: {
                "accept": "application/json;odata=verbose",
                "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                "content-Type": "application/json;odata=verbose"
            },
            data: JSON.stringify(self.getEmployeeFormData()),
            success: function (data) {
                $.when(self.addOrRemoveUserToOrFromSiteGroups()).done(function () {
                    self.redirectToList();
                });
            },
            error: function (error) {
                alert(JSON.stringify(error));
            }
        });

    };

    // Get the currently logged on user
    self.getCurrentUser = function () {
        var currentUserURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/currentUser";

        $.ajax({
            url: currentUserURL,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" },
            success: function (data) {
                self.CurrentUser(data.d.LoginName);
                self.UserID(data.d.LoginName.replace('i:0#.f|membership|', ''));                
            },
            error: function (error) {
                alert(JSON.stringify(error));
            }
        });
    };

    self.CurrentUser.subscribe(function () {
        self.getUserGroups();
    });

    self.getNameAndManagerFromProfile = function () {
        var userID = encodeURIComponent(self.CurrentUser());
        var currentUserURL = _spPageContextInfo.webAbsoluteUrl + "/_api/SP.UserProfiles.PeopleManager/getpropertiesfor(@v)?@v='" + userID + "'&$select=UserProfileProperties";
        
        $.ajax({
            url: currentUserURL,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" },
            success: function (data) {
                var properties = data.d.UserProfileProperties.results;
                var fN, lN;

                properties.forEach(function (property) {
                    if (property.Key == "FirstName") {
                        fN = property.Value;
                    }

                    if (property.Key == "LastName") {
                        lN = property.Value;
                    }

                    if (property.Key == "Manager") {

                        var managerValue = property.Value;
                        managerValue = managerValue.replace('i:0#.f|membership|', '');

                        self.EmpManager(managerValue);
                    }
                });

                self.Name(fN + " " + lN);
            },
            error: function (error) {
                alert(JSON.stringify(error));
            }
        });
    };
        
    // Correctly fill the dropdowns on edit (only valid combinations should be loaded)
    self.loadLocationDropDowns = function (locValue) {
        // Get the ID of the state that match the provided city value
        var stateid = self.getStateIDFromCityList(locValue);

        // Create promises to deal with loading country, states and cities
        isCountryIDLoaded = $.Deferred();
        isStatesLoaded = $.Deferred();
        isCitiesLoaded = $.Deferred();

        // Get the country ID based on the provided state ID value
        self.getCountryIDFromStateList(stateid);
        $.when(isCountryIDLoaded).then(function () {
            // Select the country based on the id
            self.Country(self.CountryID());
        });

        // Select the state based on the state id
        $.when(isStatesLoaded).then(function () {
            self.State(stateid);
        });

        // Select the city based on the id
        $.when(isCitiesLoaded).then(function () {
            self.Location(locValue);
        });
    };

    // Transform the skills from a ; delimited list into a list
    self.loadSkills = function (allSkills) {
        var skills = allSkills.split(";");
        var skillsCount = skills.length;

        if (skillsCount > 0) {
            var skillsCount = skillsCount - 1;
            for (i = 0; i < skillsCount; i++) {
                var technologyAndExperience = skills[i].split(",");
                self.Skills.push({ technology: technologyAndExperience[0], experience: technologyAndExperience[1] });
            }
        }

        // Indicate all data for this form is loaded
        self.isFormLoaded(true);
    };
    
    // Load all attachments of current item
    self.loadAttachments = function (aID) {
        self.AttachmentID(aID);
        var url = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('EmpAttachments')/items/?$select=File/Name,Title&$expand=File&$filter=AttachmentID eq '" + aID + "'";
        $.ajax({
            url: url,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" },
            success: function (data) {
                $.each(data.d.results, function (k, l) {
                    self.addAttachment("EmpAttachments", l.File.Name, l.Title);
                });
            },
            error: function (error) {
                alert(JSON.stringify(error));
            }
        });
    };       
    
    //Load data that's needed to initialize the form (designations)
   self.initNewFormData = function () {
        var siteURL = _spPageContextInfo.webAbsoluteUrl;
        var designationListURL = siteURL + "/_api/web/lists/getbytitle('" + designationListname + "')/items";

        $.ajax({
            url: designationListURL,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" },
            success: function (data) {
                // Add the loaded designations to the dropdown
                $.each(data.d.results, function (k, l) {
                    self.Designations.push(l.Title);
                });
                // indicate the loading of designations is done
                isDesignationsLoaded.resolve();
            },
            error: function (error) {
                alert(JSON.stringify(error));
                isDesignationsLoaded.resolve();
            }
        });

        var countryListURL = siteURL + "/_api/web/lists/getbytitle('" + countryListName + "')/items";

        $.ajax({
            url: countryListURL,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" },
            success: function (data) {
                $.each(data.d.results, function (k, l) {
                    self.Countries.push({ CountryId: l.Id, CountryName: l.Title });
                });
            },
            error: function (error) {
                alert(JSON.stringify(error));
            }
        });

        self.AttachmentID(SP.Guid.newGuid().toString());
    };

    // Load the SharePoint list item that we're editing
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
}

function initEmpForm() {
    //Initialize the OOB SharePoint JSRequest class to read information about the current page
    JSRequest.EnsureSetup();
    var itemId = JSRequest.QueryString["itemId"];

    //Load the view model
    var empModel = new EmpViewModel();
    isDesignationsLoaded = $.Deferred();

    // Load data needed to initialize the view
    empModel.initNewFormData();
    empModel.getSiteGroups();
        
    if (typeof itemId != "undefined") {
        // We're editing an existing item...load the needed data that's specific for editing an item
        empModel.ID(itemId);
        // Once the designations are loaded load the SharePoint list item we're editing
        $.when(isDesignationsLoaded).then(empModel.loadEditFormData());
        empModel.isNewForm(false);
    }
    else {
        // We're adding a new item, load the needed data that's specific for a new item
        empModel.getCurrentUser();
        empModel.addSkill();
        // Indicate the form data is loaded
        empModel.isFormLoaded(true);
    }

    // Bind data to the view
    ko.applyBindings(empModel, document.getElementById('EmpRegistrationForm'));
    
}
$(document).ready(function () {
    initEmpForm();
});
