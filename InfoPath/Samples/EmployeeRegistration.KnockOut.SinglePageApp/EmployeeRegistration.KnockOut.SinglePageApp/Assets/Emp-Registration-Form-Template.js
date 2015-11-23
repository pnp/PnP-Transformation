<script type="text/javascript" src="~sitecollection/Style%20Library/OfficeDevPnP/jquery-1.10.2.min.js"></script>
<script type="text/javascript" src="~sitecollection/Style%20Library/OfficeDevPnP/knockout-3.3.0.js"></script>
<script type="text/javascript" src="~sitecollection/Style%20Library/OfficeDevPnP/Emp-Registration-Form.js"></script>

<div data-bind="template: {name: 'EmpRegistrationForm-Template'}" id="EmpRegistrationForm">Loading...</div>

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
                    <th style="text-align: right; vertical-align: top;">Attachements</th>
                    <td>
                        <table>
                             <tbody data-bind="foreach: Attachments">
                                <tr>
                                    <td><a target='_blank' data-bind="attr: { href: attachmentURL, title: fileName }, text: fileName"></a></td>
                                    <td><a style="cursor: pointer" title="Delete Attachment" data-bind="event:{click: $root.deleteAttachment.bind($data,fileRelativeUrl)}">Delete</a></td>
                                </tr>
                            </tbody>
                        </table>
                        <table>
                            <tr>
                                <td><input id="empAttachment" type="file"  data-bind="event: {change: changeEmpAttachmentFileInfo}"></td>
                                <td><button data-bind="click: uploadAttachment">Upload</button></td>
                            </tr>
                        </table>                        
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <!-- ko if:isNewForm -->
                        <p><button id="btnSave" data-bind="click: save">Save</button></p>
                        <!-- /ko -->
                        <!-- ko ifnot:isNewForm -->
                        <p><button id="btnUpdate" data-bind="click: update">Update</button></p>
                        <!-- /ko -->
                    </td>
                </tr>
            </table>
        <!-- /ko -->
    <!-- /ko -->
</script>
