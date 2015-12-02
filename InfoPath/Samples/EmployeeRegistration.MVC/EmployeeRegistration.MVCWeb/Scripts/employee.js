var jsSPHostUrl;
var isDDLLoaded;

function getQueryStringParameter(param) {
    var params = document.URL.split("?")[1].split("&");
    var strParams = "";
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == param) {
            return singleParam[1];
        }
    }
}

function addNewSkill() {
    var i = $('#tbodySkills tr').size();
   
    var techControl = "<input type='text' name='Skills[" + i + "].Technology' id='Skills_" + i + "__Technology' class='text-box single-line' text='' />";
    var experienceControl = "<input type='text' name='Skills[" + i + "].Experience' id='Skills_" + i + "__Experience' class='text-box single-line' text='' />";
    var skillRow = "<tr><td>" + techControl + "</td><td>" + experienceControl + "</td></tr>";
    $('#tbodySkills').append(skillRow);
}

function loadDDL(actionName, selectedVal, optionName, ddlID) {
    var url = '/Employee/' + actionName + '/?SPHostUrl=' + jsSPHostUrl;

    $.getJSON(url, { selectedID: selectedVal }, function (data) {
        var items = '<option>Select ' + optionName + '</option>';
        $.each(data, function (i, item) {
            items += "<option value='" + item.Value + "'>" + item.Text + "</option>";
        });
        $(ddlID).html(items);

        if (!(typeof isDDLLoaded === "undefined")) {
            isDDLLoaded.resolve();
        }
    });
}

function LoadCities(stateVal) {
    loadDDL("EmpCities", stateVal, "City", "#Location");
}

$('#CountryID').change(function () {
    var countryVal = $("#CountryID option:selected").text();
    loadDDL("EmpStates", countryVal, "State", "#StateID");
    LoadCities("");
});

$('#StateID').change(function () {
    var stateVal = $("#StateID option:selected").text();
    LoadCities(stateVal);
});

function loadCityData()
{
    isDDLLoaded = $.Deferred();
    var stateVal = $("#StateID option:selected").text();
    LoadCities(stateVal);
     
    $.when(isDDLLoaded).then(function () {
       $("#Location").val(jsCityVal);
    });
}

function loadDropDownListData()
{
    isDDLLoaded = $.Deferred();
    var countryVal = $("#CountryID option:selected").text();
    loadDDL("EmpStates", countryVal, "State", "#StateID");

    $.when(isDDLLoaded).then(function () {
        $("#StateID").val(jsStateID);
        loadCityData();
    });
}

function getProfileData() {
    var url = '/Employee/GetNameAndManagerFromProfile/?SPHostUrl=' + jsSPHostUrl;
   
    $.getJSON(url, { userID: $('#UserID').val() }, function (data) {
        $('#EmpManager').val(data.Manager);
        $('#Name').val(data.Name);
    });
}

function bindUploadedAttachmentsToForm(data) {
    var attachmentURL = jsSPHostUrl + "/Lists/EmpAttachments/" + data.NewFileName;
    var inputAttachment = "<a target='_blank' href=" + attachmentURL + " title=" + data.FileName + ">" + data.FileName + "</a>";
    var removeAttachment = '<a style="cursor: pointer" title="Delete Attachment" onclick="deleteAttachment(this, \'' + data.fileRelativeUrl + '\')">Delete</a>';

    var attachmentRow = "<tr><td>" + inputAttachment + " </td><td>" + removeAttachment + "</td></tr>";
    $('#tbodyAttachments').append(attachmentRow);
    $("#isFileUploaded").val(true);
}

function uploadAttachment() {
    // Disable submit submit button until file gets uploaded to SharePoint library
    var btnSubmit = document.getElementById('btnSubmit');
    btnSubmit.disabled = true;

    var url = '/Employee/UploadAttachment/?SPHostUrl=' + jsSPHostUrl;
    var fileInput = document.getElementById('empAttachment');
    var attachmentID = document.getElementById('AttachmentID').value;

    var formdata = new FormData(); // Send form data to controller
    formdata.append('file', fileInput.files[0]);
    formdata.append('attachmentID', attachmentID);
   
    $.ajax({
        type: "POST",
        url: url,
        dataType: "json",
        data: formdata,
        cache: false,
        contentType: false,
        processData: false,
        success: function (d) {
            bindUploadedAttachmentsToForm(d);
            var file = $("#empAttachment");
            file.replaceWith(file.val('').clone(true));
            btnSubmit.disabled = false; // Enable submit button after file upload
        },
        error: function (xhr, textStatus, errorThrown) {
            alert(textStatus + ":" + errorThrown);
            btnSubmit.disabled = false; // Enable submit button after file upload
        }
    });
}

function deleteAttachment(elem, fileRelativeUrl) {
    // Disable submit submit button until file gets removed from SharePoint library
    var btnSubmit = document.getElementById('btnSubmit');
    btnSubmit.disabled = true;

    var url = '/Employee/DeleteAttachment/?SPHostUrl=' + jsSPHostUrl;
    var formdata = new FormData(); // Send form data to controller
    formdata.append('fileRelativeUrl', fileRelativeUrl);

    $.ajax({
        type: "POST",
        url: url,
        dataType: "json",
        data: formdata,
        cache: false,
        contentType: false,
        processData: false,
        success: function (d) {
            var rowIndex = elem.parentNode.parentNode.rowIndex;
            document.getElementById("tbAttachments").deleteRow(rowIndex);
            btnSubmit.disabled = false; // Enable submit button after file deletion
        },
        error: function (xhr, textStatus, errorThrown) {
            alert(textStatus + ":" + errorThrown + response.responseText);
            btnSubmit.disabled = false; // Enable submit button after file deletion
        }
    });
}

$(document).ready(function () {
    jsSPHostUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));

    if (jsCityVal) {
        loadDropDownListData();
    }

});