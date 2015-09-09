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

$(document).ready(function () {
    jsSPHostUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));

    if (jsCityVal) {
        loadDropDownListData();
    }

});