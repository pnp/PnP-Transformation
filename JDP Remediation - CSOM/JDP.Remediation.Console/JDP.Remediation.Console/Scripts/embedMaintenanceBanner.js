// choose the location and version of jQuery you wish to use
var JDP_jQueryUrl = "https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.0.2.min.js";

var JDP_EmbedJsFileName = "embedMaintenanceBanner.js";     // this must match the name of this file.

// Specify your banner title here
var JDP_BannerTitle = "Notice: ";
// Specify your maintenance message here
var JDP_BannerMessage = "Migration weekend is coming very soon...";

// Is MDS enabled?
if ("undefined" != typeof g_MinimalDownload && g_MinimalDownload && (window.location.pathname.toLowerCase()).endsWith("/_layouts/15/start.aspx") && "undefined" != typeof asyncDeltaManager) {
    // MDS is enabled; Register script for MDS if possible
    RegisterModuleInit(JDP_EmbedJsFileName, JDP_JavaScriptEmbed);
    JDP_JavaScriptEmbed();
}
else {
    // MDS is not enabled
    JDP_JavaScriptEmbed();
}

// For debugging
function JDP_LogMessage(msg)
{
    if (console) {
        console.log(msg);
    }
}

function JDP_JavaScriptEmbed()
{
    JDP_LoadScript(JDP_jQueryUrl, function () {

        $(document).ready(function () {
            JDP_SetCustomBanner(JDP_BannerTitle, JDP_BannerMessage);
        });
    });
}

function JDP_SetCustomBanner(title, message)
{
    var bannerHtml =
        "<div id='customBanner' class='ms-status-yellow' style='display:block; padding:10px; margin-bottom:15px; border-style:solid; border-width:1px; border-color:#d7d889;'>" +
            "<span class='ms-status-status'>" +
                "<span class='ms-status-iconSpan'><img class='ms-status-iconImg' src='/_layouts/15/images/spcommon.png'/></span>" +
                "<span class='ms-bold ms-status-title'>" + title + "</span>" +
                "<span class='ms-status-body'>" + message + "</span>" +
                "<br/>" +
            "</span>" +
        "</div>";

    if ($('#customBanner').length)
    {
        $('#customBanner').remove();
    }

    // We target a well-known DIV that  unconditionally appears in all pages.
    $('#contentBox').prepend(bannerHtml);
}

function JDP_LoadScript(url, callback)
{
    var head = document.getElementsByTagName("head")[0];
    var script = document.createElement("script");
    script.src = url;

    // Attach handlers for all browsers
    var done = false;
    script.onload = script.onreadystatechange = function () {

        if (!done && (!this.readyState || this.readyState == "loaded" || this.readyState == "complete"))
        {
            done = true;

            // Fire your callback...
            callback();

            // Handle memory leak in IE
            script.onload = script.onreadystatechange = null;
            head.removeChild(script);
        }
    };

    head.appendChild(script);
}
