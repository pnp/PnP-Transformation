<script type="text/javascript">
JSRequest.EnsureSetup();
var displayMode = JSRequest.QueryString["DisplayMode"];

if (typeof displayMode == "undefined") {
    var itemId = JSRequest.QueryString["ID"];
    window.location = _spPageContextInfo.webAbsoluteUrl + "/SitePages/EmpForm.aspx?itemId="+itemId;
}
</script>