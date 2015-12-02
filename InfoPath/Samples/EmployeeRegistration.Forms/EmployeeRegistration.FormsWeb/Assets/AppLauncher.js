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