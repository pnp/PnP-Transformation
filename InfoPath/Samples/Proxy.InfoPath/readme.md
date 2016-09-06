# InfoPath web service proxy #

### Summary ###
This sample shows how to implement a proxy service that re-implements the out-of-the-box SharePoint **GetGroupCollectionsOfUser** ASMX endpoint. In SharePoint Online InfoPath forms [can only call 10 OOB ASMX operations](https://support.microsoft.com/en-us/kb/2674193), all other out-of-the-box ASMX calls will fail. A workaround for this is to host mimic the ASMX service behavior using a custom service. 

### Applies to ###
-  Office 365 Multi Tenant (MT)
-  Office 365 Dedicated (D)
-  SharePoint 2013 on-premises


### Solution ###
Solution | Author(s)
---------|----------
Proxy.InfoPath | Ron Tielke, Bert Jansen (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0 | January 21st 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Introduction #
In SharePoint Online InfoPath Forms Services cannot call into out-of-the-box ASMX services due to the security setup of SharePoint Online. The [10 most frequently called ASMX service operations](https://support.microsoft.com/en-us/kb/2674193) have been "fixed" by a change in InfoPath Forms Services, but what if your form is calling into an ASMX service which is not in the list of 10? Since InfoPath Forms Services can still call external services the solution is to host a custom service externally and call that one. That custom service then can use the SharePoint CSOM to connect back to SharePoint and return the needed information. In order to use this sample one needs to decide on two topics:
- How does the proxy service access SharePoint
- How to verify the service caller 

Next chapters give some more context on these topics.

## Choose a method to access SharePoint ##
The Proxy must be able to access SharePoint to obtain the information (group information in this case). You can choose either of the following approaches, but regardless of your choice, it is best to follow the principal of least privilege.

### Use app-only (preferred) ###
Register and trust an app-only principal via:
- Use AppRegNew to create an App Registration (and App ID)
- Use AppInv to grant AT LEAST the following permission on the target site collection(s) that host the form:

```XML
<AppPermissionRequests AllowAppOnlyPolicy="true">
      <AppPermissionRequest Scope="http://sharepoint/content/sitecollection" Right="Read" />
</AppPermissionRequests>
```

### Use a service account ###
The Service Account needs to be granted AT LEAST the following permissions on the target site collection(s) that host the form:
- Site Permissions:**Browse User Information**  *View information about users of the Web site*
- Site Permissions:**Use Remote Interfaces**  *Use SOAP, Web DAV, the Client Object Model or SharePoint Designer interfaces to access the Web site*
- Site Permissions:**Open**  *Allows users to open a Web site, list, or folder in order to access items inside that container*

You can create a custom permission level with only these permissions and grant that level to the service account or choose for an simpler approach by simply adding the Service Account to the Owners security group.


### Update the samples web.config ###
The appSettings section in web.config from this sample is well documented. Please make the needed changes depending on the chosen setup and target platform

```XML
<appSettings>

    <!-- Set to "true" if the target SharePoint environment is v15 (either on-prem or SPO-D) -->
    <!-- Set to "false" if the target SharePoint environment is v16 (either SPO-MT or SPO-vNext) -->
    <add key="TargetFarmIsSPOD" value="false" />
    
    <!-- Configure the following Client keys if you wish to use App Registration Mode -->
    <!--
    Minimum App Permissions required on target site collection that hosts the InfoPath form:
    <AppPermissionRequests AllowAppOnlyPolicy="true">
      <AppPermissionRequest Scope="http://sharepoint/content/sitecollection" Right="Read" />
    </AppPermissionRequests>

    <add key="ClientId" value="obtain from AppRegNew.aspx" />
    <add key="ClientSecret" value="obtain from AppRegNew.aspx" />
    -->
    <add key="ClientId" value="" />
    <add key="ClientSecret" value="" />
    
    <!-- Configure the following Client keys if you wish to use Service Account Mode -->
    <!--
    minimum permissions required on target site collection that hosts the InfoPath form:
    Site Permissions:Browse User Information  -  View information about users of the Web site.  
    Site Permissions:Use Remote Interfaces    -  Use SOAP, Web DAV, the Client Object Model or SharePoint Designer interfaces to access the Web site. 
    Site Permissions:Open                     -  Allows users to open a Web site, list, or folder in order to access items inside that container.

    *******************************************************************************************************
    NOTE: 
     This Sample simply stores the service account password as cleartext for demo purposes
     Passwords should never be stored as cleartext in configuration files
     For security purposes, your implementation should take steps to encrypt the password as necessary
    *******************************************************************************************************
    
    Format for v15 (TargetFarmIsSPOD=true)
    <add key="ServiceAccountDomain" value="domain" />
    <add key="ServiceAccountUsername" value="username" />
    <add key="ServiceAccountPassword" value="********" />

    Format for v16 (TargetFarmIsSPOD=false)
    <add key="ServiceAccountDomain" value="" />
    <add key="ServiceAccountUsername" value="username@contoso.onMicrosoft.com" />
    <add key="ServiceAccountPassword" value="********" />
    -->
    <add key="ServiceAccountDomain" value="" />
    <add key="ServiceAccountUsername" value="" />
    <add key="ServiceAccountPassword" value="" />

    <!-- 
      The client (i.e., the InfoPath form) must pass this value as the ClientValidationKey in order to be considered a "valid" client. 
      Note: This is a sample value for demo purposes

    <add key="ClientValidationKey" value="ourSharedSecret" />
    -->
    <add key="ClientValidationKey" value="" />

    <!-- 
    Sample values used by the Testing Interface ...
    
    Format for v15 (TargetFarmIsSPOD=true)
    <add key="TestSiteCollectionUrl" value="https://portal.contoso.com" />
    <add key="TestUsername" value="domain\\user" />

    Format for v16 (TargetFarmIsSPOD=false)
    <add key="TestSiteCollectionUrl" value="https://contoso.sharepoint.com/sites/test" />
    <add key="TestUsername" value="user@contoso.onMicrosoft.com" />
    -->
    <add key="TestSiteCollectionUrl" value="" />
    <add key="TestUsername" value="" />

  </appSettings>
```

## Choose a method to verify the caller ##
The proxy must support anonymous access and be internet accessible. It is important that the proxy validate the caller. The SharePoint Add-in Model provides this ability natively when OAuth is used; however, the client in this case (InfoPath Form Services) does not support OAuth. For the proxy sample a shared validation key approach was choosen. The form simply passes the key to the proxy when calling the web method(s). We choose to pass a shared key instead of credentials (e.g., the ServiceAccount/Password or AppID/Secret) in order limit exposure/risk:
- If the key falls into the wrong hands, all the hacker can do is call the Proxy API.  
- If the credentials fall into the wrong hands, the hacker could write an app to do whatever the credentials would allow.

Furthermore, the validation key model allows one to enhance the client validation logic to implement a form-specific, or site-specific, validation keys if necessary.


<img src="https://telemetry.sharepointpnp.com/pnp-transformation/infopath/samples/Proxy.InfoPath" /> 
