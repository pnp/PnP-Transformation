# InfoPath scanning #
The InfoPath transformation process does require you to have a good view on what potentially problematic InfoPath forms live in your farm(s). This section will describe the tools that you can use to acquire this InfoPath information. 

**Important:**
These tools have only been developed and tested against SharePoint 2013.

## How to use
To use the InfoPath scanning components you'll need to follow these steps:

1. Create a folder `c:\scanner\infopath` on a SharePoint server of the farm that you want to scan
2. Compile the **InfoPathScraper** tool and copy the binary into the folder `c:\scanner\infopath\InfoPathScraper`
3. Execute the PowerShell script **01_Get-InfoPathFiles.ps1**: this script will download all the XSN and UDCX files in your farm
4. Execute the PowerShell script **02_Scrape-InfoPathFiles.ps1**: this script will use the **InfoPathScraper** tool to dump relevant information from the XSN files downloaded in the previous step
5. Execute the PowerShell script **03_Get-UdcxReport.ps1**: all UDXC files downloaded in step 3 will be parsed and a CSV report will be generated. This report is needed for UDCX file fixing once these files are migrated over to SharePoint Online
6. Execute the PowerShell script **04_Parse-InfoPathReport.ps1**: the data exported in step 4 (scraper output) is analyzed and all forms which potentially are problematic will be listed in a CSV report
7. **[Optional]** Execute the PowerShell script **05_Get-InfoPathUsageInformation.ps1**: you can use this script to acquire usage Information of the forms which are listed in the previous step. This information can be helpful to guide your customer in determining which forms are still business relevant as you most likely only want to spent effort on those
8. **[Optional]** Execute the PowerShell script **06_Parse-InfoPathPeoplePickerReport.ps1**: this script will generate a report of all forms that have either a people picker of group picker control. This output can be used to perform targeted authentication data fixes as described in [https://github.com/SharePoint/PnP-Transformation/tree/master/InfoPath/Migration/PeoplePickerRemediation.Console](https://github.com/SharePoint/PnP-Transformation/tree/master/InfoPath/Migration/PeoplePickerRemediation.Console "PeoplePickerRemediation.Console") project
9. **[Optional]** Execute the PowerShell script **07_Get-CrossSiteWebServiceReport.ps1**: this will list all UDCX files that that contain cross site collection references. In InfoPath in SharePoint Online you can call the 10 supported ASMX operations but only when the URL to the ASMX endpoint is in the same site collection as where the form is hosted. Using this report you can, before migration, fix potential cross site ASMX calls
10. **[Optional]** Execute the PowerShell script **08_Get-RpcCallReport.ps1**: we've seen calls to OWSSVR.dll to retrieve list data in XML format...calling OWSSVR.dll has been [deprecated in 2014](https://msdn.microsoft.com/en-us/library/office/jj164060.aspx) and is not allowed anymore from InfoPath Forms services in SharePoint Online. Using this report you can identify the forms to be fixed


## Extension to also collect user information
The presented scripts do not pull down user information like site owner or user that last modified the form. If you need that kind of information you can either change the scripts or use a separate script that can read information from the generated CSV files and extend that with user information. Below snippet, which uses [PnP PowerShell](https://github.com/sharepoint/pnp-powershell/releases/latest) will help you get started with that:

```PS
#get use credentials for authentication
$creds = Get-Credential

#site url and file url
$siteUrl = "http://[site-collection-url]"
$fileUrl = "/sites/[site-collection-name]/Shared Documents/testfile.txt"

#Connect to SharePoint 2013 environment (or SPO-D vCurrent)
SharePointPnPPowerShell2013\Connect-PnPOnline –Url $siteUrl –Credentials $creds
$ctx = Get-PnPContext

#get site collection owner information
$site = $ctx.Site
$users = $site.RootWeb.SiteUsers
$ctx.Load($users)
$ctx.ExecuteQuery()
$users.Count
$users | %{if($_.IsSiteAdmin -eq $True) {Write-Host $_.LoginName}} 

#get author and modified by values for specific file
$file = $site.RootWeb.GetFileByServerRelativeUrl($fileUrl);
$ctx.Load($file);
$ctx.Load($file.Author);
$ctx.Load($file.ModifiedBy);
$ctx.ExecuteQuery();
$file.Author | Select-Object LoginName
$file.ModifiedBy | Select-Object LoginName 

```



<img src="https://telemetry.sharepointpnp.com/pnp-transformation/infopath/scanning" /> 
