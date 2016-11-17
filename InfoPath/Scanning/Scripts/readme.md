# PowerShell scripts to detect potentially problematic InfoPath forms #

### Summary ###
This section contains a collection of scripts that can help you detect potentially problematic InfoPath forms

### Applies to ###
-  SharePoint 2013 on-premises
-  SharePoint 2016 on-premises

### Solution ###
Solution | Author(s)
---------|----------
InfoPath scanning scripts | **Microsoft**

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | February 10th 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Scripts in this section
Following scripts are used to obtain a list of potentially problematic InfoPath forms:
- **01_Get-InfoPathFiles.ps1**: this script will download all the XSN and UDCX files in your farm
- **02_Scrape-InfoPathFiles.ps1**: this script will use the **InfoPathScraper** tool to dump relevant information from the XSN files downloaded in the previous step
- **03_Get-UdcxReport.ps1**: all UDXC files downloaded using the 01_Get-InfoPathFiles.ps1 script will be parsed and a CSV report will be generated. This report is needed for UDCX file fixing once these files are migrated over to SharePoint Online
- **04_Parse-InfoPathReport.ps1**: the data exported using the 02_Scrape-InfoPathFiles.ps1 script is analyzed and all forms which potentially are problematic will be listed in a CSV report
- **05_Get-InfoPathUsageInformation.ps1**: you can use this script to acquire usage Information of the forms which are listed in the previous step. This information can be helpful to guide your customer in determining which forms are still business relevant as you most likely only want to spent effort on those
- **06_Parse-InfoPathPeoplePickerReport.ps1**: the data exported using the 02_Scrape-InfoPathFiles.ps1 script is analyzed and all forms which have an or more people picker controls will be listed in a CSV report. You can use this report to determine which InfoPath XML files do require a migration of the stored user/group account information (checkout the PeoplePickerRemediation.Console application in the migration folder)
- **07_Get-CrossSiteWebServiceReport.ps1**: this script uses the data coming out of running *3_Get-UdcxReport.ps1* and *04_Parse-InfoPathReport.ps1* to build a report of forms of which the used UDCX file needs rework. UDCX files pointing to ASMX endpoints which are on a different site collection than the actual site hosting the form will not work in SharePoint Online whereas this does work in SharePoint On-premises   


**Important:**
These scripts have only been tested against SharePoint 2013. 

<img src="https://telemetry.sharepointpnp.com/pnp-transformation/infopath/scanning/scripts" /> 
