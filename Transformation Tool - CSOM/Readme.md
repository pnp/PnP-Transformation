# Transformation Tool - CSOM #

### Summary ###
This sample shows an Transformation application that is used to perform replacement of FTC component with equivalent solution.

### Applies to ###
-  Office 365 Dedicated (D)
-  SharePoint 2013 on-premises


### Solution ###
Solution | Author(s)
---------|----------
Transformation Tool - CSOM | Infosys Ltd

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | February 22nd 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Introduction #

**Transformation tool** will help the customers to migrate their existing applications to next version of MS platform providing a solution to their custom components and smooth transition.
This tool is used to replace/transform all the FTC components used in any site (Publish/Team) with OOTB/App Components. Also it'll report where these customization have been replaced (FTC to App).

Customer can transform below customization which are used in farm/Web Application/Site/Web level: 
 
- Master Pages  
- Page/Page Layouts  
- Content Type  
- Custom Fields  
- Custom Field Type  
- Custom List and Libraries  
- Files (CSS/JS/Images etc.)   
- Web Parts
- Web Part Removal 
- Site Column Removal 

**CSOM** approach is used to make Transformation tool farm/server independent.

## 1. Master Page Replacement ##

### <span style="color:brown">Set-MasterPageUsingDiscoveryUsage:</span> ###

#### Functionality:  
- It will read the input csv file and if the old Master Page URL matches with the input Master Page URL then the old Master page is replaced with the New Master Page in that particular web. 
- This functionality reads an input file in CSV format, which should contain header columns like - *CustomMasterUrl,	CustomMasterUrlStatus,	MasterUrl,	MasterUrlStatus,	SiteCollectionOwner, SiteCollection,	WebApplication,	WebUrl*. 
- If we specify old Master Page URL as **“all”**, it will update the master page from all input web/site with `New_MasterPageDetails`.
- Checks whether the Server Relative URL is present for the old Master Page URL that is provided as input. If not it will append the Server Relative URL for the old Master Page URL.


#### Result:
Replace the old master page with new master page and creates a `MasterPage_Replace.CSV` file with the details of the Master Page Replacement

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Set-MasterPageUsingDiscoveryUsage</td>
<td>MasterPage_Usage.CSV</td>
<td>MasterPage_Replace.CSV</td>
</tr>
</table>
 

### <span style="color:brown">Set-MasterPageSiteCollectionLevel:</span> ###

#### Functionality:
- It will iterate through all the sub sites that are present in the site collection.
- Checks whether the Server Relative URL is present for the old Master Page URL that is provided as input. If not it will append the Server Relative URL for the old Master Page URL.
- Checks whether the Server Relative URL is present for the new Master Page URL that is provided as input. If not it will append the Server Relative URL for the new Master Page URL.
- Checks if new master page is available in Gallery or not.
- If `CustomMasterUrlStatus` is `true` and the old Master page URL matches with the Custom Master URL then the Custom Master URL of the web is changed with the new Master Page URL.
- If `MasterUrlStatus` is `true` and the old Master page URL matches with the Master URL then the Master URL of the web is changed with the new Master Page URL.

#### Result:
Replace the old master page with new master page and creates a `MasterPage_Replace.CSV` file with the details of the Master Page Replacement

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Set-MasterPageSiteCollectionLevel</td>
<td>N/A</td>
<td>MasterPage_Replace.CSV</td>
</tr>
</table>



### <span style="color:brown">Set-MasterPageWebLevel:</span> ###

#### Functionality:

- Checks whether the Server Relative URL is present for the old Master Page URL that is provided as input. If not it will append the Server Relative URL for the old Master Page URL.
- Checks whether the Server Relative URL is present for the new Master Page URL that is provided as input. If not it will append the Server Relative URL for the new Master Page URL.
- Checks if new master page is available in Gallery or not.
- If `CustomMasterUrlStatus` is `true` and the old Master page URL matches with the Custom Master URL then the Custom Master URL of the web is changed with the new Master Page URL.
- If `MasterUrlStatus` is `true` and the old Master page URL matches with the Master URL then the Master URL of the web is changed with the new Master Page URL.

#### Result:
Replace the old master page with new master page and creates a `MasterPage_Replace.CSV` file with the details of the Master Page Replacement

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Set-MasterPageWebLevel</td>
<td>N/A</td>
<td>MasterPage_Replace.csv</td>
</tr>
</table>



## 2. Page Layout Replacement ##

### <span style="color:brown">Set-PageLayoutUsingDiscoveryUsage:</span> ###

#### Functionality:  
- It will read the input csv file and if the old Page Layout URL matches with the input Page Layout URL then the old Page Layout is replaced with the New Page.
- This functionality reads an input file in CSV format, which should contain header columns like - *PageLayout_Name, PageLayout_ServerRelativeUrl, SiteCollection, WebApplication, WebUrl*.

#### Result:
Replace the old page layout with new page layout and creates a `PageLayouts_Replace.CSV` file with the details of the Page Layout Replacement.

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Set-PageLayoutUsingDiscoveryUsage</td>
<td>PageLayouts_AvailableCustomPageLayout_Usage.csv</td>
<td>PageLayouts_Replace.csv</td>
</tr>
</table>
 

### <span style="color:brown">Set-PageLayoutSiteCollectionLevel:</span> ###

#### Functionality:  

- It will iterate through all the Page Layouts that are present in the site collection.
- If the old Page Layout URL matches with the input Page Layout URL then the old Page Layout is replaced with the New Page Layout
#### Result:
Replace the old page layout with new page layout and creates a `PageLayouts_Replace.CSV` file with the details of the Page Layout Replacement.

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Set-PageLayoutSiteCollectionLevel</td>
<td>N/A</td>
<td>PageLayouts_Replace.CSV</td>
</tr>
</table>
 

### <span style="color:brown">Set-PageLayoutWebLevel:</span> ###

#### Functionality:  

- It will iterate through all the page layouts that are present in the web.
- If the old Page Layout URL matches with the input Page Layout URL then the old Page Layout is replaced with the New Page Layout
#### Result:
Replace the old page layout with new page layout and creates a `PageLayouts_Replace.CSV` file with the details of the Page Layout Replacement.

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Set-PageLayoutWebLevel</td>
<td>N/A</td>
<td>PageLayouts_Replace.CSV</td>
</tr>
</table>

## 3. Site Columns and Content Types Replacement ##

### <span style="color:brown">Add-SiteColumnToContentTypeUsingCSV:</span> ###

#### Functionality:  
- It will read the input csv file (`ContentType_Usage.csv`) and add the site column to the content type input.
- This functionality reads an input file in CSV format, which should contain header columns like - *ContentTypeId, ContentTypeName, SiteCollection, WebApplication, WebUrl*.

#### Result:
Site column is added to the content type and details are stored in the `SiteColumn_AddTo_ContentType_Replace.csv` file generated.

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Add_siteColumnToContentTypUsingCsv</td>
<td>SiteColumn_addIN_ContentType_Input.csv</td>
<td>SiteColumn_AddTo_ContentType_Replace.Csv</td>
</tr>
</table>
 

### <span style="color:brown">Add-SiteColumnToContentTypeWebLevel:</span> ###

#### Functionality:  
- Add the site column to the content type.

#### Result:
Site column is added to the content type and details are stored in the `SiteColumn_AddTo_ContentType_Replace.csv` file generated.

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Add_siteColumnToContentTypeWebLevel</td>
<td>SiteColumn_addIN_ContentType_Input.csv</td>
<td>SiteColumn_AddTo_ContentType_Replace.csv</td>
</tr>
</table>

### <span style="color:brown">Add-ContentTypeUsingCSV:</span> ###

#### Functionality:  
- It will read the input csv file and create the content type from the old content type. 
- This functionality reads an input file in CSV format, which should contain header columns like - *ContentTypeId, ContentTypeName, SiteCollection, WebApplication, WebUrl*.
- Check if the new content type already exists. 
- If new content type already exists, no further action required. Otherwise, creates a content type information object and load the newly created content type to write in output CSV

#### Result:
Creates the `ContentType_Usage_Replace.csv` file with the details of the content type creation.

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Add_ContentTypeUsingCSV</td>
<td>ContentType_Usage.csv</td>
<td>ContentType_Created.CSV</td>
</tr>
</table>


### <span style="color:brown">Add-ContentTypeWebLevel:</span> ###

#### Functionality:  
- Check if the new content type already exists. If new content type already exists, no further action required. Otherwise, creates a content type information object and load the newly created content type to write in output CSV

#### Result:
Creates the `ContentType_Usage_Replace.csv` file with the details of the content type creation.

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Add_ContentTypeWebLevel</td>
<td>ContentType_Usage.csv</td>
<td>ContentType_Created.CSV</td>
</tr>
</table>

### <span style="color:brown">Add-SiteColumnUsingCSV:</span> ###

#### Functionality:  
- It will read the input csv file and update the site column using CSOM

#### Result:
Creates a `SiteColumn_Usage_Replace.csv` file about site column creation.

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>AddSiteColumnUsingCSV</td>
<td>SiteColumn_CSOM_Input.csv</td>
<td>SiteColumn_Usage_Replace.csv</td>
</tr>
</table>

### <span style="color:brown">Add-SiteColumnWebLevel:</span> ###

#### Functionality:  
- Check if new site column already exists, if it doesn’t exist then get the details of old site column and update schemaXML and create new site column using CSOM.

#### Result:
Create a `SiteColumn_Usage_Replace.csv` file about site column creation and returns a list of site column base class 

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Add-SiteColumnWebLevel</td>
<td>SiteColumn_CSOM_Input.csv</td>
<td>SiteColumn_Usage_Replace.csv</td>
</tr>
</table>


## 4. List Migration ##

### <span style="color:brown">Add-ListMigrateWebLevel:</span> ###

#### Functionality:  
- Migrates list and its contents from the already existing one to new one by using the Base Type of the list as well its base template.
- A new list with exact same content, and provided name is created.

#### Result:
Creates a new list/library by copying the existing content from the original list/library. It also copy all the site columns, content types, attachments and all the files that are associated with the list and creates a `ListMigrationReport.csv` with the details of the List migration.

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Add-ListMigrateWebLevel</td>
<td>N/A</td>
<td>ListMigrationReport.csv</td>
</tr>
</table>
 
### <span style="color:brown">Add-ListMigrateUsingCSV:</span> ###

#### Functionality:  
- Migrates list and its contents from the already existing one to new one by using the Base Type of the list as well its base template.
- A new list with exact same content, and provided name is created.

#### Result:
Creates a new list/library by copying the existing content from the original list/library. It also copy all the site columns, content types, attachments and all the files that are associated with the list and creates a `ListMigrationReport.csv` with the details of the List migration.

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Add-ListMigrateUsingCSV</td>
<td>N/A</td>
<td>ListMigrationReport.csv</td>
</tr>
</table>
 

## 5. Ghosting and Un-Ghosting ##

### <span style="color:brown">Copy-UnGhostFile:</span> ###

#### Functionality:  
- This Command let first downloads the desired file from the document library using the `AbsoluteFilePath` of that file. The new version of this file is then uploaded to the same document library.

#### Result:
The newer version of the downloaded file is uploaded to the same list. The output details are logged in the .CSV file generated- `UnGhostingReport.csv`. 

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Copy-UnGhostFile</td>
<td>N/A</td>
<td>UnGhostingReport.CSV</td>
</tr>
</table>


### <span style="color:brown">Move-UnGhostFile:</span> ###

#### Functionality:  
- This Command let first downloads the desired file from the document library using the `AbsoluteFilePath` of that file. The new version of this file is then uploaded to the same document library and the newer version then overwrites old file..

#### Result:
The newer version of the downloaded file is uploaded to the same document library and the newer version then overwrites old file. The output details are logged in the .CSV file generated- `UnGhostingReport.csv`. 

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Move-UnGhostFile</td>
<td>N/A</td>
<td>UnGhostingReport.CSV</td>
</tr>
</table>


### <span style="color:brown">Get-CustomFile:</span> ###

#### Functionality:  
- This Command let is used to download the desired file from the document library using the `AbsoluteFilePath` of that file. This file is then saved/download in the `OutPutDirectory` specified

#### Result:
The desired file would be downloaded to the `OutPutDirectory` specified using the absolute URL of the File, both of which parameters are provided as input.

#### Output:

<table>
<tr>
<td><b>Command</b></td>
<td><b>Input File</b></td>
<td><b>Output File</b></td>
</tr>
<tr>
<td>Get-File</td>
<td>N/A</td>
<td>DownloadedFilesReport.csv</td>
</tr>
</table>


## 6. WebPart Transformation ##

WebPart Transformation tool is used to replace/transform all the FTC WebParts used in any site (Publish/Team) pages with OOTB WebPart or AppPart.
The flow of this tool is as follows:

- Get Source WebPart Usage (2 Options)
	- Use GetWebPartUsage Command-Let of transformation tool, this has the ZoneID Limitation
	- Use Discovery Tool Usage
- Get Source WebPart Properties
- Configure New/Target WebPart With Source WebPart
- Delete Source WebPart
- Add Target WebPart

Before WebPart Transformation, user can check if the Source WebPart can be configured with the Target WebPart using ConfigureNewWebPart Command-Let.
ConfigureNewWebPart Command-Let will output 2 xml files, Configured and Non-Configured Properties xmls. User can now check for any Non-Configured Properties and make decision accordingly.

### <span style="color:brown">Pre-Requisites:</span> ###
- Source Web Part Type: Any Web part (OOTB/Custom)
- Target Web Part Xml: Either with OOTB Web Part or App Part
- Target Web Part/App Part should be present in the site. 
- We have below Cmdlts for uploading files to the site
- Upload ApptoAppCatalog 
- Upload DependencyFile

### <span style="color:brown">Known Issues/Limitations:</span> ###
- Get-WebPartUsage – Zone ID is missing (Unable to fetch through CSOM)
- Xml Schema Difference – How do we map the properties / xml’s, if SourceWebPart (.dwp) and TargetWebPart (.webpart) or vice-versa sharing the different schema

### <span style="color:brown">Web Part Transformation (Commands-Lets):</span> ###
1.	Get-WebPartUsage
2.	Get-WebPartProperties
3.	Get-WebPartPropertiesByUsageCSV
4.	Move-DependencyFileToServerRelativeFolder
5.	New-WebPartXmlConfiguration
6.	Move-AppToAppCatalog
7.	Add-WebPart
8.	Add-WebPartsByUsageCSV
9.	Remove-WebPart
10.	Remove-WebPartsByUsageCSV
11.	Set-TargetWebPart
12.	Set-TargetWebPartsByUsageCSV
13.	Set-TargetWebPartEnd2EndByUsageCSV


## 7. Web Part Removal ##

Removal of specified Web Part and/or Custom Field used in any SharePoint site using CSOM Based API. These web part and/or custom field can be specify in CSV format file or in cmdlets itself.


### <span style="color:brown">Remove-WebPart:</span> ###

#### Functionality:  
Deletes `WebPart` from Web page, using `StorageKey` property

#### Result:
Deletes `WebPart` from Web page, using `StorageKey` property


### <span style="color:brown">Remove-WebPartsByUsageCSV:</span> ###

#### Functionality:  
Delete `WebPart` from Web page, using `WebPartType` and `Usage CSV file`.

To perform this delete action, make sure that the below columns are there in this input CSV.

- PageUrl
- StorageKey
- WebPartId
- WebPartTitle
- WebPartType
- WebUrl
- ZoneID
- ZoneIndex


#### Result:
Delete `WebPart` from Web page, using `WebPartType` and `Usage CSV file`.

## 8. CSOM Field Removal ##

Removal of specified Web Part and/or Custom Field used in any SharePoint site using CSOM Based API. These web part and/or custom field can be specify in CSV format file or in cmdlets itself.


### <span style="color:brown">Remove-SiteColumnByType:</span> ###

#### Functionality:  
In a given site collection, it removes custom field from SharePoint web [and sub sites if any]. In every web site under given site collection, it scans below items for matching custom field type and remove if any matching field.

- Lists
- List Content Types
- Web Content Type 
- Site Columns

This command-let can be executed in two modes:

**1.	Deletion mode:**
Pass –Confirm parameter to delete custom field which are matching with field type `FieldType` in a given Site Collection `SiteCollectionUrl`

**2.	Read-only mode:**
If you don’t pass –Confirm parameter, it shows usage report of custom field type in console window only.

#### Result:
In a given site collection scans through Lists, List Content Types, Web Content types and site columns of that web and sub webs and removes all reference of custom filed for a specified custom field type


### <span style="color:brown">Remove-SiteColumnByTypeUsingCSV:</span> ###

#### Functionality:  
This command-let takes csv file as input which is having “CustomFieldType” and “SiteCollectionUrl”. This csv file can be output “CustomFieldType_Usage.csv” of Discovery tool or create your own csv file with two columns i.e., “CustomFieldType” and “SiteCollectionUrl”.

This command-let can be executed in two modes:

**1.	Deletion mode:**
Pass –Confirm parameter to delete custom field which are matching with field type `FieldType` in a given Site Collection `SiteCollectionUrl`

**2.	Read-only mode:**
If you don’t pass –Confirm parameter, it shows usage report of custom field type in console window only.

#### Result:
Iterates csv for the given custom field type and retrieves the sites collections, for every site collection scans through Lists, List Content Types, Web Content types and site columns of that web and sub webs and removes all reference of custom filed for a specified custom field type.

----------

<img src="https://telemetry.sharepointpnp.com/pnp-transformation/transformationtool-csom" /> 

