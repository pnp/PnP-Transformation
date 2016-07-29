# JDP Remediation - CSOM #

### Summary ###
This sample shows an JDP Remediation - CSOM application that is used to perform FTC cleanup post to solution retraction.

### Applies to ###
-  Office 365 Dedicated (D)
-  SharePoint 2013 on-premises


### Solution ###
Solution | Author(s)
---------|----------
JDP Remediation - CSOM | Infosys Ltd, Ron Tielke (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | January 29th 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Introduction #
This is a client-side Console application that leverages the v15 CSOM client SDKs to perform operations against a remote SPO-D 2013 farm.  

The primary purpose of this tool is to allow a customer to remediate issues identified in Setup files, Features and Event Receivers, also provides usage report of Master Pages, Custom Fields and Content Types.  

On executing the Console Application, we get 3 choices of Operations, which are listed as below:
    
<span style="color:red;font-weight:bold">1</span> <span style="color:blue;">Transformation</span>  
<span style="color:red;font-weight:bold">2</span> <span style="color:blue">Clean-Up</span>  
<span style="color:red;font-weight:bold;">3</span> <span style="color:blue;">Self Service Report</span>  
<span style="color:red;font-weight:bold;">4</span> <span style="color:blue;">Exit</span>  

![](\images/\ChoiceOfOperations.png) 

As soon as you run this application, you would need to provide some inputs to execute this console application to get your desired results.

- You will have to provide Admin account details- ***User Name*** and ***Password***
- You will also have to select the corresponding selection of the operation you want to perform.

	<span style="color:blue;"><span style="text-decoration:underline;">**Example**</span>, if you want to execute <span style="text-decoration:underline">`"Clean-Up" operations`</span>, you'll have to enter <span style="color:red;font-weight:bold">2</span> and enter <span style="color:red;font-weight:bold">4</span> to <span style="text-decoration:underline">*quit*</span> the application.</span>


This application logs all exceptions if any occur while execution. This console application is intended to work against SPO-D 2013 (v15) target environments.  

Please find below small summary on each command that has been implemented in this application.

![](\images/JDP.png) 

Please find below more details on these functionalities that has been implemented in this application:  

## <span style="color:blue;">1 - Transformation</span>   ##
On selecting 1st Choice of Operation, we get the following operations as shown in the below screenshot:

![](\images/ChoiceOfOperation1.png) 

These operations are listed and explained as below:

<span style="color:red;font-weight:bold">1</span> <span style="color:green;">Add OOTB Web Part/App Part to a page</span>  
<span style="color:red;font-weight:bold">2</span> <span style="color:green">Replace FTC Web Part with OOTB Web Part/App Part on a page</span>  
<span style="color:red;font-weight:bold;">3</span> <span style="color:green;">Replace MasterPage</span>  
<span style="color:red;font-weight:bold;">4</span> <span style="color:green;">Exit</span> 

### 1. Add OOTB Web Part or App Part to a page ###

This operation adds the web part to the given page present in the given web site.
 
This functionality does not use any input file, but asks user for input of the following parameters: WebUrl, ServerRelative PageUrl, WebPart ZoneIndex, WebPart ZoneID, WebPart FileName, WebPart XmlFile Path. 

Then it generates output log. If at any time the user needs to add web part to any page, using this functionality, it will give the desired output - ***AddWebpart-yyyyMMdd_hhmmss.log*** (verbose log file).

### 2. Replace FTC Web Part with OOTB Web Part or App on a page ###

This operation will read the input file from PreMT-Scan or Discovery output file for Web Parts components (i.e. ***PreMT_MissingWebPart.csv*** or ***WebPartsUsage_Usage.csv*** respectively). It will replace old WebPart (Custom) with new WebPart (OOTB or AppPart) in the given page. First it will delete the existing WebPart present in the page, then will add the new WebPart. 

**Input**

`PreMT_MissingWebPart.csv` file of the Pre-Migration scan. 

A header row is expected with the following format:
*ContentDatabase, FeatureId, FeatureTitle, PageType, PageUrl, SiteCollection, Source, StorageKey, UpgradeStatus, WebApplication, WebPartAssembly, WebPartClass, WebPartId, WebPartTitle, WebPartType, WebUrl, ZoneID, ZoneIndex*

***OR***

`WebPartsUsage_Usage.csv` file of the Discovery scan. 

A header row is expected with the following format:
*ContentDatabase, DirName, Extension, ExtensionForFile, Id, LeafName, ListId, SetupPath, SiteCollection, WebApplication, WebUrl, WFSVC_ListFile,SetupPath, SiteCollection, WebApplication, WebUrl, WFSVC_ListFile*

Also it asks user for input of the following parameters: Input File Path, WebPart Type, Target WebPart File Name, WebPart XmlFilePath.

**Output**

- ReplaceWebPart-yyyyMMdd_hhmmss.log
- PreMT\_MissingWebPart\_ReplaceOperationStatus.csv (if input file provided is PreMT_MissingWebPart.csv)
OR
- WebParts\_Usage\_ReplaceOperationStatus.csv (if input file provided is WebParts_Usage.csv)

### 3. Replace MasterPage ###

This operation reads WebUrls from master page reports generated either from PreMT or Discovery or reads WebUrl from user. And reads custom master page to be replaced either with MasterUrl or CustomMasterUrl or both with Out Of the Box master page

On choosing this option, we would be asked how to proceed for these operations as shown below 

	1)	Process with Input file  
	2)	Process for Web Url


**Input**

- Web Url `(Mandatory for Option 2, not required for other Options)`
- PreMT\_MasterPage\_Usage.csv  `(Mandatory for Option 1, not for others)`
OR  MasterPage_Usage.csv`(Mandatory for Option 1, not for others)`

Also it asks user for input of the following parameters: whether to replace the master Url or custom Url or both, and names of master pages to replace and to be replaced.

**Output**

- ReplaceMasterPage-yyyyMMdd_hhmmss.log 

## <span style="color:Blue;">2 - Clean-Up</span> ##
On selecting 2nd Choice of Operation, we get the following operations as shown in the below screenshot:

![](\images\ChoiceOfOperations2.png) 

These operations are listed and explained as below:

<span style="color:red;font-weight:bold">1</span> <span style="color:green;">Delete Missing Setup Files</span>  
<span style="color:red;font-weight:bold">2</span> <span style="color:green">Delete Missing Features</span>  
<span style="color:red;font-weight:bold;">3</span> 
<span style="color:green;">Delete Missing Event Receivers</span>  
<span style="color:red;font-weight:bold;">4</span> 
<span style="color:green;">Delete Missing Workflow Associations</span>  
<span style="color:red;font-weight:bold;">5</span> <span style="color:green;">Delete All List Template based on Pre-Scan OR Discovery Output OR Output generated by (Self Service > Operation 1)</span>  
<span style="color:red;font-weight:bold;">6</span> <span style="color:green;">Delete Missing Webparts</span>  
<span style="color:red;font-weight:bold;">7</span> <span style="color:green;">Exit</span> 

### 1. Delete Missing Setup Files ###
This operation reads a list of setup file definitions from an input file and deletes the associated setup file from the target SharePoint environment.  

This operation is helpful in trying to remediate the Missing Setup Files reports.  It attempts to remove all specified setup files from the target SharePoint environment.  

On selecting this Operation, we get the following options as shown in the below screenshot:

![](\images\SetupFileCleanUp.png) 

This functionality reads an input file in CSV format, which should contain header columns like - *ContentDatabase, SetupFileDirName, SetupFileExtension, SetupFileName, SetupFilePath, SiteCollection, UpgradeStatus, WebApplication, WebUrl*. Generates result in output file - ***DeleteMissingSetupFiles-yyyyMMdd_hhmmss.log*** (verbose log file).

### 2. Delete Missing Features ###
This operation reads a list of feature definitions from an input file and deletes the associated feature from the webs and sites of the target SharePoint environment.  

This operation is helpful in trying to remediate the Missing Feature reports.  It attempts to remove all specified features from the target SharePoint environment.  

On selecting this Operation, we get the following options as shown in the below screenshot:

![](\images\FeatureCleanUp.png) 

This functionality reads an input file in CSV format, which should contain header columns like - *ContentDatabase, FeatureId, FeatureTitle, SiteCollection, Source, UpgradeStatus, WebApplication, WebUrl*. Generates result in output file - ***DeleteMissingFeatures-yyyyMMdd_hhmmss.log*** (verbose log file).

### 3. Delete Missing Event Receivers ###
This operation reads a list of event receiver definitions from an input file and deletes the associated event receiver from the sites, webs, and lists of the target SharePoint environment.  

This operation is helpful in trying to remediate the Missing Event Receiver reports.  It attempts to remove all specified event receivers from the target SharePoint environment.  

On selecting this Operation, we get the following options as shown in the below screenshot:

![](\images\ERCleanUp.png) 

This functionality reads an input file in CSV format, which should contain header columns like - *Assembly, ContentDatabase, EventName, HostId, HostType, SiteCollection, WebApplication, WebUrl*. Generates result in output file - ***DeleteMissingEventReceivers-yyyyMMdd_hhmmss.log*** (verbose log file).

### 4. Delete Missing Workflow Associations ###

This operation reads a list of workflow association files from an input file and deletes them from the sites, webs, and lists of the target SharePoint environment.

This operation is helpful in trying to remediate the Workflow Associations report of the Pre-Migration Scan.  It attempts to remove all specified files from the target SharePoint environment.

On selecting this Operation, we get the following options as shown in the below screenshot:

![](\images\WFAssociationCleanUp.png) 

**Input**

`PreMT_MissingWorkflowAssociaitons.csv` file of the Pre-Migration scan. A header row is expected with the following format:
*ContentDatabase, DirName, Extension, ExtensionForFile, Id, LeafName, ListId, SetupPath, SiteCollection, WebApplication, WebUrl, WFSVC_ListFile,SetupPath, SiteCollection, WebApplication, WebUrl, WFSVC_ListFile*

**Output**

- DeleteMissingWorkflowAssociations-yyyyMMdd_hhmmss.log

### 5. Delete All List Template based on Pre-Scan OR Discovery Output OR Output generated by (Self Service > Operation 1) ###

This operation reads a list of list templates having customized elements, from an input file generated by **Operation-Generate List Template Report with FTC Analysis**, or from the list templates in gallery reports generated either from PreMT or Discovery, and deletes them from the sites, webs, and lists of the target SharePoint environment.

On selecting this Operation, we get the following options as shown in the below screenshot:

![](\images\ListTemplateCleanUp.png) 

This operation is helpful in trying to remediate the Missing List Templates in Gallery report of the Pre-Migration Scan.  It attempts to remove all specified list templates from the target SharePoint environment.

**Input**

- PreMT_AllListTemplatesInGallery_Usage.csv `(from PreMT Tool)`  
OR
- AllListTemplatesInGallery_Usage.csv `(from Discovery Tool) `  
OR
- ListTemplateCustomization_Usage.csv `(Output from Operation 'Generate List Template Report with FTC Analysis')`  


**Output**

- DeleteListTemplates-yyyyMMdd_hhmmss.log

### 6. Delete Missing WebParts ###
This operation reads a list of webparts from an input file () and deletes the associated type of web parts from the sites, webs, and lists of the target SharePoint environment, which is provided by the user to enter. 

On selecting this Operation, we get the following options as shown in the below screenshot:

![](\images\WebPartCleanUp.png) 


If the type is **all** then all the web parts in the input file are deleted. Whereas, if a specific type is mentioned, then only those web parts types are deleted from the given input file. 

This operation is helpful in trying to remediate the Missing Web Parts reports.  It attempts to remove specified webparts from the target SharePoint environment.  

**Input**

- PreMT_MissingWebParts.csv `(from PreMT Tool)`  
OR
- WebParts_Usage.csv `(from Discovery Tool) `  

**Output**

- DeleteWebpartStatus.csv
- DeleteWebparts-yyyyMMdd_hhmmss.log


## <span style="color:Blue;">3 - Self Service Report</span> ##
On selecting 3rd Choice of Operation, we get the following operations as shown in the below screenshot:

![](\images\ChoiceOfOperation3.png) 

These operations are listed and explained as below:

<span style="color:red;font-weight:bold">1</span> <span style="color:green;">Generate List Tempalte Report with FTC Analysis</span>  
<span style="color:red;font-weight:bold">2</span> <span style="color:green">Generate Site Tempalte Report with FTC Analysis</span>  
<span style="color:red;font-weight:bold;">3</span> 
<span style="color:green;">Generate Site Column and Content Type Usage Report</span>  
<span style="color:red;font-weight:bold;">4</span> 
<span style="color:green;">Generate Non-Default Master Page Usage Report</span>  
<span style="color:red;font-weight:bold;">5</span> <span style="color:green;">Generate Site Collection Report (PPE Only)</span>  
<span style="color:red;font-weight:bold;">6</span> <span style="color:green;">Get Web Part Usage Report</span>  
<span style="color:red;font-weight:bold;">7</span> <span style="color:green;">Get Web Part Properties</span>  
<span style="color:red;font-weight:bold;">8</span> <span style="color:green;">Exit</span> 

### 1. Generate List Template Report with FTC Analysis ###

This operation searches for the Customized elements **(Content Types, Site Columns and Event Receivers)** after extracting the downloaded list templates.
On choosing this option, we would be asked how to proceed for this operation as shown below 

	1)	Process with Auto-generated Site Collection Report  
	2)	Process with PreMT/Discovery ListTemplate Report  
	3)	Process with SiteCollectionUrls separated by comma (,)

This operation is helpful in trying to easily see which List Template has which customized elements. 

**Input**

- Web Application Url `(Mandatory for Option 1, not for other Options)`
- Single or Multiple Site Collection Urls `(Mandatory for Option 3, not for other Options)`
- PreMT_AllListTemplatesInGallery_Usage.csv `(Mandatory for Option 2, not for others)`
OR  AllListTemplatesInGallery_Usage.csv`(Mandatory for Option 2, not for others)`
- CustomFields.csv `(Mandatory for all Options)`
- EventReceivers.csv `(Mandatory for all Options)`
- ContentTypes.csv `(Mandatory for all Options)`

**Output**

- ListTemplateCustomization_Usage.csv
- SiteCollections.txt `(Output for only Option 1)`
- DownloadAndModifyListTemplate-* yyyymmdd*-* hhhhmmss*.log 

> **Note:** If any of the input files *(ContentTypes.csv, CustomFields.csv, EventReceivers.csv)* is not present in the input folder provided by the user, or the file has no entries then corresponding element/s would not be searched to get the customization details in the list templates.

> **Example:** If user has provided only *ContentTypes.csv and CustomFields.csv* in input folder, and *EventReceivers.csv* is not provide in input folder, then  *isCustomEventReceiver* column will have `NO INPUT FILE` value in output report as user has not provided this input file. 

### 2. Generate Site Template Report with FTC Analysis ###

This operation searches for the Customized elements **(Content Types, Site Columns, Features and Event Receivers)** after extracting the downloaded site templates.
On choosing this option, we would be asked how to proceed for this operation as shown below .

	1) Process with Auto-generated Site Collection Report  
	2) Process with PreMT/Discovery SiteTemplate Report  
	3) Process with SiteCollectionUrls separated by comma (,)

This operation is helpful in trying to easily see which Site Template has which customized elements. 


**Input**

- Web Application Url `(Mandatory for Option 1, not for other Options)`
- Single or Multiple Site Collection Urls `(Mandatory for Option 3, not for other Options)`
- PreMT_AllSiteTemplatesInGallery_Usage.csv `(Mandatory for Option 2, not for other Options)`	
OR  AllSiteTemplatesInGallery_Usage.csv `(Mandatory for Option 2, not for other Options)`
- ContentTypes.csv `(Mandatory for all Options)`
- CustomFields.csv `(Mandatory for all Options)`
- EventReceivers.csv `(Mandatory for all Options)`
- Features.csv `(Mandatory for all Options)`


**Output**

- SiteTemplateCustomization_Usage.csv
- SiteCollections.txt `(Output for only Option 1)`
- DownloadAndModifySiteTemplate-yyyymmdd-hhhhmmss.log 


> **Note:** If any of the input files *(Features.csv, ContentTypes.csv, CustomFields.csv, EventReceivers.csv)* is not present in the input folder provided by the user, or the file has no entries then corresponding element/s would not be searched to get the customization details in the site templates.

> **Example:** If user has provided only *Features.csv, ContentTypes.csv and CustomFields.csv* in input folder, and *EventReceivers.csv* is not provide in input folder, then  *isCustomEventReceiver* column will have `NO INPUT FILE` value in output report as user has not provided this input file. 

### 3. Generate Site Column & Content Type Usage Report ###
This operation reads a list of site collection URLs from an input file and scans each site collection, looking for any web or list that is using either a custom Content Type or custom Site Column of interest.  It also looks for local Content Types that have been derived from the custom Content Types of interest.  

This report is helpful in trying to remediate the Missing Content Type and Missing Site Column reports.  This report tells you where within each site collection that the content types and site columns are still in use.  

This functionality reads an input file (Sites.txt, SiteColumns.csv and ContentTypes.csv). Generates result in output file - ***GenerateColumnAndTypeUsageReport-yyyyMMdd_hhmmss.log*** (verbose log file).

### 4. Generate Non-Default Master Page Usage Report ###
This operation reads a list of site collection URLs from an input file and scans each site collection, looking for any web that is using a non-default SP2013 Master Page (i.e., something other than “seattle.master”) as either its System or Site master page.  

This functionality reads an input file (Sites.txt), which contains a fully-qualified, absolute site collection URL. Generates result in output file - ***GenerateNonDefaultMasterPageUsageReport -yyyyMMdd_hhmmss.log*** (verbose log file).

### 5. Generate Site Collection Report (PPE Only) ###
This operation generates a text file containing a list of all site collections found across all web applications in the target farm.  
 
This functionality does not use any input file and generates result in output text format. If at any time the details/report of all the Site Collection of any farm are required, using this functionality will give the desired output - ***GenerateSiteCollectionReport-yyyyMMdd_hhmmss.txt*** (verbose log file).

### 6. Get Web Part Usage Report ###

This operation iterates through all the Pages in root folder of the web, “Pages” and “Site Pages” Library and gives the usage of the given Web Part.
 
This functionality does not use any input file, but asks user for input of the following parameters: WebUrl and WebPart Type 

If at any time the user needs to get the usage of any web part type on any Web Url, using this functionality, it will give the desired output -  ***WebPartUsage.csv*** and the log ***WebPartUsage-yyyyMMdd_hhmmss.log*** (verbose log file).

### 7. Get Web Part Properties ###

This operation will Returns the properties of the given Web Part in Xml format.
 
This functionality does not use any input file, but asks user for input of the following parameters: WebUrl, Server Relative PageUrl, WebPartID 

If at any time the user needs to get the web part properties, using this functionality, it will give the Property Xml file for the corresponding Web Part -  ***WebPartID(provided in input)__WebPartProperties.xml*** and the log ***WebpartProperties-yyyyMMdd_hhmmss.log*** (verbose log file).













