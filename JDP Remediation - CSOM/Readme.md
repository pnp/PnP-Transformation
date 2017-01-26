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
    
	1. Transformation  
	2. Clean-Up  
	3. Self-Service Reports 
	4. Exit  

![](\images/ChoiceOfOperations.PNG) 

As soon as you run this application, you would need to provide some inputs to execute this console application to get your desired results.

- You will have to provide Admin account details- ***User Name*** and ***Password***
- You will also have to select the corresponding selection of the operation you want to perform.

	**Example**, if you want to execute `"Clean-Up" operations`, you'll have to enter **2** and enter **4** to `Quit` the application.


This application logs all exceptions if any occur while execution. This console application is intended to work against SPO-D 2013 (v15) target environments.  

Please find below small summary on each command that has been implemented in this application.

![](\images/JDP.PNG) 

Please find below more details on these functionalities that has been implemented in this application:  

## 1 - Transformation ##
On selecting **1st Choice of Operation**, we get the following operations as shown in the below screenshot:

![](\images/ChoiceOfOperation1.PNG) 

These operations are listed and explained as below:

1. Add OOTB Web Part/App Part to a page 
2. Replace FTC Web Part with OOTB Web Part/App Part on a page 
3. Replace Master Page  
4. Exit

### 1. Add OOTB Web Part or App Part to a page ###

This operation adds the web part to the given page present in the given web site.
 
This functionality does not use any input file, but asks user for input of the following parameters: WebUrl, ServerRelative PageUrl, WebPart ZoneIndex, WebPart ZoneID, WebPart FileName, WebPart XmlFile Path. 

Then it generates output log. If at any time the user needs to add web part to any page, using this functionality, it will give the desired output - ***AddWebPart-yyyyMMdd_hhmmss.log*** (verbose log file) and **AddWebPart_SuccessFailure-yyyyMMdd_hhmmss.csv**.

### 2. Replace FTC Web Part with OOTB Web Part or App on a page ###

This operation will read the input file from PreMT-Scan or Discovery output file for Web Parts components (i.e. ***PreMT_MissingWebPart.csv*** or ***WebPartsUsage_Usage.csv*** respectively). It will replace old WebPart (Custom) with new WebPart (OOTB or AppPart) in the given page. First it will delete the existing WebPart present in the page, then will add the new WebPart. 

**Input**

`PreMT_MissingWebPart.csv` file of the Pre-Migration scan. 

A header row is expected with the following format:
*WebPartId, WebPartType, PageUrl, StorageKey, ZoneID, ZoneIndex, WebUrl*

***OR***

`WebPartsUsage_Usage.csv` file of the Discovery scan. 

A header row is expected with the following format:
*WebPartId, WebPartType, PageUrl, StorageKey, ZoneID, ZoneIndex, WebUrl*

Also it asks user for input of the following parameters: Input File Path, WebPart Type, Target WebPart File Name, WebPart XmlFilePath.

**Output**

- ReplaceWebPart-yyyyMMdd_hhmmss.log
- ReplaceWebPart_SuccessFailure-yyyyMMdd_hhmmss.csv 


### 3. Replace MasterPage ###

This operation reads WebUrls from master page reports generated either from PreMT or Discovery or reads WebUrl from user. And reads custom master page to be replaced either with MasterUrl or CustomMasterUrl or both with Out Of the Box master page.

On choosing this option, we would be asked how to proceed for these operations as shown below 

	1)	Process with Input file  
	2)	Process for Web Url


**Input**

- Web Url `(Mandatory for Option 2, not required for other Options)`
- PreMT\_MasterPage\_Usage.csv  `(Mandatory for Option 1, not for others)`
OR  MasterPage_Usage.csv`(Mandatory for Option 1, not for others)`

A header row is expected to have following columns:
*PageUrl, SiteCollection, WebUrl, MasterUrl, CustomMasterUrl*

Also it asks user for input of the following parameters: whether to replace the master Url or custom Url or both, and names of master pages to replace and to be replaced.

**Output**

- ReplaceMasterPage-yyyyMMdd_hhmmss.log 
- ReplaceMasterPage_SuccessFailure- yyyyMMdd_hhmmss.csv

## 2 - Clean-Up ##
On selecting **2nd Choice of Operation**, we get the following operations as shown in the below screen-shot:

![](\images/ChoiceOfOperations2.PNG) 

These operations are listed and explained as below:

1. Delete Missing Setup Files  
2. Delete Missing Features  
3. Delete Missing Event Receivers  
4. Delete Missing Workflow Associations  
5. Delete All List Template based on Pre-Scan OR Discovery Output OR Output generated by (Self Service > Operation 1)  
6. Delete Missing Webparts  
7. Exit 

### 1. Delete Missing Setup Files ###
This operation reads a list of setup file definitions from an input file and deletes the associated setup file from the target SharePoint environment.  

This operation is helpful in trying to remediate the Missing Setup Files reports.  It attempts to remove all specified setup files from the target SharePoint environment.  

**Input**

This functionality reads an input file *PreMT_MissingSetupFile.csv* in CSV format, which should contain header columns like - *ContentDatabase, SetupFileDirName, SetupFileExtension, SetupFileName, SetupFilePath, SiteCollection, UpgradeStatus, WebApplication, WebUrl*.

**Output**

- DeleteSetupFiles-yyyyMMdd_hhmmss.log (verbose log file)
- DeleteSetupFiles_SuccessFailure-yyyyMMdd_hhmmss.csv



### 2. Delete Missing Features ###
This operation reads a list of feature definitions from an input file and deletes the associated feature from the webs and sites of the target SharePoint environment.  

This operation is helpful in trying to remediate the Missing Feature reports.  It attempts to remove all specified features from the target SharePoint environment.  

**Input**

This functionality reads an input file *(PreMT_MissingFeature.csv OR Features_Usage.csv)* in CSV format, which should contain header columns like - *ContentDatabase, FeatureId, FeatureTitle, SiteCollection, Source, UpgradeStatus, WebApplication, WebUrl*

**Output**

- DeleteFeatures-yyyyMMdd_hhmmss.log (verbose log file)
- DeleteFeatures_SuccessFailure-yyyyMMdd_hhmmss.csv



### 3. Delete Missing Event Receivers ###
This operation reads a list of event receiver definitions from an input file and deletes the associated event receiver from the sites, webs, and lists of the target SharePoint environment.  

This operation is helpful in trying to remediate the Missing Event Receiver reports.  It attempts to remove all specified event receivers from the target SharePoint environment.  

**Input**

This functionality reads an input file *(PreMT_MissingEventReceiver.csv OR EventReceivers_Usage.csv)* in CSV format, which should contain header columns like - *Assembly, ContentDatabase, EventName, HostId, HostType, SiteCollection, WebApplication, WebUrl*.

**Output**

- DeleteEventReceivers-yyyyMMdd_hhmmss.log (verbose log file)
- DeleteEventReceivers_SuccessFailure-yyyyMMdd_hhmmss.csv


### 4. Delete Missing Workflow Associations ###

This operation reads a list of workflow association files from an input file and deletes them from the sites, webs, and lists of the target SharePoint environment.

This operation is helpful in trying to remediate the Workflow Associations report of the Pre-Migration Scan.  It attempts to remove all specified files from the target SharePoint environment.

**Input**

*PreMT_MissingWorkflowAssociaitons.csv* file of the Pre-Migration scan.
A header row is expected with the following format:
*ContentDatabase, DirName, Extension, ExtensionForFile, Id, LeafName, ListId, SetupPath, SiteCollection, WebApplication, WebUrl, WFSVC_ListFile,SetupPath, SiteCollection, WebApplication, WebUrl, WFSVC_ListFile*

**Output**

- DeleteWorkflowAssociations-yyyyMMdd_hhmmss.log
- DeleteWorkflowAssociations_SuccessFailure-yyyyMMdd_hhmmss.csv

### 5. Delete All List Template based on Pre-Scan OR Discovery Output OR Output generated by (Self Service > Operation 1) ###

This operation reads a list of list templates having customized elements, from an input file generated by **Operation-Generate List Template Report with FTC Analysis**, or from the list templates in gallery reports generated either from PreMT or Discovery, and deletes them from the sites, webs, and lists of the target SharePoint environment.

This operation is helpful in trying to remediate the All List Templates in Gallery report of the Pre-Migration or Discovery Scan.  It attempts to remove all specified list templates from the target SharePoint environment.

**Input**

- PreMT_AllListTemplatesInGallery_Usage.csv `(from PreMT Tool)`  
OR
- AllListTemplatesInGallery_Usage.csv `(from Discovery Tool) `  
OR
- ListTemplateCustomization_Usage.csv `(Output from Operation 'Generate List Template Report with FTC Analysis')`  


**Output**

- DeleteListTemplates-yyyyMMdd_hhmmss.log
- DeleteListTemplates_SucessFailure-yyyyMMdd_hhmmss.csv

### 6. Delete Missing WebParts ###
This operation reads a list of webparts from an input file (PreMT_MissingWebParts.csv OR WebParts_Usage.csv) and deletes the associated type of web parts from the sites, webs, and lists of the target SharePoint environment, which is provided by the user to enter. 

If the type is **all** then all the web parts in the input file are deleted. Whereas, if a specific type is mentioned, then only those web parts types are deleted from the given input file. 

This operation is helpful in trying to remediate the Missing Web Parts reports.  It attempts to remove specified webparts from the target SharePoint environment.  

**Input**

- PreMT_MissingWebParts.csv `(from PreMT Tool)`  
OR
- WebParts_Usage.csv `(from Discovery Tool) `  

A header row is expected to have following columns:
*PageUrl, WebUrl, StorageKey, WebPartType*

**Output**

- DeleteWebparts_SuccessFailure-yyyyMMdd_hhmmss.csv
- DeleteWebparts-yyyyMMdd_hhmmss.log


## 3 - Self-Service Reports ##
On selecting **3rd Choice of Operation**, we get the following operations as shown in the below screenshot:

![](\images/ChoiceOfOperation3.PNG) 

These operations are listed and explained as below:

1. Generate List Template Report with FTC Analysis  
2. Generate Site Template Report with FTC Analysis  
3. Generate Site Column and Content Type Usage Report  
4. Generate Non-Default Master Page Usage Report  
5. Generate Site Collection Report (PPE Only)  
6. Generate Web Part Usage Report  
7. Generate Web Part Properties Report
8. Generate Security Group Report
8. Exit 

### 1. Generate List Template Report with FTC Analysis ###
This operation downloads List Templates (as directed), extracts the files from each template, and searches the files for instances of the following Customized elements: **(Content Types, Site Columns and Event Receivers)**. Upon choosing this option, the user is prompted for how to proceed as shown below:
	1)	Process with Auto-generated Site Collection Report  
	2)	Process with PreMT/Discovery ListTemplate Report  
	3)	Process with SiteCollectionUrls separated by comma (,)

This operation is helpful in determining the customized elements present in each List Template.

**Input**
This operation reads the following input files:
- Web Application Url `(Mandatory for Option 1, not for other Options)`
- Single or Multiple Site Collection Urls `(Mandatory for Option 3, not for other Options)`
- PreMT_AllListTemplatesInGallery_Usage.csv `(Mandatory for Option 2, not for others)`
OR  AllListTemplatesInGallery_Usage.csv`(Mandatory for Option 2, not for others)`
- CustomFields.csv `(Mandatory for all Options)`
- EventReceivers.csv `(Mandatory for all Options)`
- ContentTypes.csv `(Mandatory for all Options)`

**Output**
This operation generates the following output files:
- ListTemplateCustomization_Usage.csv
- SiteCollections.txt `(Output for only Option 1)`
- DownloadAndModifyListTemplate-*yyyymmdd*-*hhhhmmss*.log 

> **Note:** If any of the input files *(ContentTypes.csv, CustomFields.csv, EventReceivers.csv)* are not present in the specified input folder, or a given file has no entries, the operation will not be able to search the list templates for the corresponding custom elements.

> **Example:** If user has provided only *ContentTypes.csv and CustomFields.csv* in input folder, and *EventReceivers.csv* is not present, the *isCustomEventReceiver* column of the report will have a value of `NO INPUT FILE`. 

### 2. Generate Site Template Report with FTC Analysis ###
This operation downloads Site Templates (as directed), extracts the files from each template, and searches the files for instances of the following Customized elements: **(Content Types, Site Columns, Features and Event Receivers)**. Upon choosing this option, the user is prompted for how to proceed as shown below:
	1) Process with Auto-generated Site Collection Report  
	2) Process with PreMT/Discovery SiteTemplate Report  
	3) Process with SiteCollectionUrls separated by comma (,)

This operation is helpful in determining the customized elements present in each Site Template.

**Input**
This operation reads the following input files, based on the selected option:
- Web Application Url `(Mandatory for Option 1, not for other Options)`
- Single or Multiple Site Collection Urls `(Mandatory for Option 3, not for other Options)`
- PreMT_AllSiteTemplatesInGallery_Usage.csv `(Mandatory for Option 2, not for other Options)` 
OR  
- AllSiteTemplatesInGallery_Usage.csv `(Mandatory for Option 2, not for other Options)`
- ContentTypes.csv `(Mandatory for all Options)`
- CustomFields.csv `(Mandatory for all Options)`
- EventReceivers.csv `(Mandatory for all Options)`
- Features.csv `(Mandatory for all Options)`

**Output**
This operation generates the following output files:
- SiteTemplateCustomization_Usage.csv
- SiteCollections.txt `(Output for only Option 1)`
- DownloadAndModifySiteTemplate-yyyymmdd-hhhhmmss.log 

> **Note:** If any of the input files *(Features.csv, ContentTypes.csv, CustomFields.csv, EventReceivers.csv)* are not present in the specified input folder, or if a given file has no entries, the operation will not be able to search the site templates for the corresponding custom elements.

> **Example:** If user has provided only *Features.csv, ContentTypes.csv and CustomFields.csv* in the input folder, and *EventReceivers.csv* is not present, the *isCustomEventReceiver* column of the report will have a value of `NO INPUT FILE`. 

### 3. Generate Site Column/Custom Fields & Content Type Usage Report ###
This operation reads a list of site collection Urls from an input file and scans each site collection. It reports any web or list that is using either a custom Content Type or custom Site Column of interest.  It also reports any local Content Type that has been derived from a custom Content Type of interest.  

This report is helpful in trying to remediate the Missing Content Type and Missing Site Column reports.  This report tells you where within each site collection that the content types and site columns are still in use.  

**Input**
This operation reads the following input files:
- Sites.txt (no header; one fully-qualified, absolute site collection Url per line)
- CustomFields.csv (header: ID,Name)
- ContentTypes.csv (header: ContentTypeID,ContentTypeName)

**Output**
This operation generates the following output files:
- SiteColumnORFieldAndContentTypeUsage- yyyyMMdd_hhmmss.csv
- GenerateColumnAndTypeUsageReport-yyyyMMdd_hhmmss.log (verbose log file)

### 4. Generate Non-Default Master Page Usage Report ###
This operation reads a list of site collection Urls from an input file and scans each site collection. It reports those webs that are using a non-default SP2013 Master Page (i.e., something other than `“seattle.master”`) as either its System or Site master page.  

> **Note:**
If both Master Page settings (**CustomMasterUrl** and **MasterUrl**) are **“Seattle.master”**, no records are displayed for the web in the output usage file.

> If either Master Page setting (**CustomMasterUrl** or **MasterUrl**) is **“Seattle.master”**, a corresponding record is displayed for the web in the output usage file.

**Input**
This operation reads the following input files:
- Sites.txt (no header; one fully-qualified, absolute site collection Url per line)

**Output**
This operation generates the following output files:
- NonDefaultMasterPageUsage- yyyyMMdd_hhmmss.csv
- GenerateNonDefaultMasterPageUsageReport -yyyyMMdd_hhmmss.log (verbose log file)

### 5. Generate Site Collection Report (PPE Only) ###
This operation reports all site collections found across all web applications in the target farm.  
 
**Input**
This operation does not use any input files; instead, it prompts the user for the following parameters:
- the fully-qualified, absolute URL of an existing site collection in the target farm

**Output**
This operation generates the following output files:
- SiteCollectionReport- yyyyMMdd_hhmmss.txt
- GenerateSiteCollectionReport-yyyyMMdd_hhmmss.log (verbose log file)

### 6. Generate Web Part Usage Report ###
This operation iterates through all Pages present in the *“root folder”*, the *“Pages”* library, and the *“Site Pages”* library of the given web and reports the usage of the given Web Part.
 
**Input**
This functionality does not use any input file; instead, it prompts the user for the following parameters: 
- WebUrl 
- WebPartType

**Output**
This operation generates the following output files:
- WebPartUsage-yyyyMMdd_hhmmss.csv
- WebPartUsage-yyyyMMdd_hhmmss.log (verbose log file)

### 7. Generate Web Part Properties Report ###
This operation reports the properties of the given Web Part in Xml format.
 
**Input**
This functionality does not use any input file; instead, it prompts the user for the following parameters: 
- WebUrl
- ServerRelativePageUrl
- WebPartID

**Output**
This operation generates the following output files:
- WebPartID(provided in input)__WebPartProperties.xml
- WebpartProperties-yyyyMMdd_hhmmss.log (verbose log file)

### 8. Generate Security Group Report ###
This operation reads a list of site collection URLs from an input file and scans each site collection, and reports those that have granted permissions to one or more Security Groups of interest.

**Input**
This operation reads the following input files:
- Sites.txt (no header; one fully-qualified, absolute site collection Url per line)
- SecurityGroups.txt (no header; one security group per line in the following format: <domain>\<groupName>)

**Output**
This operation generates the following output files:
- GenerateSecurityGroupReport-yyyyMMdd_hhmmss.csv
- GenerateSecurityGroupReport-yyyyMMdd_hhmmss.log (verbose log file)














