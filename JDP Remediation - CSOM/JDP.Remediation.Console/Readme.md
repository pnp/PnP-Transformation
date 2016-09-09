# JDP.Remediation.Console #

### Summary ###
This sample shows an JDP Remediation - CSOM application that is used to perform FTC cleanup post to solution retraction.

### Applies to ###
-  Office 365 Dedicated (D)
-  SharePoint 2013 on-premises


### Solution ###
Solution | Author(s)
---------|----------
JDP Remediation - CSOM | Ron Tielke (**Microsoft**), Infosys Ltd

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | January 29th 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------


## Introduction ##
This is a client-side Console application that leverages the v15 CSOM client SDKs to perform operations against a remote SPO-D 2013 farm.
The primary purpose of this tool is to allow a customer to remediate issues identified by the various reports of the MSO Pre-Migration Scan.  Feel free to extend the console to include additional operations as needed.


## Scope ##
This console application is intended to work against SPO-D 2013 (v15) target environments.  
It has not been tested against SPO-vNext 2013 (v16) target environments.


## Authentication ##
The console will prompt the user for an administration account.  
Be sure to specify an account that has Admin permissions on the target SharePoint environment.  This account will be used to generate Authenticated User Context instances that will be leveraged to access the target environments.  

- If you wish to target an **SPO-D (or On-Prem)** farm:
  *  use the **<domain\>\<alias\>** format for the administrator account.
- If you wish to target an **SPO-MT (or vNext)** farm:
  *  use the **<alias\>@<domain\>.com** format for the administrator account.


## Performance ##
The utility will perform its best when using a client machine that has 8GB RAM and a reliable internet connection with ample bandwidth.  

The utility is currently implemented as a single-threaded application.  As such, it will process the specified input file in a serial fashion, one line at a time.  It can take several hours to process large input files (10,000 rows or more). 
 
In these cases, the easiest and fastest way to improve performance is to partition the input file into smaller files and use multiple instances of the utility (running either on the same machine, or on multiple machines) to process each partition.


## Commands ##

On executing the Console Application, we get 3 choices of Operations, which are listed as below:

<span style="color:red;font-weight:bold">1</span> <span style="color:blue;">Transformation</span>  
<span style="color:red;font-weight:bold">2</span> <span style="color:blue">Clean-Up</span>  
<span style="color:red;font-weight:bold;">3</span> <span style="color:blue;">Self Service Report</span>  
<span style="color:red;font-weight:bold;">4</span> <span style="color:blue;">Exit</span> 

The same are shown in the Screen shot below:

![](\images/ChoiceOfOperations.PNG)

## <span style="color:blue;">1 - Transformation</span> ##

On selecting 1st Choice of Operation, we get the following operations as shown in the below screenshot:

![](\images/ChoiceOfOperation1.PNG) 

These operations are listed and explained as below:

### 1. Add OOTB Web Part or App Part to a page ###

This operation adds the web part to the given page present in the given web site.  
**Note**: Web Part should be present in the Web Part Gallery.
 
**Input**

On selecting this operation, we will be asked for no of parameters as explained below:

![](\images/\AddOOTBtoPage.PNG) 

- **Web Url**
	* Here we need to provide the Web Url where we need to add the WebPart  
	**Example:**    “https://intranet.campoc.com/sites/OffshorePoc"
- **Server Relative PageUrl:**
	* Here we need to provide the server relative url of the page in Web Url where we need to add the WebPart  
	**Example:**    “/sites/OffshorePoc/SitePages/DevHome.aspx”
- **WebPart ZoneIndex:**
	* Here we need to provide the zone index of the web part we need to add  
	**Example:**    “0"
- **WebPart ZoneID:**
	* Here we need to provide the zone ID of the web part we need to add  
	**Example:**    “LeftColumnZone"
- **WebPart File Name**
	* Here we need to provide the WebPart Name which we need to add.  
	**Example:** webPartName.webpart

**Note:** This webpart should be present in the WebPart Gallery of the given WebUrl

- **WebPart XmlFile Path:**
	* Here we need to provide the zone ID of the web part we need to add  
	**Example:** “E:\ProjectTest\Configured_5a170911-91dd-4643-adb6-289565f12867_TargetContentQuery.xml”


**Output**

- **AddWebPart-yyyyMMdd_hhmmss.log**
  * This is the verbose log file of the scan.
  * Success messages of interest:
      * SUCCESS: Added File: {0}
  * Informational messages of interest:
      * none
  * Error messages of interest:
      * Error=File Not Found
          * Cause: The specified file or folder does not exist
          * Remediation: none; the file is gone
      * Error=The file is checked out for editing
          * Cause: someone has checked out the file for editing
          * Remediation: 
              * Visit the site containing the locked file
              * Undo the check-out
              * Delete the locked file
      * (404) Not Found
          * Cause: : The specified site collection does not exist
          * Remediation: none; the site collection does not exist
      * Cannot contact site at the specified URL
          * Cause: The specified web (subweb, subsite, etc.) does not exist
          * Remediation: none; the web does not exist
- **AddWebPart\_SuccessFailure-yyyyMMdd_hhmmss.csv**
	* In this report, details of newly added web part will be present.
	* This output file would contain the below mentioned columns and the entries corresponding to webpart.
		*	**PageUrl:** Specifies the Url of the page where the web part is added.
		*	**SiteCollection:** Specifies the Site Collection where in the web part is present.
		*	**Status:** Specifies whether the status of adding of web part. It contains one of the following values
			*	**Success:** If value of this column is “Success” it implies that the adding of the web part was successful
			*	**Failure:** If value of this column is “Failure” it implies that the adding of the web part was not successful due to some error. 
		*	**WebApplication:** Specifies the Web Application of the Site Collection where in the web part is present.
		*	**WebPartFileName:** Specifies the web part name
		*	**WebUrl:** Specifies the Web Url for the Site Collection where in the web part is present
		*	**ZoneID:** Specifies the zone ID of the web part
		*	**ZoneIndex:** Specifies the Zone Index of the web part


### 2. Replace FTC Web Part with OOTB Web Part or App on a page ###
This operation will read the input file from PreMT-Scan or Discovery output file for Web Parts components (i.e. PreMT_MissingWebPart.csv or WebPartsUsage_Usage.csv respectively). It will replace old WebPart (Custom) with new WebPart (OOTB or AppPart) in the given page. First it will delete the existing WebPart present in the page, then will add the new WebPart.  

On selecting this Operation, we get the following options as shown in the below screenshot:

![](\images/\ReplaceFTCWebpart.PNG) 

Before running this function, please run GetWebPartProperties function, so that the XMLs for both source and target WebParts is available, to run delete and add function.  

**Note:** Web Part should be present in the Web Part Gallery.
 

**Input**

- **`PreMT_MissingWebPart.csv`** file of the Pre-Migration scan. 

 * A header row is expected with the following format:  
		* *WebPartId, WebPartType, PageUrl, StorageKey, ZoneID, ZoneIndex, WebUrl*

***OR***

- **`WebPartsUsage_Usage.csv`** file of the Discovery scan. 

	* A header row is expected with the following format:
		* *WebPartId, WebPartType, PageUrl, StorageKey, ZoneID, ZoneIndex, WebUrl*

Also it asks user for input of the following parameters:

- **Input File Path:**
	* Here we need to the file path of the above mentioned input file  
	**Example:**    “E:\WebParts_Usage.csv"
- **WebPart Type:**
	* Here we need to provide the WebPart Type which we need to be replaced  
	**Example:**    “ContentEditorWebPart”
- **Target WebPart File Name:**
	* Here we need to provide the WebPart Name which we need to replace  
	**Example:**    “ContentQuery.webpart"
- **WebPart ZoneID:**
	* Here we need to provide the zone ID of the web part we need to add  
	**Example:**    “LeftColumnZone"
- **WebPart XmlFile Path:**
	* Here we need to provide the Path of the Xml file corresponding to Web Part  
	**Example:** “E:\ProjectTest\Configured_5a170911-91dd-4643-adb6-289565f12867_TargetContentQuery.xml”

**Output**

-	**ReplaceWebPart\_SuccessFailure-yyyyMMdd_hhmmss.csv** 
  * In this report, all the web parts replaced will be present.
  * This output file would contain the below mentioned columns and the entries corresponding to webaprts.
      * **PageUrl:** Specifies the Url of the page where the web part is replaced
      * **WebApplication:** Specifies the Web Application of the Site Collection where in the web part is present.
      * **SiteCollection:** Specifies the Site Collection where in the web part is present.
      * **WebUrl:** Specifies the Web Url for the Site Collection where in the web part is present.
      * **Status:** Specifies whether the status of replacement of web part. It contains one of the following values
	      * Success: If value of this column is “Success” it implies that the replacement of the web part was successful
	      * Error: If value of this column is “Error” it implies that the replacement of the web part was not successful due to some error. Same error details will be mentioned in the log file as mentioned below
		* **WebPartId:** Specifies Web Part Id of the web part
		* **WebPartType:** Specifies the Web Part Type of the web part
		* **ZoneID:** Specifies the zone ID of the web part
		* **ZoneIndex:** Specifies the Zone Index of the web part  
- **ReplaceWebPart-yyyyMMdd_hhmmss.log**
	* This is the verbose log file of the scan.
  * Success messages of interest:
      * SUCCESS: Added File: {0}
  * Informational messages of interest:
      * none
  * Error messages of interest:
      * Error=File Not Found
          * Cause: The specified file or folder does not exist
          * Remediation: none; the file is gone
      * Error=The file is checked out for editing
          * Cause: someone has checked out the file for editing
          * Remediation: 
              * Visit the site containing the locked file
              * Undo the check-out
              * Delete the locked file
      * (404) Not Found
          * Cause: : The specified site collection does not exist
          * Remediation: none; the site collection does not exist
      * Cannot contact site at the specified URL
          * Cause: The specified web (subweb, subsite, etc.) does not exist
          * Remediation: none; the web does not exist


### 3. Replace MasterPage ###

This operation reads WebUrls from master page reports generated either from PreMT or Discovery or reads WebUrl from user. And reads custom master page to be replaced either with MasterUrl or CustomMasterUrl or both with Out Of the Box master page

On choosing this option, we would be asked how to proceed for this operation as shown below 

![](\images/\ReplaceMasterPage.PNG)

There are two approaches as shown in the figure:
  
*Please select any of following options:*

	1)	Process with Input file  
	2)	Process for Web Url

On selecting above option, console prompts to select another options as described below,

	1)	Replace MasterUrl
		Select this option if wanted to replace Master Url or System Master Page
	2)	Replace Custom MasterUrl
		Select this option if wanted to replace Custom Master Url or Site Master Page
	3)	Replace Both Master Urls   
		Select this option if wanted to replace both Master Url & Custom Master Url 

Also it asks user for input of the following parameters: 

- **Custom master page:**
	* Here we need to the file path of the above mentioned input file  
		**Example:**    “contoso.master"
- **OOTB master page:**
	* Here we need to provide the WebPart Type which we need to be replaced  
	**Example:**    “seattle.master”

**Input**

- **Web Url** `(Mandatory for Option 2, not required for other Options)`
- **PreMT\_MasterPage\_Usage.csv**  `(Mandatory for Option 1, not for others)`
	* This is a CSV that follows the format and content of the PreMT\_MasterPage\_Usage.csv file of the Pre-Migration scan. A header row is expected to have following columns:
		* PageUrl, SiteCollection, WebUrl, MasterUrl, CustomMasterUrl

		**OR**  

- **MasterPage_Usage.csv** `(Mandatory for Option 1, not for others)`
	* This is a CSV that follows the format and content of the MasterPage_Usage.csv file of the Discovery. A header row is expected to have following columns:
		* PageUrl, SiteCollection, WebUrl, MasterUrl, CustomMasterUrl
		
**Output**

- **ReplaceMasterPage-yyyyMMdd_hhmmss.log** 
	* This is the verbose log file of the scan.
  * Success messages of interest:
      * SUCCESS: Added File: {0}
  * Informational messages of interest:
      * none
  * Error messages of interest:
      * Error=File Not Found
          * Cause: The specified file or folder does not exist
          * Remediation: none; the file is gone
      * Error=The file is checked out for editing
          * Cause: someone has checked out the file for editing
          * Remediation: 
              * Visit the site containing the locked file
              * Undo the check-out
              * Delete the locked file
      * (404) Not Found
          * Cause: : The specified site collection does not exist
          * Remediation: none; the site collection does not exist
      * Cannot contact site at the specified URL
          * Cause: The specified web (subweb, subsite, etc.) does not exist
          * Remediation: none; the web does not exist.
- **ReplaceMasterPage\_SuccessFailure-yyyyMMdd_hhmmss.csv** 
	* In this report, all the master pages replaced will be present.
	* This output file would contain the below mentioned columns and the entries corresponding to master page.
		* **CustomMasterPageUrl:** Specifies the Url of the custom master page replaced.
		* **OOTBMasterPageUrl:** Specifies the Url of the OOTB master page 
		* **WebApplication:** Specifies the Web Application of the Site Collection where in the master page is present.
		* **SiteCollection:** Specifies the Site Collection where in the master page is present.
		* **WebUrl:** Specifies the Web Url for the Site Collection where in the master page is present.
		* **Status:** Specifies whether the status of replacement of master page. It contains one of the following values
			* **Success:** If value of this column is “Success” it implies that the replacement of the master page was successful
			* **Failure:** If value of this column is “Failure” it implies that the replacement of the master page was not successful due to some error. 


## <span style="color:Blue;">2 - Clean-Up</span> ##

On selecting 2nd Choice of Operation, we get the following operations as shown in the below screenshot:

![](\images\ChoiceOfOperations2.png) 

These operations are listed and explained as below:

### 1. Delete Missing Setup Files ###
This operation reads a list of setup file definitions from an input file and                <span style="color:red;">deletes</span> the associated setup file from the target SharePoint environment.  


On selecting this Operation, we get the following options as shown in the below screenshot:

![](\images/\SetupFileCleanUp.PNG) 

This operation is helpful in trying to remediate the Missing Setup Files report of the Pre-Migration Scan.  It attempts to remove all specified setup files from the target SharePoint environment.

#### Input ####
- **PreMT_MissingSetupFile.csv**
  * This is a CSV that follows the format and content of the PreMT_MissingSetupFile.csv file of the Pre-Migration scan. A header row is expected with the following format:
      * ContentDatabase, SetupFileDirName, SetupFileExtension, SetupFileName, SetupFilePath, SiteCollection, UpgradeStatus, WebApplication, WebUrl

#### Output ####
- **DeleteSetupFiles-yyyyMMdd_hhmmss.log**
  * This is the verbose log file of the scan.
  * Success messages of interest:
      * SUCCESS: Deleted File: {0}
  * Informational messages of interest:
      * none
  * Error messages of interest:
      * Error=File Not Found
          * Cause: The specified file or folder does not exist
          * Remediation: none; the file is gone
      * Error=Cannot remove file
          * Cause: The file is likely being used as the default Master Page of the site
          * Remediation: 
              * Use SPD to open the site containing the locked file
              * Configure **seattle.master** to be both default MPs
              * Delete the locked file
      * Error=This item cannot be deleted because it is still referenced by other pages
          * Cause: the file is being used by other pages
          * Remediation: 
              * Visit the site containing the locked file
              * Go to Site Settings and click Manage Content and Structure
                  * Or hack the URL: /_layouts/15/**siteManager**.aspx
              * Generate a References Report for the locked file
              * Remediate all references
              * Delete the locked file
      * Error=The file is checked out for editing
          * Cause: someone has checked out the file for editing
          * Remediation: 
              * Visit the site containing the locked file
              * Undo the check-out
              * Delete the locked file
      * (404) Not Found
          * Cause: : The specified site collection does not exist
          * Remediation: none; the site collection does not exist
      * Cannot contact site at the specified URL
          * Cause: The specified web (subweb, subsite, etc.) does not exist
          * Remediation: none; the web does not exist
- **DeleteSetupFiles\_SuccessFailure-yyyyMMdd_hhmmss.csv**
	* In this report, details of the deleted setup files will be present.
	* This output file would contain the below mentioned columns and the entries corresponding to deleted setup file.
		* **SetupFileDirName:** Specifies the directory of the setup file.
		* **SetupFileName:** Specifies the name of the setup file.
		* **WebApplication:** Specifies the Web Application of the Site Collection where in the setup file is present.
		* **WebUrl:** Specifies the Web Url for the Site Collection where in the setup file is present.
		* **Status:** Specifies whether the status of deletion of setup file. It contains one of the following values
			* **Success:** If value of this column is “Success” it implies that the deletion of the setup file was successful
			* **Failure:** If value of this column is “Failure” it implies that the deletion of the setup file was not successful due to some error. 


### 2. Delete Missing Features ###
This operation reads a list of feature definitions from an input file and deletes the associated feature from the webs and sites of the target SharePoint environment.  


On selecting this Operation, we get the following options as shown in the below screenshot:

![](\images/\FeatureCleanUp.PNG) 

This operation is helpful in trying to remediate the Missing Feature report of the Pre-Migration Scan.  It attempts to remove all specified features from the target SharePoint environment.


#### Input ####
- **PreMT_MissingFeature.csv**
	* This is a CSV that follows the format and content of the PreMT_MissingFeature.csv file of the Pre-Migration scan. A header row is expected with the following format:
    	* FeatureId, Scope, ContentDatabase, WebApplication, SiteCollection, WebUrl

OR

- **Features_Usage.csv**
	* This is a CSV that follows the format and content of the Features_Usage.csv file of the Discovery scan. A header row is expected with the following format:
    	* FeatureId, Scope, ContentDatabase, WebApplication, SiteCollection, WebUrl
    	

#### Output ####
- **DeleteFeatures-yyyyMMdd_hhmmss.log**
  * This is the verbose log file of the scan.
  * Success messages of interest:
      * SUCCESS: Deleted Feature {0} from web {1}
      * SUCCESS: Deleted Feature {0} from site {1}
  * Informational messages of interest:
      * WARNING: feature was not found in the web-scoped features; trying the site-scoped features...
      * WARNING: Could not delete Feature {0}; feature not found
      * WARNING: Could not delete Feature {0}; feature not found in site.Features
      * WARNING: Could not delete Feature {0}; feature not found in web.Features
  * Error messages of interest:
      * (404) Not Found
          * Cause: : The specified site collection does not exist
          * Remediation: none; the site collection does not exist
      * Cannot contact site at the specified URL
          * Cause: The specified web (subweb, subsite, etc.) does not exist
          * Remediation: none; the web does not exist
- **DeleteFeatures\_SuccessFailure-yyyyMMdd_hhmmss.csv**
	* In this report, details of the deleted features will be present.
	* This output file would contain the below mentioned columns and the entries corresponding to deleted features.
		* **FeatuerId:** Specifies the Id of the feature.
		* **Scope:** Specifies the scope of the feature.
		* **SiteCollection:** Specifies the Site Collection where in the feature is present.
		* **WebApplication:** Specifies the Web Application of the Site Collection where in the feature is present.
		* **WebUrl:** Specifies the Web Url for the Site Collection where in the feature is present.
		* **Status:** Specifies whether the status of deletion of feature. It contains one of the following values
			* **Success:** If value of this column is “Success” it implies that the deletion of the feature was successful
			* **Failure:** If value of this column is “Failure” it implies that the deletion of the feature was not successful due to some error. 


### 3. Delete Missing Event Receivers ###
This operation reads a list of event receiver definitions from an input file and deletes the associated event receiver from the sites, webs, and lists of the target SharePoint environment. 

On selecting this Operation, we get the following options as shown in the below screenshot:

![](\images/\ERCleanUp.PNG) 

This operation is helpful in trying to remediate the Missing Event Receiver report of the Pre-Migration Scan.  It attempts to remove all specified event receivers from the target SharePoint environment.


#### Input ####
- **PreMT_MissingEventReceiver.csv**
  * This is a CSV that follows the format and content of the **PreMT_MissingEventReceiver.csv** file of the Pre-Migration scan. A header row is expected with the following format:
      * Assembly, ContentDatabase, HostId, HostType, SiteCollection, WebApplication, WebUrl

OR

- **EventReceivers_Usage.csv**
  * This is a CSV that follows the format and content of the **EventReceivers_Usage.csv** file of the Discovery scan. A header row is expected with the following format:
      * Assembly, ContentDatabase, HostId, HostType, SiteCollection, WebApplication, WebUrl
      
#### Output ####
- **DeleteEventReceivers-yyyyMMdd_hhmmss.log**
  * This is the verbose log file of the scan.
  * Search the log for instances of the following significant entries:
      * SUCCESS: Deleted SITE Event Receiver [{0}] from site {1}
      * SUCCESS: Deleted WEB Event Receiver [{0}] from web {1}
      * SUCCESS: Deleted LIST Event Receiver [{0}] from list [{1}] on web {2}
  * Informational messages of interest:
      * <span style="background-color: #FFFF00">TBD</span>
  * Error messages of interest:
      * (404) Not Found
          * Cause: : The specified site collection does not exist
          * Remediation: none; the site collection does not exist
      * Cannot contact site at the specified URL
          * Cause: The specified web (subweb, subsite, etc.) does not exist
          * Remediation: none; the web does not exist
- **DeleteEventReceivers\_SuccessFailure-yyyyMMdd_hhmmss.csv**
	* In this report, details of the deleted event receivers will be present.
	* This output file would contain the below mentioned columns and the entries corresponding to deleted event receiver.
		* **Assembly:** Specifies the Assembly of the event receiver.
		* **EventName:** Specifies the name of the event receiver.
		* **HostId:** Specifies the HostId of the event receiver.
		* **HostType:** Specifies the HostType of the event receiver.
		* **SiteCollection:** Specifies the Site Collection where in the event receiver is present.
		* **WebApplication:** Specifies the Web Application of the Site Collection where in the event receiver is present.
		* **WebUrl:** Specifies the Web Url for the Site Collection where in the event receiver is present.
		* **Status:** Specifies whether the status of deletion of event receiver. It contains one of the following values
			* **Success:** If value of this column is “Success” it implies that the deletion of the event receiver was successful
			* **Failure:** If value of this column is “Failure” it implies that the deletion of the event receiver was not successful due to some error. 

### 4. Delete Missing Workflow Associations ###
This operation reads a list of workflow association files from an input file and deletes them from the sites, webs, and lists of the target SharePoint environment.


On selecting this Operation, we get the following options as shown in the below screenshot:

![](\images/\WFAssociationCleanUp.PNG) 

This operation is helpful in trying to remediate the Workflow Associations report of the Pre-Migration Scan.  It attempts to remove all specified files from the target SharePoint environment.


#### Input ####
- **PreMT_MissingWorkflowAssociaitons.csv**
  * This is a CSV that follows the format and content of the **PreMT_MissingWorkflowAssociaitons.csv** file of the Pre-Migration scan. A header row is expected with the following format:
      * ContentDatabase, DirName, Extension, ExtensionForFile, Id, LeafName, ListId, SetupPath, SiteCollection, WebApplication, WebUrl, WFSVC_ListFile, SetupPath, SiteCollection, WebApplication, WebUrl, WFSVC_ListFile

#### Output ####
- **DeleteWorkflowAssociations-yyyyMMdd_hhmmss.log**
	* This is the verbose log file of the scan.  
	* Success messages of interest:
		* SUCCESS: Deleted File: {0}
	* Informational messages of interest:
		* none
	* Error messages of interest:
		* Error=File Not Found
			* Cause: The specified file or folder does not exist
			* Remediation: none; the file is gone
		* Error=The file is checked out for editing
			* Cause: someone has checked out the file for editing
			* Remediation:
				* Visit the site containing the locked file
				* Undo the check-out
				* Delete the locked file
		* (404) Not Found
			* Cause: The specified site collection does not exist
			* Remediation: none; the site collection does not exist
		* Cannot contact site at the specified URL
			* Cause: The specified web (subweb, subsite, etc.) does not exist
			* Remediation: none; the web does not exist
- **DeleteWorkflowAssociations\_SuccessFailure-yyyyMMdd_hhmmss.csv**
	* In this report, details of the deleted workflow associations will be present.
	* This output file would contain the below mentioned columns and the entries corresponding to deleted workflow associations.
		* **DirName:** Specifies the directory name of the workflow associations.
		* **LeafName:** Specifies the leaf name of the workflow associations.
		* **SiteCollection:** Specifies the Site Collection where in the workflow associations is present.
		* **WebApplication:** Specifies the Web Application of the Site Collection where in the workflow associations is present.
		* **WebUrl:** Specifies the Web Url for the Site Collection where in the workflow associations is present.
		* **Status:** Specifies whether the status of deletion of workflow associations. It contains one of the following values
			* **Success:** If value of this column is “Success” it implies that the deletion of the workflow associations was successful
			* **Failure:** If value of this column is “Failure” it implies that the deletion of the workflow associations was not successful due to some error. 

		
### 5. Delete All List Template based on Pre-Scan OR Discovery Output OR Output generated by (Self Service > Operation 1) ###
This operation reads a list of list templates having customized elements, from an input file generated by **Operation-Generate List Template Report with FTC Analysis**, or from the list templates in gallery reports generated either from PreMT or Discovery, and deletes them from the sites, webs, and lists of the target SharePoint environment.

On selecting this Operation, we get the following options as shown in the below screenshot:

![](\images/\ListTemplateCleanUp.PNG) 

This operation is helpful in trying to remediate the Missing List Templates in Gallery report of the Pre-Migration Scan.  It attempts to remove all specified list templates from the target SharePoint environment.
 

#### Input ####

- PreMT\_AllListTemplatesInGallery_Usage.csv `(from PreMT Tool)`  
OR
- AllListTemplatesInGallery_Usage.csv `(from Discovery Tool) `  
OR
- ListTemplateCustomization_Usage.csv `(Output from Operation-Generate List Template Report with FTC Analysis)`  

#### Output ####

- **DeleteListTemplates-yyyymmdd-hhhhmmss.log**
	* This is the verbose log file of the scan.  
	* Success messages of interest:
		* SUCCESS: Deleted File: {0}
	* Informational messages of interest:
		* none
	* Error messages of interest:
		* Error=File Not Found
			* Cause: The specified file or folder does not exist
			* Remediation: none; the file is gone
		* Error=The file is checked out for editing
			* Cause: someone has checked out the file for editing
			* Remediation:
				* Visit the site containing the locked file
				* Undo the check-out
				* Delete the locked file
		* (404) Not Found
			* Cause: The specified site collection does not exist
			* Remediation: none; the site collection does not exist
		* Cannot contact site at the specified URL
			* Cause: The specified web (subweb, subsite, etc.) does not exist
			* Remediation: none; the web does not exist

### 6. Delete Missing WebParts ###
This operation reads PageUrls from web part report generated either from PreMT or Discovery and deletes webparts based on Webpart type. User will be prompted to enter Webpart type to delete which kind of webparts to delete form input report file or he can enter ‘all’ to delete all webparts in input report file. 

On choosing this option, we would be asked how to proceed for this operation as shown in the below image:

![](\images/\WebPartCleanUp.PNG)

- **WebPart Type**
	* Here we need to the give the Web Part Type that we want to delete  
	**Example:**    “WebPartType"

#### Input ####

- **PreMT_MissingWebPart.csv** `(from PreMT Tool)`  
	* This is a CSV that follows the format and content of the PreMT_MissingWebPart.csv file of the Pre-Migration scan. A header row is expected to have following columns
		* PageUrl, WebUrl, StorageKey, WebPartType  
**OR**
- **WebParts_Usage.csv** `(from Discovery Tool) `
	* This is a CSV that follows the format and content of the WebParts_Usage.csv file of the Discovery. A header row is expected to have following columns
		* PageUrl, WebUrl, StorageKey, WebPartType 

#### Output ####

- **DeleteWebparts\_SuccessFailure-yyyyMMdd_hhmmss.csv**
	* In this report, details of the deleted webpart will be present.
	* This output file would contain the below mentioned columns and the entries corresponding to deleted webpart.
		* **PageUrl:** Contains path of the page from list which contains the page and page name in which page webpart type is present 
		* **StorageKey:** contains storage key of webpart on which deletion operation is performed
		* **WebpartType:** contains type of webpart on which deletion operation is performed.
		* **WebUrl:** Specifies the Web Url for the Site Collection where in webpart is present.
		* **Status:** Specifies whether the status of deletion of webpart. It contains one of the following values
			* **Success:** If value of this column is “Success” it implies that the deletion of the webpart was successful
			* **Failure:** If value of this column is “Failure” it implies that the deletion of the webpart was not successful due to some error.
- **DeleteWebparts-yyyyMMdd_hhmmss.log**
	* This is the verbose log file of the scan.  
	* Success messages of interest:
		* SUCCESS: Deleted File: {0}
	* Informational messages of interest:
		* none
	* Error messages of interest:
		* Error=File Not Found
			* Cause: The specified file or folder does not exist
			* Remediation: none; the file is gone
		* Error=The file is checked out for editing
			* Cause: someone has checked out the file for editing
			* Remediation:
				* Visit the site containing the locked file
				* Undo the check-out
				* Delete the locked file
		* (404) Not Found
			* Cause: The specified site collection does not exist
			* Remediation: none; the site collection does not exist
		* Cannot contact site at the specified URL
			* Cause: The specified web (subweb, subsite, etc.) does not exist
			* Remediation: none; the web does not exist

## <span style="color:Blue;">3 - Self Service Report</span> ##
On selecting 3rd Choice of Operation, we get the following operations as shown in the below screenshot:

![](\images/\ChoiceOfOperation3.PNG) 

### 1. Generate List Template Report with FTC Analysis ###
This operation searches for the Customized elements **(Content Types, Site Columns and Event Receivers)** after extracting the downloaded list templates.
On choosing this option, we would be asked how to proceed for this operation as shown below 

![](\images/\GenerateListTemplateReport.PNG)

	1)	Process with Auto-generated Site Collection Report  
	2)	Process with PreMT/Discovery ListTemplate Report  
	3)	Process with SiteCollectionUrls separated by comma (,)
	4)	Exit to Self Service Report Menu

#### Input ####

- Web Application Url `(Mandatory for Option 1, not for other Options)`
- Single or Multiple Site Collection Urls `(Mandatory for Option 3, not for other Options)`
- PreMT\_AllListTemplatesInGallery_Usage.csv `(Mandatory for Option 2, not for others)`
	OR  AllListTemplatesInGallery_Usage.csv`(Mandatory for Option 2, not for others)`
- CustomFields.csv `(Mandatory for all Options)`
- EventReceivers.csv `(Mandatory for all Options)`
- ContentTypes.csv `(Mandatory for all Options)`


#### Output ####

- **ListTemplateCustomization_Usage.csv**
	- If customization is available in any list templates record for any component - ContentTypes, Event Receiver and Site Column, then we'll report information related to that List Template. 
- **SiteCollections.txt** `(Output for only Option 1)`
	- This is file will list all the Site Collections that will be processed for the Web Application Url.
- **DownloadAndModifyListTemplate-yyyymmdd-hhhhmmss.log**
	* This is the verbose log file of the scan.  
	* Success messages of interest:
		* SUCCESS: Deleted File: {0}
	* Informational messages of interest:
		* none
	* Error messages of interest:
		* Error=File Not Found
			* Cause: The specified file or folder does not exist
			* Remediation: none; the file is gone
		* Error=The file is checked out for editing
			* Cause: someone has checked out the file for editing
			* Remediation:
				* Visit the site containing the locked file
				* Undo the check-out
				* Delete the locked file
		* (404) Not Found
			* Cause: The specified site collection does not exist
			* Remediation: none; the site collection does not exist
		* Cannot contact site at the specified URL
			* Cause: The specified web (subweb, subsite, etc.) does not exist
			* Remediation: none; the web does not exist

> **Note:** If any of the input files *(ContentTypes.csv, CustomFields.csv, EventReceivers.csv)* is not present in the input folder provided by the user, or the file has no entries then corresponding element/s would not be searched to get the customization details in the list templates.

> **Example:** If user has provided only *ContentTypes.csv and CustomFields.csv* in input folder, and *EventReceivers.csv* is not provide in input folder, then  *isCustomEventReceiver* column will have `NO INPUT FILE` value in output report as user has not provided this input file. 

### 2. Generate Site Template Report with FTC Analysis ###
This operation searches for the Customized elements **(Content Types, Site Columns, Features and Event Receivers)** after extracting the downloaded site templates.

On choosing this option, we would be asked how to proceed for this operation as shown below 

![](\images/\GenerateSiteTemplateReport.PNG)

	1)	Process with Auto-generated Site Collection Report  
	2)	Process with PreMT/Discovery SiteTemplate Report  
	3)	Process with SiteCollectionUrls separated by comma (,)
	4)	Exit to Self Service Report Menu

#### Input ####

- Web Application Url `(Mandatory for Option 1, not for other Options)`
- Single or Multiple Site Collection Urls `(Mandatory for Option 3, not for other Options)`
- PreMT\_AllSiteTemplatesInGallery_Usage.csv `(Mandatory for Option 2, not for other Options)`	
	OR    
AllSiteTemplatesInGallery_Usage.csv `(Mandatory for Option 2, not for other Options)`
- ContentTypes.csv `(Mandatory for all Options)`
- CustomFields.csv `(Mandatory for all Options)`
- EventReceivers.csv `(Mandatory for all Options)`
- Features.csv `(Mandatory for all Options)`


#### Output ####

- **SiteTemplateCustomization_Usage.csv**
	- If customization is available in any site templates record for any component - Features, ContentTypes, Event Receiver and Site Column, then we'll report information related to that Site Template. 
- **SiteCollections.txt** `(Output for only Option 1)`
	- This is file will list all the Site Collections that will be processed for the Web Application Url.
- **DownloadAndModifySiteTemplate-yyyymmdd-hhhhmmss.log**
	* This is the verbose log file of the scan.  
	* Success messages of interest:
		* SUCCESS: Deleted File: {0}
	* Informational messages of interest:
		* none
	* Error messages of interest:
		* Error=File Not Found
			* Cause: The specified file or folder does not exist
			* Remediation: none; the file is gone
		* Error=The file is checked out for editing
			* Cause: someone has checked out the file for editing
			* Remediation:
				* Visit the site containing the locked file
				* Undo the check-out
				* Delete the locked file
		* (404) Not Found
			* Cause: The specified site collection does not exist
			* Remediation: none; the site collection does not exist
		* Cannot contact site at the specified URL
			* Cause: The specified web (subweb, subsite, etc.) does not exist
			* Remediation: none; the web does not exist

> **Note:** If any of the input files *(Features.csv, ContentTypes.csv, CustomFields.csv, EventReceivers.csv)* is not present in the input folder provided by the user, or the file has no entries then corresponding element/s would not be searched to get the customization details in the site templates.

> **Example:** If user has provided only *Features.csv, ContentTypes.csv and CustomFields.csv* in input folder, and *EventReceivers.csv* is not provide in input folder, then  *isCustomEventReceiver* column will have `NO INPUT FILE` value in output report as user has not provided this input file.

### 3. Generate Site Column & Content Type Usage Report ###
This operation reads a list of site collection URLs from an input file and scans each site collection, looking for any web or list that is using either a custom Content Type or custom Site Column of interest.  It also looks for local Content Types that have been derived from the custom Content Types of interest.  
This report is helpful in trying to remediate the Missing Content Type and Missing Site Column reports of the Pre-Migration scan.  This report tells you where within each site collection that the content types and site columns are *still in use*.

**General Remediation:**  

- Visit the affected site collection
- If the data associated with these instances is still needed, migrate it to new content types and site columns, or move it into a spreadsheet.
- Delete all instances and empty BOTH recycle bins. 
- Clean up the definitions themselves via a temporary Sandbox Solution that uses the original Feature ID to re-deploy ONLY the affected content types and site columns.  
- Simply activate, then de-activate, the temporary sandbox feature to remove the definitions. 
- If the definitions remain, you still have some instances to delete.

#### Input ####
**1) Sites.txt**

  * This is NOT a CSV file, so no header row expected
  * Each line of this file should contain a fully-qualified, absolute site collection URL.
  * In general, this should contain the list of site collections identified in the following files of the Pre-Migration Scan:
	- PreMT_MissingSiteColumn.csv
	- PreMT_MissingContentType.csv
  * Avoid duplicate entries


**2) SiteColumns.csv**

  * The file defines the custom Site Columns of interest.
  * This is a CSV that follows the format and content of the 'PreMT_MissingSiteColumn.csv' file of the Pre-Migration scan.  A header row is expected with the following format:
SiteColumnID, SiteColumnName	
  
	- **SiteColumnID:**
      * This column should contain the GUID of the site column,
          * Take this value from the **PreMT_MissingSiteColumn.csv** file
	- **SiteColumnName**
      * This column should contain the display name of the site column
          * Take this value from the **PreMT_MissingSiteColumn.csv** file

 **3) ContentTypes.csv**

  * The file defines the custom Content Types of interest.
  * This is a CSV that follows the format and content of the **PreMT_MissingContentType.csv** file of the Pre-Migration scan. A header row is expected with the following format:
      * ContentTypeId,ContentTypeName
  * **ContentTypeId**
      * This column should contain the ID of the content type
          * Take this value from the **PreMT_MissingContentType.csv** file
  * **ContentTypeName**
      * This column should contain the display name of the content type
          * Take this value from the **PreMT_MissingContentType.csv** file

#### Output ####
- **GenerateColumnAndTypeUsageReport-yyyyMMdd_hhmmss.log**
  * This is the verbose log file of the scan.
  * Success messages of interest:
      * FOUND: Site Column [{1}] on WEB: {0}
      * FOUND: Site Column [{2}] on LIST [{0}] of WEB: {1}
      * FOUND: Content Type [{1}] on WEB: {0}
      * FOUND: Child Content Type [{2}] of [{1}] on WEB: {0}
      * FOUND: Content Type [{2}] on LIST [{0}] of WEB: {1}
      * FOUND: Child Content Type [{3}] of [{2}] on LIST [{0}] of WEB: {1}
  * Informational messages of interest:
      * <span style="background-color: #FFFF00">TBD</span>
  * Error messages of interest:
      * <span style="background-color: #FFFF00">TBD</span>

### 4. Generate Non-Default Master Page Usage Report ###
This operation reads a list of site collection URLs from an input file and scans each site collection, looking for any web that is using a non-default SP2013 Master Page (i.e., something other than seattle.master) as either its System or Site master page. 

On choosing this option, we would be asked how to proceed for this operation as shown below 

![](\images/\GenerateNonDefaultMasterPageUsageReport.PNG)


#### Input ####
- **Sites.txt**
  * This is NOT a CSV file, so no header row expected
  * Each line of this file should contain a fully-qualified, absolute site collection URL.
  * Avoid duplicate entries

#### Output ####
- **GenerateNonDefaultMasterPageUsageReport-yyyyMMdd_hhmmss.log**
  * This is the verbose log file of the scan.
  * Success messages of interest:
      * FOUND: System Master Page setting (Prop=MasterUrl) of web {0} is {1}
      * FOUND: Site Master Page setting (Prop=CustomMasterUrl) of web {0} is {1}
  * Informational messages of interest:
      * <span style="background-color: #FFFF00">TBD</span>
  * Error messages of interest:
      * <span style="background-color: #FFFF00">TBD</span>

### 5. Generate Site Collection Report (PPE-Only) ###
This operation generates a text file containing a list of all site collections found across all web applications in the target farm. 

On choosing this option, we would be asked how to proceed for this operation as shown below 

![](\images/\GenerateSiteCollectionReport.PNG)


#### Input ####
- **None**

#### Output ####
- **GenerateSiteCollectionReport-yyyyMMdd_hhmmss.txt**
  * This is NOT a CSV file, so no header row generated
  * Each line of the file will contain a fully-qualified, absolute site collection URL.

#### NOTES ####
- This operation is intended for use only in PPE; use on PROD at your own risk.  For PROD, it is safer to generate the report via the o365 Self-Service Admin Portal.
- This operation leverages the Search Index; as such, it might take up to 20-minutes for a newly-created site to appear in the report. 

### 6. Get Web Part Usage Report ###

This operation iterates through all the Pages in root folder of the web, “Pages” and “Site Pages” Library and gives the usage of the given Web Part.

#### Input ####
On selecting this operation, we will be asked for no of parameters as explained below:

![](\images/\GenerateWebPartUsage.PNG)

- **Web Url:**
	* Here we need to provide the Web Url where we need to add the WebPart 
	**Example:**    “https://intranet.campoc.com/sites/OffshorePoc"
- **WebPart Type:**
	* Here we need to provide the WebPart Type which we need to be replaced  
	**Example:**    “ContentEditorWebPart”

#### Output ####
- **WebPartUsage.csv**
  * In this report, we’ll get the Usage report of Web Parts in the farm.
  * This output file would contain the below mentioned columns and the entries corresponding to list templates.
  		* **PageUrl:** Specifies the Url of the page where the web part is present.
  		* **WebApplication:** Specifies the Web Application of the Site Collection where in the Web Part is present.
  		* **SiteCollection:** Specifies the Site Collection where in the Web Part is present.
  		* **WebUrl:** Specifies the Web Url for the Site Collection where in the WebPart is present.
  		* **WebPartId:** Specifies Web Part Id of the web part
  		* **WebPartTitle:** Specifies the Web Part title of the web part
  		* **WebPartType:** Specifies the Web Part Type of the web part
  		* **ZoneIndex:** Specifies the Zone Index of the web part  	
- **WebPartUsage-yyyyMMdd_hhmmss.log**
	* This is the verbose log file of the scan.  
	* Success messages of interest:
		* SUCCESS: Deleted File: {0}
	* Informational messages of interest:
		* none
	* Error messages of interest:
		* Error=File Not Found
			* Cause: The specified file or folder does not exist
			* Remediation: none; the file is gone
		* Error=The file is checked out for editing
			* Cause: someone has checked out the file for editing
			* Remediation:
				* Visit the site containing the locked file
				* Undo the check-out
				* Delete the locked file
		* (404) Not Found
			* Cause: The specified site collection does not exist
			* Remediation: none; the site collection does not exist
		* Cannot contact site at the specified URL
			* Cause: The specified web (subweb, subsite, etc.) does not exist
			* Remediation: none; the web does not exist

### 7. Get Web Part Properties ###

This operation will Returns the properties of the given Web Part in Xml format.

#### Input ####
On selecting this operation, we will be asked for no of parameters as explained below:

![](\images/\GenerateWebPartProperties.PNG)

- **Web Url:**
	* Here we need to provide the Web Url where we need to add the WebPart 
	**Example:**    “https://intranet.campoc.com/sites/OffshorePoc"
- **Server Relative PageUrl:**
	* Here we need to provide the server relative url of the page in Web Url where we need to add the WebPart  
	**Example:**    “/sites/OffshorePoc/SitePages/DevHome.aspx”
- **WebPart ID:**
	* Here we need to provide the xml file path of the web part we need to add
	**Example:**    “de5c57f9-7991-4ba7-b141-9db5d48393fc"

#### Output ####
- **WebPartID(provided in input)_WebPartProperties.xml**
  * Creates Property Xml file for the corresponding Web Part. This file will be available inside SourceWebPartXmls Folder.  	
- **WebpartProperties-yyyyMMdd_hhmmss.log**
	* This is the verbose log file of the scan.  
	* Success messages of interest:
		* SUCCESS: Deleted File: {0}
	* Informational messages of interest:
		* none
	* Error messages of interest:
		* Error=File Not Found
			* Cause: The specified file or folder does not exist
			* Remediation: none; the file is gone
		* Error=The file is checked out for editing
			* Cause: someone has checked out the file for editing
			* Remediation:
				* Visit the site containing the locked file
				* Undo the check-out
				* Delete the locked file
		* (404) Not Found
			* Cause: The specified site collection does not exist
			* Remediation: none; the site collection does not exist
		* Cannot contact site at the specified URL
			* Cause: The specified web (subweb, subsite, etc.) does not exist
			* Remediation: none; the web does not exist













 

