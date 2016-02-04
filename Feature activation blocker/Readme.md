# Feature Activation Blocker #

### Summary ###
This sample shows an DisableFeatureActivation application that is used to block feature activation.

### Applies to ###
-  Office 365 Dedicated (D)
-  SharePoint 2013 on-premises


### Solution ###
Solution | Author(s)
---------|----------
Feature Activation Blocker | Infosys Ltd

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | February 04th 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# The Process for Disable Features Activation/Deactivation #
The process of disabling the Site-Web scoped features activation/deactivation is performed with the solution ”JDP.Transformation.DisableFeatureActivation.wsp” installation and enabling the web application feature “JDP.Transformation.DisableFeatureActivation_Control”.  

Package should be prepared with the following: 
<ol type="a">
  <li>Create List and Enable, Disable webapplication feature scripts in the Installation Scripts folder. </li>
  <li>The Deployment Guide Post-Deployment section with instructions on how to disable the feature activation/deactivation.</li>
  <li>The Test Results document in Review Reports folder.</li>
</ol>

Disable Features activation/deactivation validation Steps:  
1. Ensure that the solution whose feature activation/deactivation to be disabled should be in deployed state  
2. "JDP.Transformation.DisableFeatureActivation.wsp” solution deployment to be performed using MSOCAF tool  
3. Update the variables $urls with the web application urls in “EnableWebApplicationFeature.Ps1”. Execute “EnableWebApplicationFeature.Ps1”. This script activates the feature   “JDP.Transformation.DisableFeatureActivation_Control” for the provided urls in the variables $urls.  
4. Update the variables $urls with the web application urls in “CreateFeaturesList.Ps1”. Execute “CreateFeaturesList.Ps1”. This script creates list “Disable Features List” in the root site collection for the provided urls in the variables $urls.  
5. Add the feature IDs in the List “Disable Features List” to the fields “Title”,”FeatureID”. The feature activation/deactivation is disabled for the feature IDs added in this list. <span style="text-decoration:underline">**Note:**</span> Customer can follow below options to create/update the list of features to be disabled  
<span style="text-decoration:underline">Option 1:</span> Customer provide the list of features (i.e. feature ID) to be updated in the list under ‘FeatureID’ column  
<span style="text-decoration:underline">Option 2:</span> After this change is processed, customer update the list directly  
<span style="text-decoration:underline">Option 3:</span> Customer creates the list before this change & update the list of feature IDs to be disabled  
6. Verify the specific site-web features added to the list, the Activate/Deactivate button should be disabled  

Rollback steps:  
1. Update the variables $urls with the web application urls in “DeleteFeaturesList.ps1”. Execute “DeleteFeaturesList.ps1”. This script deletes **“Disable Features List”** list from all web sites where it was created as part of deployment.   
2. Update the variables $urls with the web application urls in “DisableWebApplicationFeature.Ps1”. Execute “DisableWebApplicationFeature.Ps1”. This script deactivates the feature “JDP.Transformation.DisableFeatureActivation_Control” for the provided urls in the variables $urls.  
3. Rollback the ”JDP.Transformation.DisableFeatureActivation.wsp” solution using MSOCAF tool. 


