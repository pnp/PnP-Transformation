# Feature Activation Blocker #

### Summary ###
This sample shows an DisableFeatureActivation solution that is used to block feature activation.

### Applies to ###
-  Office 365 Dedicated (D)
-  SharePoint 2013 on-premises


### Solution ###
Solution | Author(s)
---------|----------
Feature Activation Blocker | Antons Mislevics (**Microsoft**), Infosys Ltd

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | February 04th 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# The Process for Disable Features Activation/Deactivation #
The process of disabling the Site-Web scoped features activation/deactivation is performed with the solution ”DisableFeatureActivation.wsp” installation and enabling the web application feature “JDP.Transformation.DisableFeatureActivation_Control”.  

Package should be prepared with the following: 
<ol type="a">
  <li>Enable, Disable webapplication feature scripts in the Installation Scripts folder. </li>
  <li>The Deployment Guide Post-Deployment section with instructions on how to disable the feature activation/deactivation.</li>
</ol>

Disable Features activation/deactivation validation Steps:  
<ol type="1">
   <li>Ensure that the solution whose feature activation/deactivation to be disabled should be in deployed state</li> <li>”DisableFeatureActivation.wsp” solution deployment to be performed using MSOCAF tool Update the variables $urls with the web application urls in “EnableWebApplicationFeature.Ps1”. Execute “EnableWebApplicationFeature.Ps1”. This script activates the feature “JDP.Transformation.DisableFeatureActivation_Control” for the provided urls in the variables $urls. This feature activation creates lists “Disable Features List” and "Disable Feature Message" in the root site collection for the provided urls in the variables $urls.</li>
	<li>Customer can follow below options to create/update the lists that required to disable features and to display customer message on features page.
<ol> 
Option 1: 
<ol type="a">
<li>Customer provide the list of features (i.e. feature ID) to be updated in the “Disable Features List” list under ‘FeatureID’ column</li>
<li>Customer provide Title and Description to be updated in the “Disable Feature Message” list with single item</li>
</ol>  
Option 2: After this change is processed, customer updates both “Disable Features List” and “Disable Feature Message” lists directly<br/>
Option 3: Customer creates “Disable Features List” and “Disable Feature Message” lists before this change. Updates the “Disable Features List” list with list of feature IDs to be disabled. Updates “Disable Feature Message” list with custom message to be displayed on features page.  
</ol>
</li>
	<li>Verify the specific site-web features added to the list, the Activate/Deactivate button should be disabled</li>
</ol>

Rollback steps:  
<ol type="1">
<li>Update the variables $urls with the web application urls in “DisableWebApplicationFeature.Ps1”. Execute “DisableWebApplicationFeature.Ps1”. This script deactivates the feature “JDP.Transformation.DisableFeatureActivation_Control” for the provided urls in the variables $urls. This feature deactivation deletes the lists “Disable Features List” and "Disable Feature Message".</li>
<li>Rollback the “JDP.Transformation.DisableFeatureActivation.wsp” solution using MSOCAF tool. </li>
</ol>


<img src="https://telemetry.sharepointpnp.com/pnp-transformation/feature-activation-blocker" /> 