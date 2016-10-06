# HTTP Remote Operations Pattern #

### Summary ###
This sample demonstrates how to perform remote operations towards configurations which are not natively available using CSOM. Technique is to mimic operations what users do when they use browser.

### Applies to ###
-  Office 365 Multi Tenant (MT)
-  Office 365 Dedicated (D)
-  SharePoint 2013 on-premises


### Solution ###
Solution | Author(s)
---------|----------
JDP.Transformation.HttpCommands | Vesa Juvonen, Bert Jansen, Antons Mislevics (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | September 26th 2016 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Overview #
SharePoint remote APIs are really powerful and can be used for managing broad spectrum of the different options in site or in list level. There are however still some limitations compared to the functionalities which have been available using server side APIs. These limitations are being addressed gradually in the client side object model or in other remote techniques, but as a short term workaround we can quite often use so called HTTP remote operations pattern to achieve the needed functionality.

**Notice** that this pattern should only be used when there�s no alternatives with out of the box APIs. This pattern will also cause dependency on the html and control structure on the pages, so if control identifiers will change, code has to be changed as well.

# Solution #
This solution implements HTTP remote operations pattern for the following scenarios:
- Add an app;
- Register an app;
- Trust an app;
- Get a list of available apps;
- Activate sandbox solution;
- Deactivate sandbox solution.

<img src="https://telemetry.sharepointpnp.com/pnp-transformation/JDP.Transformation.HttpCommands" />
