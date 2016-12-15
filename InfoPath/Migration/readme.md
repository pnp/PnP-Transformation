This section contains utilities that are needed when you migrate InfoPath forms to SharePoint Online (DvNext / MT).

# EmpRegConsole #
Once you transform an InfoPath form to a SharePoint Add-In you also need to migrate the data from InfoPath XML to a new model. This sample provides you more input on how to that.

# UdcxRemediation.Console #
Forms calling into OOB ASMX services work differently in SharePoint Online versus SharePoint on-premises. The udcx remediation tool will help to automatically fix the udcx files in your environment.

# PeoplePickerRemediation.Console #
Migrated InfoPath form's people picker control contains On-Prem user information. The People Picker remediation tool will help to automatically fix data of People Picker control in your environment.

# OWSSVR (RPC calls)
Forms using the RPC calls by calling owssvr.dll will not work in SharePoint Online. The guidance provided will show you how to remediate these forms by replacing the RPC call with an InfoPath SharePointDataConnection

<img src="https://telemetry.sharepointpnp.com/pnp-transformation/infopath/migration" /> 
