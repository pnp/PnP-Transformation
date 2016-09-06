# Migrate InfoPath XML to SharePoint list items #

### Summary ###
This console application shows how you can migrate data coming from InfoPath XML files into SharePoint list items

### Applies to ###
-  Office 365
-  SharePoint 2013/2016 on-premises


### Solution ###
Solution | Author(s)
---------|----------
EmpRegConsole | Raja Shekar Bhumireddy, Bert Jansen (**Microsoft**)

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | August 27th 2015 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Introduction#
When an existing InfoPath form is transformed to a new solution like a SharePoint Add-In it often also means that the existing data created using that InfoPath form needs to be retained. InfoPath by default persists data using the InfoPath XML specification...so we have a set of InfoPath XML files. This simple console application shows how you can read those InfoPath XML files, read the relevant data from them and create equivalent items in a SharePoint list.

# Important coding techniques #
## Getting the data from the InfoPath XML files ##
Below snippet shows how to grab relevant properties from the InfoPath XML file.

```C#
File ipFile = item.File;
ClientResult<System.IO.Stream> streamItem = ipFile.OpenBinaryStream();
clientContext.Load(ipFile);
clientContext.ExecuteQuery();

System.IO.MemoryStream memStream = new System.IO.MemoryStream();
streamItem.Value.CopyTo(memStream);
memStream.Position = 0;

XmlDocument ipXML = new XmlDocument();            
ipXML.Load(memStream);
XmlNamespaceManager ns = new XmlNamespaceManager(ipXML.NameTable);           
ns.AddNamespace("my", ipXML.DocumentElement.NamespaceURI);

XPathNavigator empNavigator = ipXML.CreateNavigator();
employee = new Employees(); 
employee.Name = empNavigator.SelectSingleNode("/my:EmployeeForm/my:txtName", ns).Value;
employee.UserID = empNavigator.SelectSingleNode("/my:EmployeeForm/my:txtUserID", ns).Value;
employee.Manager = empNavigator.SelectSingleNode("/my:EmployeeForm/my:txtManager", ns).Value;
employee.Number = empNavigator.SelectSingleNode("/my:EmployeeForm/my:txtEmpNumber", ns).Value;
employee.Designation = empNavigator.SelectSingleNode("/my:EmployeeForm/my:ddlDesignation", ns).Value;
employee.Locaiton = empNavigator.SelectSingleNode("/my:EmployeeForm/my:ddlCity", ns).Value;

XmlNodeList nodeSkills = ipXML.SelectNodes("//my:Skill", ns);
StringBuilder sbSkills = new StringBuilder();
foreach (XmlNode nodeSkill in nodeSkills)
{
    XmlNodeList lstSkill = nodeSkill.ChildNodes;
    sbSkills.Append(lstSkill[0].InnerText).Append(",").Append(lstSkill[1].InnerText).Append(";");
} //foreach (XmlNode nodeSkill in nodeSkills)

employee.Skills= sbSkills.ToString();
```

Notice how we transform a multi-value InfoPath item (```Skill```) to a comma delimited list in this implementation, but one could also store data in different SharePoint lists if that would be more appropriate.

## Authenticating to SharePoint ##
This sample is taken a simple approach that works in on-premises by simply taking the credentials of the user executing the code:

```C#
using (ClientContext clientContext = new ClientContext(sharePointSiteUrl))
{
    Employees employee;
    Web web = clientContext.Web;
    List libInfoPath = web.Lists.GetByTitle(infoPathLibName); // EmpRegLib

    var ipItems = libInfoPath.GetItems(CamlQuery.CreateAllItemsQuery());
    clientContext.Load(ipItems);
    clientContext.ExecuteQuery();

    foreach (ListItem item in ipItems)
    {                                               
        ReadInfoPathFile(clientContext, web, item, out employee);
        AddInfoPathToList(clientContext, web, employee);
    } // foreach (ListItem item in ipItems)
}
```

If you're migrating from an Office 365 list with InfoPath XML files or in case you want to specify a different user please checkout the [AuthenticationManager class](https://github.com/OfficeDev/PnP-Sites-Core/blob/dev/Core/OfficeDevPnP.Core/AuthenticationManager.cs) and [documentation](https://github.com/OfficeDev/PnP-Sites-Core/blob/dev/Core/README.md#authenticationmanagercs) from the [PnP-Sites-Core repository](https://github.com/OfficeDev/PnP-Sites-Core/tree/dev).

<img src="https://telemetry.sharepointpnp.com/pnp-transformation/infopath/migration/empregconsole" /> 