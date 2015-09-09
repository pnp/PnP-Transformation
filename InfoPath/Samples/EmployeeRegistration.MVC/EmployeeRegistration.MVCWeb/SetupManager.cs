using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System;
using System.Configuration;
using System.Web;
using System.Web.Hosting;

namespace EmployeeRegistration.MVCWeb
{
    public static class SetupManager
    {

        static SetupManager()
        {
            Initialized = false;
        }

        public static bool Initialized
        {
            get; set;
        }

        #region Application provisioning
        public static void Provision(ClientContext ctx)
        {
            if (!SetupManager.Initialized)
            {
                // Provision lists
                SetupManager.ProvisionLists(ctx);

                // upload assets and provision the application
                ProvisionEmployeeRegistrationApplication(ctx);

                SetupManager.Initialized = true;
            }
        }

        private static void ProvisionEmployeeRegistrationApplication(ClientContext ctx)
        {
            // upload files to the style library
            List styleLibrary = ctx.Web.Lists.GetByTitle("Style Library");
            ctx.Load(styleLibrary, l => l.RootFolder);
            Folder pnpFolder = styleLibrary.RootFolder.EnsureFolder("OfficeDevPnP");

            string fileName = "AppLauncher.js";
            File assetFile = pnpFolder.GetFile(fileName);
            if (assetFile != null)
            {
                assetFile.CheckOut();
            }

            string clientId = ConfigurationManager.AppSettings["ClientId"];
            Uri redirectUri = new Uri(HttpContext.Current.Request.Url.AbsoluteUri);
            string redirectUrl = String.Format("{0}://{1}{2}", redirectUri.Scheme, redirectUri.Authority, redirectUri.LocalPath);

            string localFilePath = "Assets/" + fileName;
            string newLocalFilePath = Utilities.ReplaceTokensInAssetFile(ctx, localFilePath, clientId, redirectUrl);

            assetFile = pnpFolder.UploadFile(fileName, newLocalFilePath, true);
            assetFile.CheckIn("Uploaded by provisioning engine.", CheckinType.MajorCheckIn);
            ctx.ExecuteQuery();
            System.IO.File.Delete(newLocalFilePath);

            Web web = ctx.Web;
            ctx.Load(web, w => w.ServerRelativeUrl);
            ctx.ExecuteQuery();
            Console.WriteLine("");

            // Inject the content editor web parts on the new and edit form pages
            if (!SetupManager.IsWebPartOnPage(ctx, String.Format("{0}/Lists/Employees/newform.aspx", web.ServerRelativeUrl), "Employee Registration"))
            {
                string newFormUrl = string.Format("{0}/{1}", web.ServerRelativeUrl, "Lists/Employees/newform.aspx");
                SetupManager.CloseAllWebParts(ctx, newFormUrl);
                ProvisionWebPart(ctx, newFormUrl, "AppLauncher.js");
            }
            else
            {
                Console.WriteLine("The New form page was already customized");
            }

            if (!SetupManager.IsWebPartOnPage(ctx, String.Format("{0}/Lists/Employees/editform.aspx", web.ServerRelativeUrl), "Employee Registration"))
            {
                string editFormUrl = string.Format("{0}/{1}", web.ServerRelativeUrl, "Lists/Employees/editform.aspx");
                SetupManager.CloseAllWebParts(ctx, editFormUrl);
                ProvisionWebPart(ctx, editFormUrl, "AppLauncher.js");
            }
            else
            {
                Console.WriteLine("The New form page was already customized");
            }
        }

        #endregion

        #region provision data model
        public static void ProvisionLists(ClientContext ctx)
        {
            //Create country list
            Console.WriteLine("Provisioning lists:");
            List countryList = null;
            if (!ctx.Web.ListExists("EmpCountry"))
            {
                Console.WriteLine("Country list...");
                countryList = ctx.Web.CreateList(ListTemplateType.GenericList, "EmpCountry", false, false, "Lists/EmpCountry", false);

                //Provision country list items
                ListItemCreationInformation newCountryCreationInfomation;
                newCountryCreationInfomation = new ListItemCreationInformation();
                ListItem newCountry = countryList.AddItem(newCountryCreationInfomation);
                newCountry["Title"] = "Belgium";
                newCountry.Update();
                newCountry = countryList.AddItem(newCountryCreationInfomation);
                newCountry["Title"] = "United States of America";
                newCountry.Update();
                newCountry = countryList.AddItem(newCountryCreationInfomation);
                newCountry["Title"] = "India";
                newCountry.Update();
                ctx.Load(countryList);
                ctx.ExecuteQueryRetry();
            }
            else
            {
                countryList = ctx.Web.GetListByUrl("Lists/EmpCountry");
                Console.WriteLine("Country list was already available");
            }

            List stateList = null;
            if (!ctx.Web.ListExists("EmpState"))
            {
                Console.WriteLine("State list...");
                stateList = ctx.Web.CreateList(ListTemplateType.GenericList, "EmpState", false, false, "Lists/EmpState", false);
                Field countryLookup = stateList.CreateField(@"<Field Type=""Lookup"" DisplayName=""Country"" ID=""{BDEF775C-AB4B-4E86-9FB8-0A2DE40FE832}"" Name=""Country""></Field>", false);
                ctx.Load(stateList);
                ctx.Load(countryLookup);
                ctx.Load(stateList.DefaultView, p => p.ViewFields);
                ctx.ExecuteQueryRetry();

                // Add field to default view               
                stateList.DefaultView.ViewFields.Add("Country");
                stateList.DefaultView.Update();
                ctx.ExecuteQueryRetry();

                // configure country lookup field
                FieldLookup countryField = ctx.CastTo<FieldLookup>(countryLookup);
                countryField.LookupList = countryList.Id.ToString();
                countryField.LookupField = "Title";
                countryField.Indexed = true;
                countryField.IsRelationship = true;
                countryField.RelationshipDeleteBehavior = RelationshipDeleteBehaviorType.Restrict;
                countryField.Update();
                ctx.ExecuteQueryRetry();

                //Provision state list items
                ListItemCreationInformation newStateCreationInfomation;
                newStateCreationInfomation = new ListItemCreationInformation();
                ListItem newState = stateList.AddItem(newStateCreationInfomation);
                newState["Title"] = "Washington";
                newState["Country"] = "2;#United States of America";
                newState.Update();
                newState = stateList.AddItem(newStateCreationInfomation);
                newState["Title"] = "Limburg";
                newState["Country"] = "1;#Belgium";
                newState.Update();
                newState = stateList.AddItem(newStateCreationInfomation);
                newState["Title"] = "Tennessee";
                newState["Country"] = "2;#United States of America";
                newState.Update();
                newState = stateList.AddItem(newStateCreationInfomation);
                newState["Title"] = "Karnataka";
                newState["Country"] = "3;#India";
                newState.Update();

                ctx.ExecuteQueryRetry();
            }
            else
            {
                countryList = ctx.Web.GetListByUrl("Lists/EmpState");
                Console.WriteLine("State list was already available");
            }

            List cityList = null;
            if (!ctx.Web.ListExists("EmpCity"))
            {
                Console.WriteLine("City list...");
                cityList = ctx.Web.CreateList(ListTemplateType.GenericList, "EmpCity", false, false, "Lists/EmpCity", false);
                Field stateLookup = cityList.CreateField(@"<Field Type=""Lookup"" DisplayName=""State"" ID=""{F55BED78-CAF9-4EDF-92B9-C46BDC032DD5}"" Name=""State""></Field>", false);
                ctx.Load(cityList);
                ctx.Load(stateLookup);
                ctx.Load(cityList.DefaultView, p => p.ViewFields);
                ctx.ExecuteQueryRetry();

                // Add field to default view               
                cityList.DefaultView.ViewFields.Add("State");
                cityList.DefaultView.Update();
                ctx.ExecuteQueryRetry();

                // configure state lookup field
                FieldLookup stateField = ctx.CastTo<FieldLookup>(stateLookup);
                stateField.LookupList = stateList.Id.ToString();
                stateField.LookupField = "Title";
                stateField.Indexed = true;
                stateField.IsRelationship = true;
                stateField.RelationshipDeleteBehavior = RelationshipDeleteBehaviorType.Restrict;
                stateField.Update();
                ctx.ExecuteQueryRetry();

                //Provision city list items
                ListItemCreationInformation newCityCreationInfomation;
                newCityCreationInfomation = new ListItemCreationInformation();
                ListItem newCity = cityList.AddItem(newCityCreationInfomation);
                newCity["Title"] = "Bree";
                newCity["State"] = "2;#Limburg";
                newCity.Update();
                newCity = cityList.AddItem(newCityCreationInfomation);
                newCity["Title"] = "Redmond";
                newCity["State"] = "1;#Washington";
                newCity.Update();
                newCity = cityList.AddItem(newCityCreationInfomation);
                newCity["Title"] = "Franklin";
                newCity["State"] = "3;#Tennessee";
                newCity.Update();
                newCity = cityList.AddItem(newCityCreationInfomation);
                newCity["Title"] = "Bangalore";
                newCity["State"] = "4;#Karnataka";
                newCity.Update();

                ctx.ExecuteQueryRetry();
            }
            else
            {
                cityList = ctx.Web.GetListByUrl("Lists/EmpCity");
                Console.WriteLine("City list was already available");
            }

            List designationList = null;
            if (!ctx.Web.ListExists("EmpDesignation"))
            {
                Console.WriteLine("Designation list...");
                designationList = ctx.Web.CreateList(ListTemplateType.GenericList, "EmpDesignation", false, false, "Lists/EmpDesignation", false);
                ctx.Load(designationList);
                ctx.ExecuteQueryRetry();

                //Provision designation list items
                ListItemCreationInformation newDesignationCreationInfomation;
                newDesignationCreationInfomation = new ListItemCreationInformation();
                ListItem newDesignation = designationList.AddItem(newDesignationCreationInfomation);
                newDesignation["Title"] = "Service Engineer";
                newDesignation.Update();
                newDesignation = designationList.AddItem(newDesignationCreationInfomation);
                newDesignation["Title"] = "Service Engineer II";
                newDesignation.Update();
                newDesignation = designationList.AddItem(newDesignationCreationInfomation);
                newDesignation["Title"] = "Senior Service Engineer";
                newDesignation.Update();
                newDesignation = designationList.AddItem(newDesignationCreationInfomation);
                newDesignation["Title"] = "Principal Service Engineer";
                newDesignation.Update();
                newDesignation = designationList.AddItem(newDesignationCreationInfomation);
                newDesignation["Title"] = "Program Manager";
                newDesignation.Update();
                newDesignation = designationList.AddItem(newDesignationCreationInfomation);
                newDesignation["Title"] = "Senior Program Manager";
                newDesignation.Update();

                ctx.ExecuteQueryRetry();
            }
            else
            {
                designationList = ctx.Web.GetListByUrl("Lists/EmpDesignation");
                Console.WriteLine("Designation list was already available");
            }

            List employeeList = null;
            if (!ctx.Web.ListExists("Employees"))
            {
                Console.WriteLine("Employee list...");
                employeeList = ctx.Web.CreateList(ListTemplateType.GenericList, "Employees", false, false, "Lists/Employees", false);
                ctx.Load(employeeList);
                ctx.ExecuteQueryRetry();

                employeeList.CreateField(@"<Field Type=""Text"" DisplayName=""Number"" ID=""{0FFF75B6-0E57-46BE-84D7-E5225A55E914}"" Name=""EmpNumber""></Field>", false);
                employeeList.CreateField(@"<Field Type=""Text"" DisplayName=""UserID"" ID=""{D3199531-C091-4359-8CDE-86FB3924F65E}"" Name=""UserID""></Field>", false);
                employeeList.CreateField(@"<Field Type=""Text"" DisplayName=""Manager"" ID=""{B7E2F4D9-AEA2-40DB-A354-94EEDAFEF35E}"" Name=""EmpManager""></Field>", false);
                employeeList.CreateField(@"<Field Type=""Text"" DisplayName=""Designation"" ID=""{AB230804-C137-4ED8-A6D6-722037BDDA3D}"" Name=""Designation""></Field>", false);
                employeeList.CreateField(@"<Field Type=""Text"" DisplayName=""Location"" ID=""{2EE32832-5EF0-41D0-8CD3-3DE7B9616C21}"" Name=""Location""></Field>", false);
                employeeList.CreateField(@"<Field Type=""Text"" DisplayName=""Skills"" ID=""{89C02660-822D-4F41-881D-1D533C56017E}"" Name=""Skills""></Field>", false);
                employeeList.Update();
                ctx.Load(employeeList.DefaultView, p => p.ViewFields);
                ctx.ExecuteQueryRetry();

                // Add fields to view
                employeeList.DefaultView.ViewFields.Add("EmpNumber");
                employeeList.DefaultView.ViewFields.Add("UserID");
                employeeList.DefaultView.ViewFields.Add("EmpManager");
                employeeList.DefaultView.ViewFields.Add("Designation");
                employeeList.DefaultView.ViewFields.Add("Location");
                employeeList.DefaultView.ViewFields.Add("Skills");
                employeeList.DefaultView.Update();
                ctx.ExecuteQueryRetry();

            }
            else
            {
                employeeList = ctx.Web.GetListByUrl("Lists/Employees");
                Console.WriteLine("Employee list was already available");
            }
        }
        #endregion

        #region web part manipulation
        private static void ProvisionWebPart(ClientContext ctx, string relativePageUrl, string scriptFile, bool isWikiPage = false)
        {
            Console.WriteLine("Provisioning web part...");

            string webPartXml = System.IO.File.ReadAllText(HostingEnvironment.MapPath(String.Format("~/{0}", @"Assets\EmployeeRegistration.dwp")));
            //replace tokens
            string scriptUrl = String.Format("~sitecollection/Style Library/OfficeDevPnP/{0}", scriptFile);
            scriptUrl = Utilities.ReplaceTokens(ctx, scriptUrl);
            webPartXml = webPartXml.Replace("%ContentLink%", scriptUrl);

            OfficeDevPnP.Core.Entities.WebPartEntity webPart = new OfficeDevPnP.Core.Entities.WebPartEntity()
            {
                WebPartZone = "Main",
                WebPartIndex = 20,
                WebPartTitle = "Employee Registration",
                WebPartXml = webPartXml
            };

            Console.WriteLine("Adding employee registration web part to " + relativePageUrl);
            if (isWikiPage)
            {
                ctx.Web.AddWebPartToWikiPage(relativePageUrl, webPart, 1, 1, false);
            }
            else
            {
                ctx.Web.AddWebPartToWebPartPage(relativePageUrl, webPart);
            }
            Console.WriteLine("");
        }

        public static bool IsWebPartOnPage(ClientContext ctx, string relativePageUrl, string title)
        {
            var webPartPage = ctx.Web.GetFileByServerRelativeUrl(relativePageUrl);
            ctx.Load(webPartPage);
            ctx.ExecuteQuery();

            if (webPartPage == null)
            {
                return false;
            }

            LimitedWebPartManager limitedWebPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            ctx.Load(limitedWebPartManager.WebParts, wps => wps.Include(wp => wp.WebPart.Title));
            ctx.ExecuteQueryRetry();

            if (limitedWebPartManager.WebParts.Count >= 0)
            {
                for (int i = 0; i < limitedWebPartManager.WebParts.Count; i++)
                {
                    WebPart oWebPart = limitedWebPartManager.WebParts[i].WebPart;
                    if (oWebPart.Title.Equals(title, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public static void CloseAllWebParts(ClientContext ctx, string relativePageUrl)
        {
            var webPartPage = ctx.Web.GetFileByServerRelativeUrl(relativePageUrl);
            ctx.Load(webPartPage);
            ctx.ExecuteQuery();

            if (webPartPage == null)
            {
                return;
            }

            LimitedWebPartManager limitedWebPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            ctx.Load(limitedWebPartManager.WebParts, wps => wps.Include(wp => wp.WebPart.Title));
            ctx.ExecuteQueryRetry();

            if (limitedWebPartManager.WebParts.Count >= 0)
            {
                for (int i = 0; i < limitedWebPartManager.WebParts.Count; i++)
                {
                    limitedWebPartManager.WebParts[i].CloseWebPart();
                    limitedWebPartManager.WebParts[i].SaveWebPartChanges();
                }
                ctx.ExecuteQuery();
            }
        }
        #endregion

    }
}
