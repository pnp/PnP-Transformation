using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System;
using System.Security;

namespace KnockOut.SinglePageApp
{
    class Program
    {

        // TODO: update values before running the sample or blank them which 
        // triggers a prompt to ask for the value
        static string siteUrl = "https://bertonline.sharepoint.com/sites/devinfopath";
        static string username = "bert.jansen@bertonline.onmicrosoft.com";
        static SecureString password = null;

        static void Main(string[] args)
        {
            ConsoleColor current = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Provisioning of the KnockOut.SinglePageApp sample started...");
            Console.ForegroundColor = current;
            Console.WriteLine("");

            #region Get information about the site to deploy to
            // Request Office365 site from the user
            if (String.IsNullOrEmpty(siteUrl))
            {
                siteUrl = GetSite();
            }

            // Prompt for username 
            if (String.IsNullOrEmpty(username))
            {
                username = GetUserName();
            }

            // Prompt for password
            if (password == null || password.Length == 0)
            {
                password = GetPassword();
            }
            #endregion

            #region Deploy assets and lists
            ClientContext ctx = CreateContext();
            // Provision supporting js files to the Style Library
            ProvisionAssets(ctx);

            // Provision lists and items
            ProvisionLists(ctx);

            // Provision the employee registration application
            ProvisionEmployeeRegistrationApplication(ctx);
            #endregion
            
            Console.WriteLine("Provisioning of the KnockOut.SinglePageApp sample is done.");
            Console.WriteLine("");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Open a browser and navigate to {0} to start the application.", String.Format("{0}/Lists/Employees", siteUrl));
            Console.ForegroundColor = current;
            Console.WriteLine("");
            Console.WriteLine("Press <enter> to continue...");
            Console.ReadLine();
        }

        #region helper methods
        private static void ProvisionEmployeeRegistrationApplication(ClientContext ctx)
        {
            Web web = ctx.Web;
            ctx.Load(web, w => w.ServerRelativeUrl);
            ctx.ExecuteQuery();
            Console.WriteLine("");

            string empFormPage = ctx.Web.AddWikiPage("Site Pages", "EmpForm.aspx");
            if (!String.IsNullOrEmpty(empFormPage))
            {
                Console.WriteLine("Provisioning EmpForm.aspx...");
                ProvisionWebPart(ctx, string.Format("{0}/{1}", web.ServerRelativeUrl, empFormPage), "Emp-Registration-Form-Template.js");
            }
            else
            {
                Console.WriteLine("EmpForm.aspx was already added");
            }

            if (!IsWebPartOnPage(ctx, String.Format("{0}/Lists/Employees/newform.aspx",web.ServerRelativeUrl), "Employee Registration"))
            {
                string newFormUrl = string.Format("{0}/{1}", web.ServerRelativeUrl, "Lists/Employees/newform.aspx");
                CloseAllWebParts(ctx, newFormUrl);
                ProvisionWebPart(ctx, newFormUrl, "Emp-NewForm.js");
            }
            else
            {
                Console.WriteLine("The New form page was already customized");
            }

            if (!IsWebPartOnPage(ctx, String.Format("{0}/Lists/Employees/editform.aspx", web.ServerRelativeUrl), "Employee Registration"))
            {
                string editFormUrl = string.Format("{0}/{1}", web.ServerRelativeUrl, "Lists/Employees/editform.aspx");
                CloseAllWebParts(ctx, editFormUrl);
                ProvisionWebPart(ctx, editFormUrl, "Emp-EditForm.js");
            }
            else
            {
                Console.WriteLine("The Edit form page was already customized");
            }
        }

        private static void ProvisionWebPart(ClientContext ctx, string relativePageUrl, string scriptFile)
        {
            Console.WriteLine("Provisioning web part...");

            string webPartXml = System.IO.File.ReadAllText(@"Assets\EmployeeRegistration.dwp");
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
            ctx.Web.AddWebPartToWebPartPage(relativePageUrl, webPart);
            Console.WriteLine("");
        }

        private static bool IsWebPartOnPage(ClientContext ctx, string relativePageUrl, string title)
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

        private static void CloseAllWebParts(ClientContext ctx, string relativePageUrl)
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

        private static void ProvisionAssets(ClientContext ctx)
        {
            Console.WriteLine("Provisioning assets:");

            string[] fileNames = {"jquery-1.10.2.min.js",
                                  "knockout-3.3.0.js",
                                  "Emp-EditForm.js",
                                  "Emp-NewForm.js",
                                  "Emp-Registration-Form-Template.js",
                                  "Emp-Registration-Form.js"};

            List styleLibrary = ctx.Web.Lists.GetByTitle("Style Library");
            ctx.Load(styleLibrary, l => l.RootFolder);
            Folder pnpFolder = styleLibrary.RootFolder.EnsureFolder("OfficeDevPnP");
            foreach (string fileName in fileNames)
            {
                Console.WriteLine(fileName);

                File assetFile = pnpFolder.GetFile(fileName);
                if (assetFile != null)
                    assetFile.CheckOut();

                string localFilePath = "Assets/" + fileName;
                string newLocalFilePath = Utilities.ReplaceTokensInAssetFile(ctx, localFilePath);

                assetFile = pnpFolder.UploadFile(fileName, newLocalFilePath, true);
                assetFile.CheckIn("Uploaded by provisioning engine.", CheckinType.MajorCheckIn);
                ctx.ExecuteQuery();
                System.IO.File.Delete(newLocalFilePath);
            }
            Console.WriteLine("");
        }

        private static void ProvisionLists(ClientContext ctx)
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


        private static ClientContext CreateContext()
        {
            ClientContext ctx = new ClientContext(siteUrl);
            ctx.Credentials = new SharePointOnlineCredentials(username, password);
            ctx.ExecuteQuery();

            Console.WriteLine("Connected to {0}", siteUrl);
            Console.WriteLine("");

            return ctx;
        }

        private static SecureString GetPassword()
        {
            SecureString sStrPwd = new SecureString();
            try
            {
                Console.Write("SharePoint Password : ");

                for (ConsoleKeyInfo keyInfo = Console.ReadKey(true); keyInfo.Key != ConsoleKey.Enter; keyInfo = Console.ReadKey(true))
                {
                    if (keyInfo.Key == ConsoleKey.Backspace)
                    {
                        if (sStrPwd.Length > 0)
                        {
                            sStrPwd.RemoveAt(sStrPwd.Length - 1);
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                            Console.Write(" ");
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        }
                    }
                    else if (keyInfo.Key != ConsoleKey.Enter)
                    {
                        Console.Write("*");
                        sStrPwd.AppendChar(keyInfo.KeyChar);
                    }

                }
                Console.WriteLine("");
            }
            catch (Exception e)
            {
                sStrPwd = null;
                Console.WriteLine(e.Message);
            }

            return sStrPwd;
        }

        private static string GetUserName()
        {
            string strUserName = string.Empty;
            try
            {
                Console.Write("SharePoint User Name : ");
                strUserName = Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                strUserName = string.Empty;
            }
            return strUserName;
        }

        private static string GetSite()
        {
            string siteUrl = string.Empty;
            try
            {
                Console.Write("Give Office365 site URL : ");
                siteUrl = Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                siteUrl = string.Empty;
            }
            return siteUrl;
        }

        #endregion


    }
}
