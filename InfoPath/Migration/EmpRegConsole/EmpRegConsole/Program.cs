using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml;
using System.Web;
using System.Xml.XPath;

namespace EmpRegConsole
{
    class Program
    {
        static string sharePointSiteUrl = string.Empty;
        static string infoPathLibName = string.Empty;
        static string listName = string.Empty;

        static void Main(string[] args)
        {
            GetInputVariables();
            GetInfoPathAndStoreDataInList();
            Console.WriteLine("--------------------------------------------");
            Console.WriteLine("Successfully added InfoPath data to list");
            Console.WriteLine("--------------------------------------------");
        }

        private static void GetInputVariables()
        {
            bool isValidInputs;

            do
            {
                isValidInputs = false;

                Console.WriteLine("Enter sharepoint site url:");
                sharePointSiteUrl = Console.ReadLine();
                Console.WriteLine("Enter source infopath library name or title:");
                infoPathLibName = Console.ReadLine();
                Console.WriteLine("Enter target sharepoint list name or title:");
                listName = Console.ReadLine();

                Console.WriteLine("------------------------------------------");
                Console.WriteLine("Please confirm above inputs (y/n)?");
                Console.WriteLine("------------------------------------------");
                ConsoleKeyInfo confirmKey = Console.ReadKey(true);

                if (confirmKey.Key == ConsoleKey.N)
                {
                    Console.WriteLine("-----------------------------------------------------------------");
                    Console.WriteLine("Press 'E' to enter inputs again or Press 'X' to exit application?");
                    Console.WriteLine("-----------------------------------------------------------------");
                    confirmKey = Console.ReadKey(true);

                    if (confirmKey.Key == ConsoleKey.X)
                    {
                        Environment.Exit(0); // 0 - Success 
                    }

                    isValidInputs = (confirmKey.Key == ConsoleKey.E);
                }

            } while (isValidInputs);


        }

        private static void GetInfoPathAndStoreDataInList()
        {
            try
            {
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

                } // using (ClientContext clientContext
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error Message: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        } // GetInfoPathAndStoreDataInList

        private static void AddInfoPathToList(ClientContext clientContext, Web web, Employees emp)
        {
            Console.WriteLine("Adding infopath data to list....");
            List lstEmployees = web.Lists.GetByTitle(listName); // EmployeesConsoleTest
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem listItem = lstEmployees.AddItem(itemCreateInfo);
            listItem["Title"] = emp.Name;
            listItem["UserID"] = emp.UserID;
            listItem["EmpManager"] = emp.Manager;
            listItem["EmpNumber"] = emp.Number;
            listItem["Designation"] = emp.Designation;
            listItem["Location"] = emp.Locaiton;
            listItem["Skills"] = emp.Skills;

            listItem.Update();
            clientContext.ExecuteQuery();
        } // AddInfoPathToList

        private static void ReadInfoPathFile(ClientContext clientContext, Web web, ListItem item, out Employees employee)
        {
            Console.WriteLine("Reading InfoPath file " + item["Title"]);
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
        } // ReadInfoPathFile
    }
}
