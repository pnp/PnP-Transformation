using JDP.Remediation.Console.Common.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace JDP.Remediation.Console
{
    public static class Program
    {
        /// <summary>
        /// Gets or sets the domain provided by user
        /// </summary>
        public static string AdminDomain
        {
            get;
            set;
        }
        /// <summary>
        /// Gets or sets the username provided by user
        /// </summary>
        public static string AdminUsername
        {
            get;
            set;
        }
        /// <summary>
        /// Gets or sets the password provided by user
        /// </summary>
        public static SecureString AdminPassword
        {
            get;
            set;
        }

        private static void ShowUsage()
        {
            System.Console.WriteLine();
            System.Console.WriteLine();
            System.Console.ForegroundColor = System.ConsoleColor.White;
            System.Console.WriteLine("#### JDP Remediation Console ####");
            System.Console.ResetColor();
            System.Console.ForegroundColor = System.ConsoleColor.Yellow;
            System.Console.WriteLine("Please type an operation number and press [Enter] to execute the specified operation:");
            System.Console.WriteLine("1. Transformation");
            System.Console.WriteLine("2. Clean-Up");
            System.Console.WriteLine("3. Self-Service Reports");
            System.Console.WriteLine("4. Exit");
            System.Console.ResetColor();
            System.Console.WriteLine();
        }

        /// <summary>
        /// Main method
        /// </summary>
        /// <param name="args">Event Arguments</param>
        public static void Main(string[] args)
        {
            string input = String.Empty;
            string input1 = string.Empty;
            string input2 = string.Empty;

            GetCredentials();
            //Excception CSV Creation Command
            ExceptionCsv objException = ExceptionCsv.CurrentInstance;
            objException.CreateLogFile(Environment.CurrentDirectory);

            do
            {
                ShowUsage();
                input = System.Console.ReadLine();
                switch (input.ToUpper(System.Globalization.CultureInfo.CurrentCulture))
                {
                    case "1":
                        do
                        {
                            System.Console.ForegroundColor = System.ConsoleColor.Green;
                            System.Console.WriteLine("");
                            System.Console.WriteLine("Your selected Operation [1: \"Transformation\"] supports the following operations.");
                            System.Console.WriteLine("Please type an operation number and press [Enter] to execute the specified operation:");
                            System.Console.WriteLine("1. Add OOTB Web Part or App Part to a page");
                            System.Console.WriteLine("2. Replace FTC Web Part with OOTB Web Part or App Part on a page");
                            System.Console.WriteLine("3. Replace MasterPage");
                            System.Console.WriteLine("4. Reset Device Channel Mapping Files");
                            System.Console.WriteLine("5. Manage [Add|Remove] Maintenance Banners");
                            System.Console.WriteLine("6. Exit");
                            System.Console.ResetColor();
                            input1 = System.Console.ReadLine();
                            switch (input1)
                            {
                                case "1":
                                    AddWebPart.DoWork();
                                    break;
                                case "2":
                                    ReplaceWebPart.DoWork();
                                    break;
                                case "3":
                                    ReplaceMasterPage.DoWork();
                                    break;
                                case "4":
                                    ResetDeviceChannelMappingFiles.DoWork();
                                    break;
                                case "5":
                                    ManageMaintenanceBanners.DoWork();
                                    break;
                                case "6":
                                    break;
                                default:
                                    break;
                            }
                        } while (input1.ToUpper(System.Globalization.CultureInfo.CurrentCulture) != "6");
                        break;

                    case "2":
                        do
                        {
                            System.Console.ForegroundColor = System.ConsoleColor.Magenta;
                            System.Console.WriteLine("");
                            System.Console.WriteLine("Your selected Operation [2: \"Clean-Up\"] supports the following operations.");
                            System.Console.WriteLine("Please type an operation number and press [Enter] to execute the specified operation:");
                            System.Console.WriteLine("1. Delete Missing Setup Files");
                            System.Console.WriteLine("2. Delete Missing Features");
                            System.Console.WriteLine("3. Delete Missing Event Receivers");
                            System.Console.WriteLine("4. Delete Workflow Associations");
                            System.Console.WriteLine("5. Delete List Templates");
                            System.Console.WriteLine("6. Delete Missing Webparts");
                            System.Console.WriteLine("7. Exit");
                            System.Console.ResetColor();
                            input1 = System.Console.ReadLine();
                            switch (input1)
                            {
                                case "1":
                                    DeleteMissingSetupFiles.DoWork();
                                    break;
                                case "2":
                                    DeleteMissingFeatures.DoWork();
                                    break;
                                case "3":
                                    DeleteMissingEventReceivers.DoWork();
                                    break;
                                case "4":
                                    DeleteMissingWorkflowAssociations.DoWork();
                                    break;
                                case "5":
                                    DownloadAndModifyListTemplate.DeleteListTemplate();
                                    break;
                                case "6":
                                    DeleteWebparts.DoWork();
                                    break;
                                case "7":
                                    break;
                                default:
                                    break;
                            }
                        } while (input1.ToUpper(System.Globalization.CultureInfo.CurrentCulture) != "7");
                        break;

                    case "3":
                        do
                        {
                            System.Console.ForegroundColor = System.ConsoleColor.DarkCyan;
                            System.Console.WriteLine("");
                            System.Console.WriteLine("Your selected Operation [3: \"Self-Service Reports\"] supports the following operations.");
                            System.Console.WriteLine("Please type an operation number and press [Enter] to execute the specified operation:");
                            System.Console.WriteLine("1. Generate List Template Report with FTC Analysis");
                            System.Console.WriteLine("2. Generate Site Template Report with FTC Analysis");
                            System.Console.WriteLine("3. Generate Site Column/Custom Field & Content Type Usage Report");
                            System.Console.WriteLine("4. Generate Non-Default Master Page Usage Report");
                            System.Console.WriteLine("5. Generate Site Collection Report (PPE-Only)");
                            System.Console.WriteLine("6. Generate Web Part Usage Report");
                            System.Console.WriteLine("7. Generate Web Part Properties Report");
                            System.Console.WriteLine("8. Generate Security Group Report");
                            System.Console.WriteLine("9. Generate Pivot Reports ");
                            System.Console.WriteLine("10. Exit ");
                            System.Console.ResetColor();
                            input1 = System.Console.ReadLine();
                            switch (input1)
                            {
                                case "1":
                                    DownloadAndModifyListTemplate.DoWork();
                                    break;
                                case "2":
                                    DownloadAndModifySiteTemplate.DoWork();
                                    break;
                                case "3":
                                    GenerateColumnAndTypeUsageReport.DoWork();
                                    break;
                                case "4":
                                    GenerateNonDefaultMasterPageUsageReport.DoWork();
                                    break;
                                case "5":
                                    System.Console.ForegroundColor = ConsoleColor.Yellow;
                                    System.Console.WriteLine("This operation is intended for use only in PPE; use on PROD at your own risk.");
                                    System.Console.WriteLine("For PROD, it is safer to generate the report via the o365 Self-Service Admin Portal.");
                                    System.Console.ResetColor();
                                    System.Console.ForegroundColor = System.ConsoleColor.Cyan; 
                                    System.Console.WriteLine("Press \"y\" only if you wish to continue.  Press any other key to abort this operation.");
                                    System.Console.ResetColor();

                                    input2 = System.Console.ReadLine();
                                    if (input2.ToUpper(System.Globalization.CultureInfo.CurrentCulture) != "Y")
                                    {
                                        System.Console.WriteLine("Operation aborted by user.");
                                        break;
                                    }
                                    GenerateSiteCollectionReport.DoWork();
                                    break;
                                case "6":
                                    WebPartUsage.DoWork();
                                    break;
                                case "7":
                                    WebPartProperties.DoWork();
                                    break;
                                case "8":
                                    GenerateSecurityGroupReport.DoWork();
                                    break;
                                case "9":
                                    GeneratePivotReport.DoWork();
                                    break;
                                case "10":
                                    break;
                                default:
                                    break;
                            }
                        } while (input1.ToUpper(System.Globalization.CultureInfo.CurrentCulture) != "10");
                        break;

                    case "4":
                        break;

                    default:
                        break;
                }
            }
            while (input.ToUpper(System.Globalization.CultureInfo.CurrentCulture) != "4");
        }

        /// <summary>
        /// get credentials
        /// </summary>
        public static void GetCredentials()
        {
            ConsoleKeyInfo key;
            bool retryUserNameInput = false;
            string account = String.Empty;
            string password = String.Empty;

            do
            {
                System.Console.ForegroundColor = System.ConsoleColor.Cyan;
                System.Console.WriteLine(@"Please enter the Admin account: ");
                System.Console.ForegroundColor = System.ConsoleColor.Yellow;
                System.Console.WriteLine(@"- Use [domain\alias] format for SPO-D & On-Prem farms");
                System.Console.WriteLine(@"- Use [alias@domain.com] format for SPO-MT & vNext farms");
                System.Console.ResetColor();

                account = System.Console.ReadLine();

                if (account.Contains('\\'))
                {
                    string[] segments = account.Split('\\');
                    AdminDomain = segments[0];
                    AdminUsername = segments[1];
                    break;
                }
                if (account.Contains("@"))
                {
                    AdminUsername = account;
                    break;
                }
            }
            while (retryUserNameInput);

            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            System.Console.WriteLine("Please enter the Admin password: ");
            System.Console.ResetColor();
            
            do
            {
                key = System.Console.ReadKey(true);

                if (key.Key != ConsoleKey.Backspace)
                {
                    password += key.KeyChar;
                    System.Console.Write("*");
                }
                else if (key.Key == ConsoleKey.Backspace)
                {
                    if (password.Length > 0)
                    {
                        password = password.Substring(0, password.Length - 1);
                        System.Console.CursorLeft--;
                        System.Console.Write('\0');
                        System.Console.CursorLeft--;
                    }
                }
            }
            while (key.Key != ConsoleKey.Enter);

            System.Console.WriteLine("");

            AdminPassword = Helper.CreateSecureString(password.TrimEnd('\r'));
        }
    }
}
