using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace PeoplePickerRemediation.Console
{
    class Program
    {
        /// <summary>
        // Return a value of TRUE if you want the utility to leverage the App Model for AuthN/AuthZ
        // Return a value of FALSE if you want the utility to leverage User Credentials for AuthN/AuthZ
        /// </summary>
        public static bool UseAppModel
        {
            get
            {
                try
                {
                    return System.Configuration.ConfigurationManager.AppSettings["UseAppModel"].ToBoolean();
                }
                catch
                {
                    return true;
                }
            }
        }

        /// <summary>
        /// Gets or sets the domain provided by user
        /// </summary>
        ///
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
        public static void Main(string[] args)
        {
            if (UseAppModel == false)
            {
                GetCredentials();
            }

            string inputFilePath = GetInputFilePath();

            if (PeoplePickerRemediation.ValidateLDAPVariable())
            {
                if (ValidateDirectorySearcherUser())
                {
                    PeoplePickerRemediation.DoWork(inputFilePath);
                }

                System.Console.WriteLine("Execution has completed");
            }
        }

        private static bool ValidateDirectorySearcherUser()
        {
            bool continueStatus = false;
            string accountID = "";
            System.Console.WriteLine("LDAP Directory Lookup Validation:");
            System.Console.WriteLine(String.Format("-AppSettings[LocalAdLdapQuery] = {0}", ConfigurationManager.AppSettings["LocalAdLdapQuery"].ToString()));
            System.Console.WriteLine("Please provide a user account to test [format = domain\\alias]: ");
            accountID = System.Console.ReadLine();
            string upn = PeoplePickerRemediation.GetUserPrinicpalNameFromDirectorySearcher(accountID);
            System.Console.WriteLine(string.Format("{0} user's UPN value is {1}", accountID, String.IsNullOrEmpty(upn) ? "Empty" : upn));
            System.Console.WriteLine("");
            System.Console.WriteLine(string.Format("Do you want to continue and process the InfoPath Form Libraries? [Y/N]?"));
            string continuedStatus = System.Console.ReadLine();
            if (!string.IsNullOrEmpty(upn) && (continuedStatus.Contains("y") || continuedStatus.Contains("Y")))
            {
                continueStatus = true;
            }
            System.Console.WriteLine("");
            System.Console.WriteLine("");
            return continueStatus;
        }

        private static string GetInputFilePath()
        {
            bool retryFilePathInput;
            string inputFilePath;

            do
            {
                retryFilePathInput = false;
                System.Console.WriteLine(@"Please enter the Path containing the InfoPath Form Libraries List: ");
                System.Console.WriteLine(@"- Give the path in the following format [Folder path containing the InfoPath Form Libraries List]\[CSV File Name]");
                System.Console.WriteLine(@"- Example: E:\PeoplePickerRemediation\PeoplePickerReport.csv");

                inputFilePath = System.Console.ReadLine();

                if (inputFilePath == "")
                {
                    retryFilePathInput = true;
                    System.Console.WriteLine(@"Please make sure the File Path is not empty");
                }
                else
                {
                    if (!File.Exists(inputFilePath))
                    {
                        retryFilePathInput = true;
                        System.Console.WriteLine("");
                        System.Console.WriteLine(@"Please make sure the File Path is in a valid format");
                    }
                }
            }
            while (retryFilePathInput);

            System.Console.WriteLine("");
            return inputFilePath;
        }

        private static void GetCredentials()
        {
            ConsoleKeyInfo key;
            bool retryUserNameInput = false;
            string account = String.Empty;
            string password = String.Empty;

            do
            {
                System.Console.WriteLine(@"Please enter the Admin account: ");
                System.Console.WriteLine(@"- Use [domain\alias] for SPO-D & On-Prem farms");
                System.Console.WriteLine(@"- Use [alias@domain.com] for SPO-MT & vNext farms");

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

            System.Console.WriteLine("Please enter the Admin password: ");

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

            AdminPassword = Helper.CreateSecureString(password.TrimEnd('\r'));
            System.Console.WriteLine("");
            System.Console.WriteLine("");
        }
    }
}
