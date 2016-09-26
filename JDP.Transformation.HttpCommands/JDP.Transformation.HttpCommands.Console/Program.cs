using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDP.Transformation.HttpCommands
{
    class Program
    {
        static void Main(string[] args)
        {
            string siteUrl = "https://<tenant>.sharepoint.com/sites/dev";
            string username = "<user>@<tenant>.onmicrosoft.com";
            string password = "<password>";
            string domain = "<domain>";
            string[] solutionNames = { "sample.wsp" };

            // NOTE:
            // sample below is using Office365 authentication
            // for windows authentication use:
            // RequestActivateSandboxSolution activateSandbox = new RequestActivateSandboxSolution(siteUrl, 
            //      AuthenticationType.NetworkCredentials, username, password, domain);

            foreach (string solutionName in solutionNames)
            {
                Console.WriteLine("Activating solution: " + solutionName);
                var activateSandbox = new RequestActivateSandboxSolution(siteUrl,
                    AuthenticationType.Office365, username, password);
                activateSandbox.SolutionName = solutionName;
                activateSandbox.Execute();
                Console.WriteLine("Done.");
            }


            /*foreach (string solutionName in solutionNames)
            {
                Console.WriteLine("Deactivating solution: " + solutionName);
                var deactivateSandbox = new RequestDeactivateSandboxSolution(siteUrl,
                    AuthenticationType.Office365, username, password);
                deactivateSandbox.SolutionName = solutionName;
                deactivateSandbox.Execute();
                Console.WriteLine("Done.");
            }*/

            Console.ReadLine();
        }
    }
}
