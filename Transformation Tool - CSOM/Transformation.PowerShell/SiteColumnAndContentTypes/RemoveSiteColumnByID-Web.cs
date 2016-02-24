using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;
using Transformation.PowerShell.Common;

namespace Transformation.PowerShell.SiteColumnAndContentTypes
{
    [Cmdlet(VerbsCommon.Remove, "SiteColumnByID")]
    public class RemoveSiteColumnByID_Web : TrasnformationPowerShellCmdlet
    {

        [Parameter(Mandatory = true, Position = 0, HelpMessage = "Provide output directory path. Log file will be created in output directory")]
        public string OutPutDirectory;

        [Parameter(Mandatory = true, Position = 1, HelpMessage = "Enter SharePoint site collection url. Ex: https://sharepoint.com/[sites/]Web_Site")]
        public string SiteCollectionUrl;

        [Parameter(Mandatory = true, Position = 2, HelpMessage = "The unique id of the site column. Ex: 8478039d-fbd5-421d-bd6c-87a07d7ce499")]
        public string FieldID;

        [Parameter(Mandatory = true, Position = 3)]
        public string SharePointOnline_OR_OnPremise;

        [Parameter(Mandatory = true, Position = 4)]
        public string UserName;

        [Parameter(Mandatory = true, Position = 5)]
        public string Password;

        [Parameter(Mandatory = true, Position = 6)]
        public string Domain;

        [Parameter(Position = 7, HelpMessage = "-Confirm is optional parameter. Site column will be removed from web, if you specify -Confirm option. Otherwise, it shows usage report of site column")]
        public SwitchParameter Confirm;

        protected override void ProcessRecord()
        {
            SiteColumnAndContentTypeHelper obj = new SiteColumnAndContentTypeHelper();
            obj.IterateAllWebsAndRemoveSiteColumnByID(OutPutDirectory.Trim(), SiteCollectionUrl.Trim(), new Guid(FieldID)
                                                        , SharePointOnline_OR_OnPremise.Trim(), UserName.Trim()
                                                        , Password.Trim(), Domain.Trim(), Confirm);

            if (!Confirm)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("[Hint] Use -Confirm option to remove site column");
                Console.ResetColor();
            }
        } //  ProcessRecord()
    }
}
