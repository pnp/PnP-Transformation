using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.WebPart
{
    [Cmdlet(VerbsCommon.Get, "WebPartPropertiesByUsageCSV")]
    public class GetWebPartPropertiesByUsageCSV : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string SourceWebPartType;

        [Parameter(Mandatory = true, Position = 1)]
        public string WebPartUsageFilePath;

        [Parameter(Mandatory = true, Position = 2)]
        public string OutPutDirectory;

        [Parameter(Mandatory = true, Position = 3)]
        public string SharePointOnline_OR_OnPremise;

        [Parameter(Mandatory = true, Position = 4)]
        public string UserName;

        [Parameter(Mandatory = true, Position = 5)]
        public string Password;

        [Parameter(Mandatory = true, Position = 6)]
        public string Domain;

        protected override void ProcessRecord()
        {
            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();
            webPartTransformationHelper.GetWebPartProperties_UsingCSV(SourceWebPartType, WebPartUsageFilePath, OutPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain, "CSVUpdates");
        }
    }
}
