using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.WebPart
{
    [Cmdlet(VerbsCommon.Set, "TargetWebPart")]
    public class ReplaceWebPart : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string WebUrl;        

        [Parameter(Mandatory = true, Position = 1)]
        public string SourceWebPartStorageKey;

        [Parameter(Mandatory = true, Position = 2)]
        public string WebPartZoneIndex;

        [Parameter(Mandatory = true, Position = 3)]
        public string WebPartZoneID;

        [Parameter(Mandatory = true, Position = 4)]
        public string ServerRelativePageUrl;

        [Parameter(Mandatory = true, Position = 5)]
        public string TargetWebPartFileName;

        [Parameter(Mandatory = true, Position = 6)]
        public string TargetWebPartXmlFile;

        [Parameter(Mandatory = true, Position = 7)]
        public string OutPutDirectory;

        [Parameter(Mandatory = true, Position = 8)]
        public string SharePointOnline_OR_OnPremise;

        [Parameter(Mandatory = true, Position = 9)]
        public string UserName;

        [Parameter(Mandatory = true, Position = 10)]
        public string Password;

        [Parameter(Mandatory = true, Position = 11)]
        public string Domain;
        protected override void ProcessRecord()
        {
            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();
            webPartTransformationHelper.ReplaceWebPart(WebUrl, TargetWebPartFileName, TargetWebPartXmlFile, SourceWebPartStorageKey, WebPartZoneIndex, WebPartZoneID, ServerRelativePageUrl, OutPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain, "web");

        }  
    }
}
