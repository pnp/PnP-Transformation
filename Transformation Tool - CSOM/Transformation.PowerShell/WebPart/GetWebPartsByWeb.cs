using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;
using Transformation.PowerShell.Common;

namespace Transformation.PowerShell.WebPart
{
    [Cmdlet(VerbsCommon.Get, "WebPartsByWeb")]
    public class GetWebPartsByWeb : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string WebUrl;

        [Parameter(Mandatory = true, Position = 1)]
        public string OutPutDirectory;

        [Parameter(Mandatory = true, Position = 2)]
        public string SharePointOnline_OR_OnPremise;

        [Parameter(Mandatory = true, Position = 3)]
        public string UserName;

        [Parameter(Mandatory = true, Position = 4)]
        public string Password;

        [Parameter(Mandatory = true, Position = 5)]
        public string Domain;

        protected override void ProcessRecord()
        {
            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();
            webPartTransformationHelper.GetWebPartsByWeb(WebUrl, OutPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
            
        }    
    }
}
