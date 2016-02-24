using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.SiteColumnAndContentTypes
{
    [Cmdlet(VerbsCommon.Set, "ReplaceContentTypeInListUsingCSV")]
    public class ReplaceContentTypeInList_CSV : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string OutPutDirectory;
        [Parameter(Mandatory = true, Position = 1)]
        public string SharePointOnline_OR_OnPremise;
        [Parameter(Mandatory = true, Position = 2)]
        public string UserName;
        [Parameter(Mandatory = true, Position = 3)]
        public string Password;
        [Parameter(Mandatory = true, Position = 4)]
        public string Domain;
        
        protected override void ProcessRecord()
        {
            SiteColumnAndContentTypeHelper obj = new SiteColumnAndContentTypeHelper();
            obj.ReplaceContentTypeinList_ForCSV(OutPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }
    }
}
