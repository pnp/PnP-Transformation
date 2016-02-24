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
    [Cmdlet(VerbsCommon.Set, "ReplaceContentTypeInListWebLevel")]
    public class ReplaceContentTypeInList_Web : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string OutPutDirectory;
        [Parameter(Mandatory = true, Position = 1)]
        public string WebUrl;
        [Parameter(Mandatory = true, Position = 2)]
        public string ListName;
        [Parameter(Mandatory = true, Position = 3)]
        public string OldContentTypeID;
        [Parameter(Mandatory = true, Position = 4)]
        public string NewContentTypeName;
        [Parameter(Mandatory = true, Position = 5)]
        public string SharePointOnline_OR_OnPremise;
        [Parameter(Mandatory = true, Position = 6)]
        public string UserName;
        [Parameter(Mandatory = true, Position = 7)]
        public string Password;
        [Parameter(Mandatory = true, Position = 8)]
        public string Domain;
        
        protected override void ProcessRecord()
        {
            SiteColumnAndContentTypeHelper obj = new SiteColumnAndContentTypeHelper();
            obj.ReplaceContentTypeinList_ForWeb(OutPutDirectory, WebUrl, ListName, OldContentTypeID, NewContentTypeName, Constants.ActionType_Web.ToLower(), SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }
    }
}
