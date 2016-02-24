using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Management.Automation;
using Transformation.PowerShell.Base;
using Transformation.PowerShell.Common;

namespace Transformation.PowerShell.MasterPage
{
    [Cmdlet(VerbsCommon.Set, "MasterPageWebLevel")]
   
    public class UpdateMasterPageWebLevel : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string OutPutDirectory;
        [Parameter(Mandatory = true, Position = 1)]
        public string WebUrl;
        [Parameter(Mandatory = true, Position = 2)]
        public string New_MasterPageURL;
        [Parameter(Position = 3)]
        public string Old_MasterPageURL;
        [Parameter(Position = 4)]
        public bool CustomMasterUrlStatus;
        [Parameter(Position = 5)]
        public bool MasterUrlStatus;
        [Parameter(Mandatory = true, Position = 6)]
        public string SharePointOnline_OR_OnPremise;
        [Parameter(Mandatory = true, Position = 7)]
        public string UserName;
        [Parameter(Mandatory = true, Position = 8)]
        public string Password;
        [Parameter(Mandatory = true, Position = 9)]
        public string Domain;
        
        protected override void ProcessRecord()
        {
            MasterPageHelper objMasterHelper = new MasterPageHelper();
            objMasterHelper.ChangeMasterPageForWeb(OutPutDirectory, WebUrl, New_MasterPageURL, Old_MasterPageURL, CustomMasterUrlStatus, MasterUrlStatus, Constants.ActionType_Web.ToLower(), SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }
    }
}
