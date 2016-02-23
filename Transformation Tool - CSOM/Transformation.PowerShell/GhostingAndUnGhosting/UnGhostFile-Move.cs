using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;
using Transformation.PowerShell.Common;

namespace Transformation.PowerShell.GhostingAndUnGhosting
{
    [Cmdlet(VerbsCommon.Move, "UnGhostFile")]
    public class UnGhostFile_Move : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string AbsoluteFilePath;
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
            GhostingAndUnGhostingHelper obj = new GhostingAndUnGhostingHelper();
            obj.UnGhostFile(AbsoluteFilePath.Trim(), OutPutDirectory.Trim(), Constants.UnGhostFileOperation_Move, SharePointOnline_OR_OnPremise.Trim(), UserName.Trim(), Password.Trim(), Domain.Trim());
        }
    }
}
