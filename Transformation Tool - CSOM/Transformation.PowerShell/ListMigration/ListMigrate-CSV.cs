using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;
using Transformation.PowerShell.Common;

namespace Transformation.PowerShell.ListMigration
{
    [Cmdlet(VerbsCommon.Add, "ListMigrateUsingCSV")]
    public class ListMigrate_CSV : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string OutPutDirectory;
        [Parameter(Mandatory = true, Position = 1)]
        public string ListUsageFilePath;
        [Parameter(Mandatory = true, Position = 2)]
        public string OldList_Title;
        [Parameter(Position = 3)]
        public string OldList_ID;
        [Parameter(Mandatory = true, Position = 4)]
        public string NewList_Title;
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
            ListMigrationHelper obj = new ListMigrationHelper();
            obj.ListMigration_UsingCSV(OldList_Title.Trim(), OldList_ID, NewList_Title.Trim(), ListUsageFilePath.ToString().Trim(), OutPutDirectory.Trim(), SharePointOnline_OR_OnPremise.Trim(), UserName.Trim(), Password.Trim(), Domain.Trim());
        }
    }
}
