using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.SiteColumnAndContentTypes
{
    [Cmdlet(VerbsCommon.Add, "SiteColumnUsingCSV")]
    public class CreateSiteColumn_CSV : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string OldSiteColumn_InternalName;
        [Parameter(Position = 1)]
        public string OldSiteColumn_ID;
        [Parameter(Mandatory = true, Position = 2)]
        public string NewSiteColumn_InternalName;
        [Parameter(Mandatory = true, Position = 3)]
        public string NewSiteColumn_DisplayName;
        [Parameter(Mandatory = true, Position = 4)]
        public string SiteColumnUsageFilePath;
        [Parameter(Mandatory = true, Position = 5)]
        public string OutPutDirectory;
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
            SiteColumnAndContentTypeHelper obj = new SiteColumnAndContentTypeHelper();
            obj.SiteColumns_CreateSiteColumnsAndDuplicateDetailsFromOldSiteColumn_UsingCSV(OldSiteColumn_InternalName.Trim(), OldSiteColumn_ID, NewSiteColumn_InternalName.Trim(), NewSiteColumn_DisplayName.Trim(), SiteColumnUsageFilePath.Trim(), OutPutDirectory.Trim(), SharePointOnline_OR_OnPremise.Trim(), UserName.Trim(), Password.Trim(), Domain.Trim());
        }
    }
}
