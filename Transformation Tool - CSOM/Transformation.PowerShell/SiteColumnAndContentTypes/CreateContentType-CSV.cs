using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.SiteColumnAndContentTypes
{
    [Cmdlet(VerbsCommon.Add, "ContentTypeUsingCSV")]
    public class CreateContentType_CSV : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string OldContentTypeName;
        [Parameter(Mandatory = true, Position = 1)]
        public string NewContentTypeName;
        [Parameter(Mandatory = true, Position = 2)]
        public string ContentTypeUsageFilePath;
        [Parameter(Mandatory = true, Position = 3)]
        public string OutPutDirectory;
        [Parameter(Mandatory = true, Position = 4)]
        public string SharePointOnline_OR_OnPremise;
        [Parameter(Mandatory = true, Position = 5)]
        public string UserName;
        [Parameter(Mandatory = true, Position = 6)]
        public string Password;
        [Parameter(Mandatory = true, Position = 7)]
        public string Domain;
        
        protected override void ProcessRecord()
        {
            SiteColumnAndContentTypeHelper obj = new SiteColumnAndContentTypeHelper();
            obj.ContentType_CreateContentTypeAndDuplicateDetailsFromOldContentType_ForCSV(OldContentTypeName.Trim(), NewContentTypeName.Trim(), ContentTypeUsageFilePath.Trim(), OutPutDirectory.Trim(), SharePointOnline_OR_OnPremise.Trim(), UserName.Trim(), Password.Trim(), Domain.Trim());
        }
    }
}
