using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.SiteColumnAndContentTypes
{
    [Cmdlet(VerbsCommon.Add, "SiteColumnToContentTypeUsingCSV")]
    public class AddSiteColumnToContentTyp_CSV : TrasnformationPowerShellCmdlet
    {
      
        [Parameter(Mandatory = true, Position = 0)]
        public string ContentTypeName;
        [Parameter(Mandatory = true, Position = 1)]
        public string SiteColumnName;
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
            obj.AddSiteColumnToContentType_ForCSV(ContentTypeName,SiteColumnName,ContentTypeUsageFilePath, OutPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }
    }
}
