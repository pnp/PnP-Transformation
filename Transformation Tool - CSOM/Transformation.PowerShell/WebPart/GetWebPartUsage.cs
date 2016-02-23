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
    [Cmdlet(VerbsCommon.Get, "WebPartUsage")]
    public class GetWebPartUsage : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string SiteUrl;

        [Parameter(Mandatory = true, Position = 1)]
        public bool ExpandSubSites;

        [Parameter(Mandatory = true, Position = 2)]        
        public string WebPartType;

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
            WebPartTransformationHelper webPartTransformationHelper=new WebPartTransformationHelper("GetWebPartUsage");
            if (SharePointOnline_OR_OnPremise.ToUpper().Equals("OP"))
            {
                webPartTransformationHelper.UseNetworkCredentialsAuthentication(UserName, Password, Domain);
            }
            else if (SharePointOnline_OR_OnPremise.ToUpper().Equals("OL"))
            {
                webPartTransformationHelper.UseOffice365Authentication(UserName, Password);
            }
            
            //Deleted the Web Part Usage File
            webPartTransformationHelper.DeleteUsageFiles_WebPartHelper(OutPutDirectory, Constants.WEBPART_USAGE_ENTITY_FILENAME);

            webPartTransformationHelper.AddSite(SiteUrl);
            webPartTransformationHelper.WebPartType = WebPartType;
            webPartTransformationHelper.OutPutDirectory = OutPutDirectory;
            webPartTransformationHelper.ExpandSubSites = ExpandSubSites;
            webPartTransformationHelper.headerWebPart = false;
            webPartTransformationHelper.Run();
        }
    }
}
