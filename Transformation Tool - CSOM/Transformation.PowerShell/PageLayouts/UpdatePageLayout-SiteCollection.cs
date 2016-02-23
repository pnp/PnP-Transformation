using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;
using Transformation.PowerShell.PageLayouts;
using Transformation.PowerShell.PageLayout;

namespace Transformation.PowerShell.PageLayouts
{
    [Cmdlet(VerbsCommon.Set, "PageLayoutSiteCollectionLevel")]
    public class UpdatePageLayout_SiteCollection : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string OutPutDirectory;
        [Parameter(Mandatory = true, Position = 1)]
        public string SiteCollectionUrl;
        [Parameter(Mandatory = true, Position = 2)]
        public string OldPageLayoutUrl;
        [Parameter(Mandatory = true, Position = 3)]
        public string NewPageLayoutUrl;
        [Parameter(Position = 4)]
        public string NewPageLayoutDescription;
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
            PageLayoutHelper obj = new PageLayoutHelper();
            obj.ChangePageLayoutForPagesUsingOldPageLayoutInSiteCollection(OutPutDirectory, SiteCollectionUrl, OldPageLayoutUrl, NewPageLayoutUrl, NewPageLayoutDescription, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
        }
    }    
}
