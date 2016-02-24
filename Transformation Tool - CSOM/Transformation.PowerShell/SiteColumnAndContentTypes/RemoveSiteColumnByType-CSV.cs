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
    [Cmdlet(VerbsCommon.Remove, "SiteColumnByTypeUsingCSV")]
    public class RemoveSiteColumnByType_CSV : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0, HelpMessage = "Provide output directory path. Log file will be created in output directory")]
        public string OutPutDirectory;
        
        [Parameter(Mandatory = true, Position = 1, HelpMessage = "Field type of the site column. Ex: Text, Note, Integer, Choice etc.,")]
        public string FieldType;
        
        [Parameter(Mandatory = true, Position = 2)]
        public string SiteColumnUsageFilePath;

        [Parameter(Mandatory = true, Position = 3)]
        public string SharePointOnline_OR_OnPremise;

        [Parameter(Mandatory = true, Position = 4)]
        public string UserName;

        [Parameter(Mandatory = true, Position = 5)]
        public string Password;

        [Parameter(Mandatory = true, Position = 6)]
        public string Domain;

        [Parameter(Position = 7, HelpMessage = "-Confirm is optional parameter. Site column will be removed from web, if you specify -Confirm option. Otherwise, it shows usage report of site column")]
        public SwitchParameter Confirm;

        protected override void ProcessRecord()
        {
            SiteColumnAndContentTypeHelper obj = new SiteColumnAndContentTypeHelper();
            obj.RemoveSiteColumnByTypeUsingCSV(OutPutDirectory.Trim(), FieldType, SiteColumnUsageFilePath, SharePointOnline_OR_OnPremise.Trim()
                                                , UserName.Trim(), Password.Trim(), Domain.Trim(), Confirm);
        }
    }
}

