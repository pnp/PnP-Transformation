using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.WebPart
{
    [Cmdlet(VerbsCommon.New, "WebPartXmlConfiguration")]
    public class ConfigureNewWebPartXml : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string TargetWebPartXmlFilePath;

        [Parameter(Mandatory = true, Position = 1)]
        public string SourceXmlFilesDirectory;

        //[Parameter(Mandatory = true, Position = 2)]
        //public string TargetXmlFilesDirectory;

        [Parameter(Mandatory = true, Position = 2)]
        public string OutPutDirectory;
        protected override void ProcessRecord()
        {
            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();
            webPartTransformationHelper.ConfigureNewWebPartXml(TargetWebPartXmlFilePath, SourceXmlFilesDirectory, OutPutDirectory);            
        }
    }
}
