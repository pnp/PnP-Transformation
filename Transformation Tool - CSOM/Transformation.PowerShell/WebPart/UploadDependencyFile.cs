using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.WebPart
{
    [Cmdlet(VerbsCommon.Move, "DependencyFileToServerRelativeFolder")]
    public class UploadDependencyFile : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string  WebUrl;

        [Parameter(Mandatory = true, Position = 1)]
        public string FolderServerRelativeUrl;  
      
        [Parameter(Mandatory = true, Position = 2)]
        public string FileName;

        [Parameter(Mandatory = true, Position = 3)]
        public string LocalFilePath; 

        [Parameter(Mandatory = true, Position = 4)]
        public bool OverwriteIfExists;

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
            WebPartTransformationHelper webPartTransformationHelper=new WebPartTransformationHelper();
            webPartTransformationHelper.UploadDependencyFile(WebUrl, FolderServerRelativeUrl, FileName, LocalFilePath, OverwriteIfExists, OutPutDirectory, SharePointOnline_OR_OnPremise, UserName, Password, Domain);            
        }

    }
}
