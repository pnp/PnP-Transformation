using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Transformation.PowerShell.Base;
using Transformation.PowerShell.Common.CSV;

namespace Transformation.PowerShell.WebPart
{
    [Cmdlet(VerbsCommon.Remove, "WebPartsByUsageCSV")]
    public class DeleteWebPartByUsageCSV : TrasnformationPowerShellCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string SourceWebPartType;

        [Parameter(Mandatory = true, Position = 1)]
        public string WebPartUsageFilePath;

        [Parameter(Mandatory = true, Position = 2)]
        public string OutPutDirectory;

        [Parameter(Mandatory = true, Position = 3)]
        public string SharePointOnline_OR_OnPremise;

        [Parameter(Mandatory = true, Position = 4)]
        public string UserName;

        [Parameter(Mandatory = true, Position = 5)]
        public string Password;

        [Parameter(Mandatory = true, Position = 6)]
        public string Domain;

        protected override void ProcessRecord()
        {
            WebPartTransformationHelper webPartTransformationHelper = new WebPartTransformationHelper();
            string csvFile = OutPutDirectory + @"\" + System.IO.Path.GetFileNameWithoutExtension(WebPartUsageFilePath) + "_DeleteOperationStatus" + "_" + DateTime.Now.ToString("dd_MM_yyyy_hh_ss") + ".csv";


            if (String.Equals(Transformation.PowerShell.Common.Constants.ActionType_All, SourceWebPartType, StringComparison.CurrentCultureIgnoreCase))
            {
                //Reading Input File
                IEnumerable<WebPartDiscoveryInput> objWPDInput;

                objWPDInput = ImportCsv.ReadMatchingColumns<WebPartDiscoveryInput>(WebPartUsageFilePath, Transformation.PowerShell.Common.Constants.CsvDelimeter);

                if (objWPDInput.Any())
                {
                    IEnumerable <string> webPartTypes = objWPDInput.Select(x => x.WebPartType);

                    webPartTypes = webPartTypes.Distinct();

                    foreach (string webPartType in webPartTypes)
                    {
                        webPartTransformationHelper.DeleteWebPart_UsingCSV(webPartType, WebPartUsageFilePath, OutPutDirectory, csvFile, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
                    }
                }
            }
            else
            {
                webPartTransformationHelper.DeleteWebPart_UsingCSV(SourceWebPartType, WebPartUsageFilePath, OutPutDirectory, csvFile, SharePointOnline_OR_OnPremise, UserName, Password, Domain);
            }

        }
    }
}
