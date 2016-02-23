using System.Management.Automation;

namespace Transformation.PowerShell.Base
{
    public class TrasnformationPowerShellCmdlet : PSCmdlet
    {
        protected virtual void ExecuteCmdlet()
        {
        }

        protected override void ProcessRecord()
        {
            ExecuteCmdlet();
        }
    }
}