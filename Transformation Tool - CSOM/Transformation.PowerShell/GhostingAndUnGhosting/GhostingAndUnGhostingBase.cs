using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.GhostingAndUnGhosting
{
    public class GhostingAndUnGhostingBase : Elementbase
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        
    }
    public class DownloadFileBase : Elementbase
    {
        public string GivenFilePath { get; set; }
        public string FileName { get; set; }
        public string DownloadedFilePath { get; set; }

    }
}
