using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UdcxRemediation.Console.Common.Base
{
    public class Elementbase
    {
        public string SiteUrl { get; set; }

        public string WebUrl { get; set; }
    }

    public class UdcxReportInput : Elementbase
    {
        public string DirName { get; set; }
        public string LeafName { get; set; }
        public string Authentication { get; set; }       
    }

    public class UdcxReportOutput : UdcxReportInput
    {
        public string Status { get; set; }
        public string ErrorDetails { get; set; }
    }

}
