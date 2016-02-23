using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.MasterPage
{
    public class MasterPageBase : Elementbase
    {
        public string MasterUrl { get; set; }
        public string OLD_MasterUrl { get; set; }
        public string CustomMasterUrl { get; set; }
        public string OLD_CustomMasterUrl { get; set; }
    }

    public class MasterPageInput : Inputbase
    {
        public string MasterUrl { get; set; }
        public string MasterUrlStatus { get; set; }
        public string CustomMasterUrl { get; set; }
        public string CustomMasterUrlStatus { get; set; }
    }
}