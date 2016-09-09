using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDP.Remediation.Console.Common.Base
{
    public class Elementbase
    {
        public string ContentDatabase { get; set; }
        public string WebApplication { get; set; }
        public string SiteCollection { get; set; }
        public string WebUrl { get; set; }
    }
    
    public class MissingEventReceiversInput : Elementbase
    {
        public string Assembly { get; set; }
        public string EventName { get; set; }
        //public string EventReceiverId { get; set; }
        public string HostId { get; set; }
        public string HostType { get; set; }
        //public string Scope { get; set; }

    }
    public class MissingFeaturesInput : Elementbase
    {
        public string FeatureId { get; set; }
        //public string FeatureTitle { get; set; }
        public string Scope { get; set; }
        //public string Source { get; set; }
        //public string UpgradeStatus { get; set; }

    }
    public class MissingSetupFilesInput : Elementbase
    {
        public string SetupFileDirName { get; set; }
        //public string SetupFileExtension { get; set; }
        //public string SetupFileId { get; set; }
        public string SetupFileName { get; set; }
        //public string SetupFilePath { get; set; }
    }
}
