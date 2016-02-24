using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.WebPart
{
    public class WebPartTransformationBase
    {
        
    }
    public class WebPartUsageEntity
    {
        public string WebURL { get; set; }
        public string WebPartID { get; set; }
        public string WebPartTitle { get; set; }
        public string PageUrl { get; set; }
        public string ZoneIndex { get; set; }
        public string WebPartType { get; set; }
    }

    public class WebPartInfo
    {
        public string WebURL { get; set; }
        public string WebPartID { get; set; }
        public string WebPartTitle { get; set; }
        public string PageUrl { get; set; }
        public string ZoneIndex { get; set; }
        public string WebPartType { get; set; }
        public string WebPartTypeFullName { get; set; }
        public string AssemblyQualifiedName { get; set; }
        public bool IsClosed { get; set; }
        public string RepresentedWebPartType { get; set; }
        public bool IsVisible { get; set; }
    }

    public class WebPartDiscoveryInput 
    {
        public string WebUrl { get; set; }
        public string WebPartId { get; set; }
        public string ZoneID { get; set; }
        public string ZoneIndex { get; set; }
        public string PageUrl { get; set; }
        public string WebPartType { get; set; }
        public string StorageKey { get; set; }
    }
    public class TranformWebPartStatusBase : Elementbase
    {
        public string WebPartId { get; set; }
        public string ZoneID { get; set; }
        public string ZoneIndex { get; set; }
        public string PageUrl { get; set; }
        public string WebPartType { get; set; }
        public string Status { get; set; }
    }
}
