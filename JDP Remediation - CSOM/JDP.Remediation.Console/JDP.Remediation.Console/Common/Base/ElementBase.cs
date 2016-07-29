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
    public class TranformWebPartStatusBase
    {
        public string WebPartId { get; set; }
        public string ZoneID { get; set; }
        public string ZoneIndex { get; set; }
        public string PageUrl { get; set; }
        public string WebPartType { get; set; }
        public string Status { get; set; }
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
    public class WebPartUsageEntity
    {
        public string WebURL { get; set; }
        public string WebPartID { get; set; }
        public string WebPartTitle { get; set; }
        public string PageUrl { get; set; }
        public string ZoneIndex { get; set; }
        public string WebPartType { get; set; }
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

    public class MissingWorkflowAssociationsInput : Elementbase
    {
        public string DirName { get; set; }
        public string LeafName { get; set; }
    }
    public class ContentTypeBase
    {
        public string ContentTypeID { get; set; }
        public string ContentTypeName { get; set; }
        public string FeatureId { get; set; }

        public string FeatureName { get; set; }

        public string SolutionName { get; set; }

        public string IsDocumentSetContentType { get; set; }
    }
    public class CustomFieldBase
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string FeatureId { get; set; }
        public string FeatureName { get; set; }
        public string SolutionName { get; set; }
        public string Type { get; set; }
        public string Description { get; set; }
    }
    public class EventReceiversBase
    {
        public string FeatureID { get; set; }
        public string FeatureName { get; set; }
        public string SolutionName { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string ListTemplateId { get; set; }
        public string ListURL { get; set; }
        public string Assembly { get; set; }
        public string Class { get; set; }
    }
    public class PreMTMissingListTemplatesInGalleryInput : Elementbase
    {
        public string ListTemplateID { get; set; }
        public string ListTemplateName { get; set; }
        public string ListGalleryPath { get; set; }
        public string TimeLastModified { get; set; }
    }
    public class CustomizationHeadersBase : Elementbase
    {
        public string IsCustomizationPresent { get; set; }
        public string IsCustomizedContentType { get; set; }
        public string IsCustomizedEventReceiver { get; set; }
        public string IsCustomizedSiteColumn { get; set; }
    }
    public class EventReceiverInput
    {
        public string Assembly { get; set; }
    }
    public class ContentTypeInput
    {
        public string ContentTypeID { get; set; }
    }
    public class CustomFieldInput
    {
        public string ID { get; set; }
    }
    public class FeatureInput
    {
        public string FeatureID { get; set; }
    }

    public class ListTemplateFTCAnalysisOutputBase : ListCustomizationHeader
    {
        public string ListTemplateName { get; set; }
        public string ListGalleryPath { get; set; }
        public string WebApplication { get; set; }
        public string SiteCollection { get; set; }
        public string WebUrl { get; set; }
        public string CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
    }

    public class SiteTemplateFTCAnalysisOutputBase : SiteCustomizationHeader
    {
        public string SiteTemplateName { get; set; }
        public string SiteTemplateGalleryPath { get; set; }
        public string WebApplication { get; set; }
        public string SiteCollection { get; set; }
        public string WebUrl { get; set; }
        public string CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
    }

    public class ListCustomizationHeader
    {
        public string IsCustomizationPresent { get; set; }
        public string IsCustomizedContentType { get; set; }
        public string IsCustomizedEventReceiver { get; set; }
        public string IsCustomizedSiteColumn { get; set; }
        public string CTHavingCustomEventReceiver { get; set; }
    }

    public class SiteCustomizationHeader : ListCustomizationHeader
    {
        public string IsCustomizedFeature { get; set; }
    }
    public class MissingMasterPageInput
    {
        public string PageUrl { get; set; }
        public string SiteCollection { get; set; }
        public string WebUrl { get; set; }
        public string MasterUrl { get; set; }
        public string CustomMasterUrl { get; set; }
    }

    public class WebpartInput
    {
        public string PageUrl { get; set; }
        public string WebUrl { get; set; }
        public string StorageKey { get; set; }
        public string WebPartType { get; set; }
    }

    public class WebpartDeleteOutputBase : WebpartInput
    {
        public string Status { get; set; }
    }
}
