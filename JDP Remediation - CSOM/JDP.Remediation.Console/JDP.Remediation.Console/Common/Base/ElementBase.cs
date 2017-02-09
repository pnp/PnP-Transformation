using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDP.Remediation.Console.Common.Base
{
    public class ElementBase
    {
        public string ContentDatabase { get; set; }
        public string WebApplication { get; set; }
        public string SiteCollection { get; set; }
        public string WebUrl { get; set; }
    }
   
    public class ReplaceMasterPageStatusBase
    {
        public string CustomMasterPageUrl { get; set; }
        public string OOTBMasterPageUrl { get; set; }
        public string Status { get; set; }
        public string WebApplication { get; set; }
        public string SiteCollection { get; set; }
        public string WebUrl { get; set; }
        public string ExecutionDateTime { get; set; }
    }
    public class AddWebPartStatusBase
    {
        public string ZoneID { get; set; }
        public string ZoneIndex { get; set; }
        public string PageUrl { get; set; }
        public string WebPartFileName { get; set; }
        public string Status { get; set; }
        public string WebApplication { get; set; }
        public string SiteCollection { get; set; }
        public string WebUrl { get; set; }
        public string ExecutionDateTime { get; set; }
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
        public string ExecutionDateTime { get; set; }
    }
    public class MissingEventReceiversInput : ElementBase
    {
        public string Assembly { get; set; }
        public string EventName { get; set; }
        //public string EventReceiverId { get; set; }
        public string HostId { get; set; }
        public string HostType { get; set; }
        //public string Scope { get; set; }

    }
    public class MissingEventReceiversOutput
    {
        public string Assembly { get; set; }
        public string EventName { get; set; }
        public string HostId { get; set; }
        public string HostType { get; set; }
        public string Status { get; set; }
        public string WebApplication { get; set; }
        public string SiteCollection { get; set; }
        public string WebUrl { get; set; }
        public string ExecutionDateTime { get; set; }
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
    public class MissingFeaturesInput : ElementBase
    {
        public string FeatureId { get; set; }
        //public string FeatureTitle { get; set; }
        public string Scope { get; set; }
        //public string Source { get; set; }
        //public string UpgradeStatus { get; set; }

    }
    public class MissingFeatureOutput
    {
        public string FeatureId { get; set; }
        public string Scope { get; set; }
        public string WebApplication { get; set; }
        public string WebUrl { get; set; }
        public string SiteCollection { get; set; }
        public string Status { get; set; }
        public string ExecutionDateTime { get; set; }
    }
    public class MissingSetupFilesInput : ElementBase
    {
        public string SetupFileDirName { get; set; }
        //public string SetupFileExtension { get; set; }
        //public string SetupFileId { get; set; }
        public string SetupFileName { get; set; }
        //public string SetupFilePath { get; set; }
    }
    public class MissingSetupFilesOutput
    {
        public string SetupFileDirName { get; set; }
        public string SetupFileName { get; set; }
        public string Status { get; set; }
        public string WebUrl { get; set; }
        public string WebApplication { get; set; }
        public string ExecutionDateTime { get; set; }
    }

    public class MissingWorkflowAssociationsInput : ElementBase
    {
        public string DirName { get; set; }
        public string LeafName { get; set; }
    }
    public class MissingWorkflowAssociationsOutput
    {
        public string DirName { get; set; }
        public string LeafName { get; set; }
        public string Status { get; set; }
        public string WebApplication { get; set; }
        public string SiteCollection { get; set; }
        public string WebUrl { get; set; }
        public string ExecutionDateTime { get; set; }
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
    public class PreMTMissingListTemplatesInGalleryInput : ElementBase
    {
        public string ListTemplateID { get; set; }
        public string ListTemplateName { get; set; }
        public string ListGalleryPath { get; set; }
        public string TimeLastModified { get; set; }
    }
    public class CustomizationHeadersBase : ElementBase
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
    public class ListTemplateDeleteOutput
    {
        public string ListTemplateName { get; set; }
        public string ListGalleryPath { get; set; }
        public string WebApplication { get; set; }
        public string SiteCollection { get; set; }
        public string WebUrl { get; set; }
        public string Status { get; set; }
        public string ExecutionDateTime { get; set; }
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
        public string ExecutionDateTime { get; set; }
    }

    public class InputContentTypeBase
    {
        public string ContentTypeID { get; set; }
        public string ContentTypeName { get; set; }

    }

    public class InputCustomFieldBase
    {
        public string ID { get; set; }
        public string Name { get; set; }
    }

    public class ContentTypeCustomFieldOutput
    {
        public string ComponentName { get; set; }
        public string ListId { get; set; }
        public string ListTitle { get; set; }
        public string ContentTypeOrCustomFieldId { get; set; }
        public string ContentTypeOrCustomFieldName { get; set; }
        public string WebUrl { get; set; }
        public string SiteCollection { get; set; }

    }

    public class NonDefaultMasterpageOutput
    {
        public string MasterUrl { get; set; }
        public string CustomMasterUrl { get; set; }
        public string SiteCollection { get; set; }
        public string WebUrl { get; set; }
    }

    public class GenerateSiteCollectionnOutput
    {
        public string SiteCollectionUrl { get; set; }
    }

    public class GenerateSecurityGroupOutput
    {
        public string SecurityGroupName { get; set; }
        public string SiteCollectionUrl { get; set; }
    }

    public class LockedMasterPageFilesInput : ElementBase
    {
        public string SetupFileDirName { get; set; }
        public string SetupFileExtension { get; set; }
        public string SetupFileName { get; set; }
    }
    public class LockedMasterPageFilesOutput
    {
        public string SetupFileDirName { get; set; }
        public string SetupFileName { get; set; }
        public string Status { get; set; }
        public string WebUrl { get; set; }
        public string WebApplication { get; set; }
        public string ExecutionDateTime { get; set; }
        public string MappingFile { get; set; }
        public string MappingBackup { get; set; }
        public string MappingMasterPageRef { get; set; }
    }

    public class ManageMaintenanceBannersOutput
    {
        public string SiteCollectionUrl { get; set; }
        public string BannerOperation { get; set; }
        public string ScriptLinkName { get; set; }
        public string ScriptLinkFile { get; set; }
        public string Status { get; set; }
    }

}
