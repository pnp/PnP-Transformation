using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDP.Remediation.Console
{
    /// <summary>
    /// This class holds all constants used by the program.  No code
    /// </summary>
    public class Constants
    {
        public const string Quote = "\"";
        public const string EscapedQuote = "\"\"";
        public static char[] CharactersThatMustBeQuoted = { ',', '"', '\n' };
        public static readonly string CsvDelimeter = ",";
        public static readonly string NotApplicable = "N/A";
        public static readonly string CSVExtension = ".csv";
        public static readonly long ExceptionFileSizeinKb = 4096; //4 MB
        public static readonly bool Logging = true;
        public static readonly string Exception = "Exception-" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + CSVExtension;
        public static readonly string TraceLogFileSuffix = "TraceLog";


        /// <summary>
        // Excel/CSV Cell CharacterLimit. According to Microsoft's documentation: 
        // https://support.office.com/en-us/article/Excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
        // Excel cannot read more than 32767 characters in a single cell
        // Total number of characters that a cell can contain: 32,767 characters
        /// </summary>
        public static readonly int CharacterLimitForCsvCell = 32758;
        public static readonly long OutputFileSizeinKb = 8192; //8 MB

        //private static readonly string portalWebAppRoot = "https://intranet.poc.com";
        //public static readonly string PortalRootSiteUrl = portalWebAppRoot + "/";
        public static readonly string WEBPART_SERVICE = "/_vti_bin/webpartpages.asmx";
        public static readonly string PropertyBagInheritMaster = "__InheritsMasterUrl";
        public static readonly string PropertyBagInheritCustomMaster = "__InheritsCustomMasterUrl";
        public static readonly string SOURCE_WEBPART_XML_DIR = "SourceWebPartXmls";
        internal const string WEBPART_PROPERTIES_FILENAME = "WebPartProperties.xml";

        public static readonly string MASTERPAGE_CONTENT_TYPE = "0x01010500B45822D4B60B7B40A2BFCC0995839404";

        // these are simple Text files and are NOT expected to contain a Header row
        public static readonly string UsageReport_SitesInputFileName = "sites.txt";

        public static readonly string AddWebPartStatusFileName = "AddWebPart_SuccessFailure-";
        public static readonly string ReplaceWebPartStatusFileName = "ReplaceWebPart_SuccessFailure-";
        public static readonly string ReplaceMasterPageFileName = "ReplaceMasterPage_SuccessFailure-";

        public static readonly string Success = "Success";
        public static readonly string Failure = "Failure";

        public static readonly string LockedMasterPagesFilesInputFileName = "LockedMasterPageFiles.csv";
        public static readonly string DeviceChannelMappingFileName = "__DeviceChannelMappings.aspx";

        public static readonly string MissingSetupFilesInputFileName = "PreMT_MissingSetupFile.csv";
        public static readonly string MissingWorkflowAssociationsInputFileName = "PreMT_MissingWorkflowAssociations.csv";
        
        public static readonly string WEBPART_USAGE_ENTITY_FILENAME = "WebPartUsage-";
        public static readonly string TARGET_WEBPART_XML_DIR = "TargetConfiguredWebPartXmls";

        // OOB Site-level Features IDs of interest...
        public static readonly Guid SharePointEnterpriseFeatures_SiteFeatureID = new Guid("8581a8a7-cf16-4770-ac54-260265ddb0b2");
        public static readonly Guid SharePointStandardFeatures_SiteFeatureID = new Guid("b21b090c-c796-4b0f-ac0f-7ef1659c20ae");
        public static readonly Guid SearchWebPartsAndTemplates_SiteFeatureID = new Guid("9c0834e1-ba47-4d49-812b-7d4fb6fea211");
        public static readonly Guid SitePolicy_SiteFeatureID = new Guid("2fcd5f8a-26b7-4a6a-9755-918566dba90a");

        //Webpart Deletion
        public static readonly string WebpartType_All = "all";

        //SucessFailure .csv files
        public static readonly string DeleteEventReceiversStatus = "DeleteEventReceivers_SuccessFailure-";
        public static readonly string DeleteFeatureStatus = "DeleteFeatures_SuccessFailure-";
        public static readonly string DeleteSetupFileStatus = "DeleteSetupFiles_SuccessFailure-";
        public static readonly string DeleteWorkflowAssociationsStatus = "DeleteWorkflowAssociations_SuccessFailure-";
        public static readonly string DeleteListTemplateStatus = "DeleteListTemplates_SuccessFailure-";
        public static readonly string DeleteWebpartStatus = "DeleteWebparts_SuccessFailure-";

        //ColumnORFieldAndTypeUsage
        public static readonly string ContentTypeAndCustomFieldFileName = "SiteColumnOrFieldAndContentTypeUsage-";
        public static readonly string ContentTypes = "ContentType";
        public static readonly string CustomFields = "SiteColumn/CustomField";

        public static string WspExtension = ".wsp";
        public static string StpExtension = ".stp";

        //NonDefaultMasterPage
        public static readonly string NonDefaultMasterPageFileName = "NonDefaultMasterPageUsage-";

        //Generate Site Collections Reports
        public static readonly string GenSiteCollectionFileName = "SiteCollectionsReport-";

        #region ListTemplate and SiteTemplate
        public static readonly string NoInputFile = "No Input File";
        //Input files for ListTemplateIntermediate Report & SiteTemplate Intermediate Report
        public static readonly string ContentTypeInput = "ContentTypes.csv";
        public static readonly string CustomFieldsInput = "CustomFields.csv";
        public static readonly string EventReceiversInput = "EventReceivers.csv";
        public static readonly string FeaturesInput = "Features.csv";
        public static readonly string SiteCollectionsTextFile = "SiteCollections.txt";

        #region ListTemplate
        //Download Folder
        public static readonly string DownloadPathListTemplates = "L";
        public static readonly string ListTemplateCustomizationUsage = "ListTemplateCustomization_Usage.csv";
        #endregion ListTemplate

        #region SiteTemplate
        //Download Folder
        public static readonly string DownloadPathSiteTemplates = "S";
        public static readonly string SiteTemplateCustomizationUsage = "SiteTemplateCustomization_Usage.csv";
        #endregion SiteTemplate

        #endregion ListTemplate and SiteTemplate

        //Generate Security Group Reports
        public static readonly string SecurityGroupsInputFileName = "securityGroups.txt";

        public static readonly string[] LegacyMasterPageFilesToDelete = 
        { 
            "Sample.master"
        };

        // Maintenance Banner
        public static readonly string EmbedJsFileName = "embedMaintenanceBanner.js";
        public static readonly string ScriptLinkLocation = "ScriptLink";
        public static readonly string ScriptLinkDescription = "MaintenanceBanner";


        //Pivot Constants

        /// <summary>
        // Excel/CSV Cell CharacterLimit. According to Microsoft's documentation: 
        // https://support.office.com/en-us/article/Excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
        // Excel cannot read more than 32767 characters in a single cell
        // Total number of characters that a cell can contain: 32,767 characters
        /// </summary>

        public static readonly long MergingFileSizeinKb = 700001;//700 MB

        //Pivot Configuration
        public static string SliceAndDiceSheet_Suffix = "_SDView";
        public static int SheetNameMaxLength = 31;
        public static string DTFileName = "Discovery Usage";
        public static string PreMTFileName = "Pre Migration Scan";
    }
}
