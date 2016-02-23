namespace Transformation.PowerShell.Common
{
    public static class Constants
    {
        public const string Quote = "\"";
        public const string EscapedQuote = "\"\"";
        public static char[] CharactersThatMustBeQuoted = { ',', '"', '\n' };
        public static readonly string CsvDelimeter = ",";
        public static readonly string NotApplicable = "N/A";
        
        public static readonly long ExceptionFileSizeinKb = 4096; //4 MB
        public static readonly string TraceLogFileSuffix = "TraceLog";
        public static readonly bool Logging = true;

        internal const string LOGGING_SOURCE = "WebPartTransformation";
        internal const string WEBPART_SERVICE = "/_vti_bin/webpartpages.asmx";
        internal const string TEAMSITE_PAGES_LIBRARY = "Site Pages";
        internal const string WEBPART_USAGE_ENTITY_FILENAME = "WebPartUsage.csv";
        internal const string WEBPART_PROPERTIES_FILENAME = "WebPartProperties.xml";
        internal const string OnPremise = "OP";
        internal const string Online = "OL";

        internal const string TARGET_WEBPART_XML_DIR = "TargetConfiguredWebPartXmls";
        internal const string SOURCE_WEBPART_XML_DIR = "SourceWebPartXmls";

        internal const string LOG_DIR = "Logs";

        internal const string ActionType_Web = "web";
        internal const string ActionType_SiteCollection = "sitecollection";
        internal const string ActionType_All = "all";
        internal const string ActionType_Blank = "";
        internal const string ActionType_CSV = "csv";

        internal const string Input_All = "all";
        internal const string Input_Blank = "";

        internal const string UnGhostFileOperation_Copy = "COPY";
        internal const string UnGhostFileOperation_Move = "MOVE";

        public static string[] GetOutPutFiles(string outPutFolder = "")
        {
            return new[]
            {
                outPutFolder + @"\" + MasterPageUsage,
                outPutFolder + @"\" + PageLayoutUsage,
                outPutFolder + @"\" + PagesUsage
            };
        }

        #region InputCSVFiles

        public static readonly string MasterPageInput = "MasterPage_Usage.csv";
        public static readonly string PageLayoutInput = "PageLayouts_Usage.csv";
        public static readonly string PagesInput = "Pages_Usage.csv";
        public static readonly string ContentTypeInput = "ContentType_Usage.csv";
        public static readonly string ContentTypeDuplicateInput = "ContentType_CSOM_Input.csv";
        public static readonly string SiteColumnDuplicateInput = "SiteColumn_CSOM_Input.csv";
        public static readonly string SiteColumnAddINContentTypeInput = "SiteColumn_addIN_ContentType_Input.csv";
        public static readonly string ContentType_Add_To_ListInput = "ContentType_Add_To_ListInput.csv";
        public static readonly string WebPart_DiscoveryFile_Input = "WebParts_Usage.csv";

        #endregion

        #region OutputCSVFiles

        public static readonly string MasterPageUsage = "MasterPage_Replace.csv";

        public static readonly string PageLayoutUsage = "PageLayouts_Replace.csv";
        public static readonly string PagesUsage = "Pages_Replace.csv";
        public static readonly string ContentTypeCreated = "ContentType_Created.csv";
        public static readonly string Exception = "Exception";
        public static readonly string ContentTypeDuplicateOutput = "ContentType_Usage_Replace.csv";
        public static readonly string SiteColumnDuplicateOutput = "SiteColumn_Usage_Replace.csv";
        public static readonly string SiteColumnAddINContentTypeOutput = "SiteColumn_AddTo_ContentType_Replace.csv";
        public static readonly string ContentType_Add_To_ListOutput = "ContentType_Add_To_List_Replace.csv";

        public static readonly string ListMigration_Output = "ListMigrationReport.csv";
        public static readonly string UnGhosting_Output = "UnGhostingReport.csv";
        public static readonly string UnGhosting_DownloadFileReport = "DownloadedFilesReport.csv";

        public static readonly string CustomField_Deletion_Output = "CustomFieldDeletionReport.csv";
        #endregion
    }
}