using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.SiteColumnAndContentTypes
{
    public class SiteColumnBase : Elementbase
    {
        public string Old_SiteColumn_Title { get; set; }
        public string Old_SiteColumn_InternalName { get; set; }
        public string Old_SiteColumn_ID { get; set; }
        public string Old_SiteColumn_Type { get; set; }
        public string old_SiteColumn_Scope { get; set; }

        public string New_SiteColumn_Title { get; set; }
        public string New_SiteColumn_InternalName { get; set; }
        public string New_SiteColumn_ID { get; set; }
        public string New_SiteColumn_Type { get; set; }
        public string New_SiteColumn_Scope { get; set; }
    }
    public class SiteColumnInput : Inputbase
    {
        public string CustomFieldTitle { get; set; }
        public string CustomFieldInternalName { get; set; }
        public string CustomFieldId { get; set; }
        public string CustomFieldType { get; set; }
        public string CustomFieldScope { get; set; }
        public string CustomFieldListTitle { get; set; }
    }
    public class AddSiteColumnToContentTypeBase : Elementbase
    {
        public string ContentTypeName { get; set; }
        public string SiteColumnName { get; set; }

        public string ContentTypeID { get; set; }
        public string SiteColumnID { get; set; }
    }
       
    public class UpdateContentTypeinListBase : Elementbase
    {
        public string ListName { get; set; }
        public string oldContentTypeId { get; set; }
        public string newContentTypeName { get; set; }
    }
    public class UpdateContentTypeinListInput : Inputbase
    {
        public string ListName { get; set; }
        public string oldContentTypeId { get; set; }
        public string newContentTypeName { get; set; }
    }

    public class ContentTypeBase : Elementbase
    {
        public string OldContentTypeID { get; set; }
        public string OldContentTypeName { get; set; }
        public string NewContentTypeID { get; set; }
        public string NewContentTypeName { get; set; }
        
    }
    public class ContentTypeInput : Inputbase
    {
        public string ContentTypeId { get; set; }
        public string ContentTypeName { get; set; }
    }

    public class CustomFieldIDInput : CustomFieldBase
    {
        public string CustomFieldId { get; set; }
    }

    public class CustomFieldIdAndTypeInput : CustomFieldBase
    {
        public string CustomFieldType { get; set; }

        public string CustomFieldId { get; set; }
    }

    public class CustomFieldOutput
    {
        public string WebUrl { get; set; }
        public string CustomFieldUsedIn { get; set; }
        public string CustomFieldScope { get; set; }
        public string CustomField { get; set; }
        public string CustomFieldName { get; set; }

        public string StatusMessage { get; set; }

        public string FieldID { get; set; }

        public string ListID { get; set; }

        public string ListTitle { get; set; }

        public string ExceptionMessage { get; set; }
    }
}
