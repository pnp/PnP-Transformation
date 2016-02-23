using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.PageLayout
{
    public class PageLayoutBase : Elementbase
    {
        public string PageId { get; set; }
        public string PageTitle { get; set; }
        public string NewPageLayoutUrl { get; set; }     
        public string NewPageLayoutDescription { get; set; }
        public string OldPageLayoutDescription { get; set; }
        public string OldPageLayoutUrl { get; set; }
    }

    public class PageLayoutInput : Inputbase
    {
        public string PageLayout_Name { get; set; }
        public string PageLayout_ServerRelativeUrl { get; set; }
    }
    
    public class PageInput : Inputbase
    {
        public string PageName { get; set; }
        public string PageUrl { get; set; }
        public string PageTitle { get; set; }
        public string PageId { get; set; }
        public string PageServerRelativeUrl { get; set; }
        public string PageLayoutServerRelativeUrl { get; set; }
    }
}