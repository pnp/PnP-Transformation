using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Transformation.PowerShell.Base;

namespace Transformation.PowerShell.ListMigration
{
    public class ListMigrationBase : Elementbase
    {
        public string New_ListTitle { get; set; }
        public string New_ListID { get; set; }
        public string New_ListBaseTemplate { get; set; }
        public string Old_ListTitle { get; set; }
        public string Old_ListID { get; set; }
        public string Old_ListBaseTemplate { get; set; }
        
    }

    public class ListMigrationInput : Inputbase
    {
        public string ListTitle { get; set; }
        public string ListId { get; set; }
        public string ListBaseTemplate { get; set; }
        public string AssociatedContentTypes { get; set; }
    }
}
