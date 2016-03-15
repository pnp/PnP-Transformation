using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PeoplePickerRemediation.Console.Common.Base
{
    public class PeoplePickerListsInput
    {
        public string WebUrl { get; set; }
        public string ListName { get; set; }
    }
    public class PeoplePickerListOutput
    {

        public string WebUrl { get; set; }
        public string ListName { get; set; }
        public string ItemID { get; set; }
        public string Users { get; set; }
        public string Groups { get; set; }
        public string Status { get; set; }
        public string ErrorDetails { get; set; }

    }

}
