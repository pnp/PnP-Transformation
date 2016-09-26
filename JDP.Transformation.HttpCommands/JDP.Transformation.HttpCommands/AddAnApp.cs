using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDP.Transformation.HttpCommands
{
    /// <summary>
    /// Add an app using addanapp.aspx
    /// </summary>
    public class AddAnApp : RemoteOperation
    {
        #region CONSTRUCTORS

        public AddAnApp(string TargetUrl, AuthenticationType authType, string User, string Password, string Domain = "")
            : base(TargetUrl, authType, User, Password, Domain)
        {
        }

        #endregion

        #region PROPERTIES

        public override string OperationPageUrl
        {
            get
            {
                return string.Format("/_layouts/15/addanapp.aspx");
            }
        }

        public string appid
        {
            get;
            set;
        }
        public string catalog
        {
            get;
            set;
        }
        public string oID
        {
            get;
            set;
        }

        #endregion

        #region METHODS

        public override void SetPostVariables()
        {
            // Set operation specific parameters
            this.PostParameters.Add("__EVENTTARGET", "");
            this.PostParameters.Add("__EVENTARGUMENT", "");
            this.PostParameters.Add("task", "AppDownload");
            this.PostParameters.Add("appid", appid);
            this.PostParameters.Add("oID", oID);
            this.PostParameters.Add("catalog", catalog);
        }

        #endregion
    }
}
