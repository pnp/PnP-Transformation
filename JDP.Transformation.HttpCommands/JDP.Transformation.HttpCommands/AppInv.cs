using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDP.Transformation.HttpCommands
{
    /// <summary>
    /// Trust an app using appinv.aspx
    /// </summary>
    public class AppInv : RemoteOperation
    {
        #region CONSTRUCTORS

        public AppInv(string TargetUrl, AuthenticationType authType, string User, string Password, string Domain = "")
            : base(TargetUrl, authType, User, Password, Domain)
        {
        }

        #endregion

        #region PROPERTIES

        public override string OperationPageUrl
        {
            get
            {
                return string.Format("/_layouts/15/appinv.aspx?catalog={0}&appcatalogid={1}", CatalogNo, CatalogId.Replace("-", "%2D"));
            }
        }

        public string CatalogNo
        {
            get;
            set;
        }
        public string CatalogId
        {
            get;
            set;
        }

        #endregion

        #region METHODS

        public override void SetPostVariables()
        {
            // Set operation specific parameters
            this.PostParameters.Add("__EVENTTARGET", "ctl00$PlaceHolderMain$BtnAllow");
            this.PostParameters.Add("__EVENTARGUMENT", "");
        }

        #endregion
    }
}
