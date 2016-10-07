using System.Security;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;

namespace JDP.Transformation.HttpCommands
{
    public class RequestActivateSandboxSolution : RemoteOperation
    {

        #region CONSTRUCTORS

        public RequestActivateSandboxSolution(string TargetUrl, AuthenticationType authType, string User, string Password, string Domain = "") 
            : base(TargetUrl, authType, User, Password, Domain)
        {

        }

        #endregion

        #region PROPERTIES

        private int? solutionId;
        public int SolutionId
        {
            get { return solutionId.GetValueOrDefault(); }
            set { solutionId = value; }
        }

        public string SolutionName
        {
            get;
            set;
        }

        public override string OperationPageUrl
        {
            get
            {
                //if solutionname is set, use that to lookup the SolutionId
                if (!string.IsNullOrEmpty(SolutionName) && !solutionId.HasValue)
                {
                    using (ClientContext ctx = new ClientContext(TargetSiteUrl))
                    {
                        var spoPassword = new SecureString();
                        foreach (char c in this.Password)
                        {
                            spoPassword.AppendChar(c);
                        }
                        if (this.AuthType == AuthenticationType.Office365)
                        {
                            ctx.Credentials = new Microsoft.SharePoint.Client.SharePointOnlineCredentials(this.User, spoPassword);
                        }
                        else
                        {
                            ctx.Credentials = new NetworkCredential(this.User, spoPassword, this.Domain);
                        }

                        List solutionGallery = ctx.Web.Lists.GetByTitle("Solution Gallery");

                        CamlQuery query = new CamlQuery();
                        query.ViewXml = "<View><Query><Where><Eq><FieldRef Name=\"FileLeafRef\" /><Value Type=\"Text\">" + SolutionName + "</Value></Eq></Where></Query></View>";
                        ListItemCollection items = solutionGallery.GetItems(query); 

                        ctx.Load(items);
                        ctx.ExecuteQuery();
                        if (items.Count == 1)
                        {
                            SolutionId = items[0].Id;
                        }
                        else
                        {
                            //no match
                            SolutionId = -1;
                        }
                    }
                }

                return "/_catalogs/solutions/Forms/Activate.aspx?Op=ACT&ID=" + SolutionId.ToString();
            }
        }

        #endregion

        #region METHODS

        private string targetControlName;

        public override void AnalyzeRequestResponse(string page)
        {
            base.AnalyzeRequestResponse(page);

            //need to find the EventTarget control name something like ctl00$ctl34$g_9419a5d1_c889_46a9_ab51_c7ab392b6fb1$ctl00$ctl00$ctl00$toolBarTbl$RptControls$diidIOActivateSolutionItem
            string pattern = @"javascript:__doPostBack\(\&\#39;(.*?diidIOActivateSolutionItem)\&\#39;";
            RegexOptions regexOptions = RegexOptions.None;
            Regex regex = new Regex(pattern, regexOptions);
            Match m = regex.Match(page);
            if (m.Success)
            {
                targetControlName = m.Groups[1].Value;
            }
        }

        public override void SetPostVariables()
        {
            // Set operation specific parameters,
            this.PostParameters.Add("__EVENTTARGET", targetControlName.Replace("$", "%24"));
        }



        #endregion

    }
}
