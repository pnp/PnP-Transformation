namespace JDP.Transformation.HttpCommands
{
    public class OpenSiteClosedByPolicy : RemoteOperation
    {
        public OpenSiteClosedByPolicy(string TargetUrl, AuthenticationType authType, string User, string Password, string Domain = "") : base(TargetUrl, authType, User, Password, Domain)
        {
        }

        public override string OperationPageUrl
        {
            get
            {
                return @"/_layouts/15/ProjectPolicyAndLifecycle.aspx";
            }
        }

        public override void SetPostVariables()
        {
            // Set operation specific parameters
            this.PostParameters.Add("ctl00$PlaceHolderMain$ctl00$buttonOpenProject", "Open this site");
        }
    }
}