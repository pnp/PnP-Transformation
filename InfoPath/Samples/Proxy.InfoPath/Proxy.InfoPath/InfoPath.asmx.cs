using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Web;
using System.Web.Configuration;
using System.Web.Services;

using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client;

namespace Proxy.InfoPath
{
    /// <summary>
    /// Summary description for InfoPath
    /// </summary>
    [WebService(Namespace = "http://microsoft.mso.jdp.org/samples")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class InfoPath : System.Web.Services.WebService
    {
        // if true, we are targeting a v15 farm (SPO-D or on-prem); if false, we are targeting a v16 farm (SPO-MT or SPO-vNext)
        private static readonly bool TargetFarmIsSPOD = WebConfigurationManager.AppSettings["TargetFarmIsSPOD"].ToBoolean();

        // The client (i.e., the InfoPath form) must pass this value as the ClientValidationKey in order to be considered a "valid" client.
        private static readonly string ClientValidationKey = WebConfigurationManager.AppSettings["ClientValidationKey"];

        // We return this error message when the client validation key has not been configured.
        private static readonly string ClientConfigurationError = "Configuration failed. You do not have permission to perform this action or access this resource";
        // We return this error message when the client does not provide the correct client validation key.
        private static readonly string ClientValidationError = "Validation failed. You do not have permission to perform this action or access this resource";

        // This is the prefix used for accounts on v15 Farms (SPO-D and on-prem)
        private static readonly string AccountPrefix15 = "i:0#.w|";
        // This is the prefix used for accounts on v16 Farms (SPO-MT and SPO-vNext)
        private static readonly string AccountPrefix16 = "i:0#.f|membership|";

        //*******************************************************************************************************
        //NOTE: 
        //    This Sample simply stores the service account password as cleartext for demo purposes
        //    Passwords should never be stored as cleartext in configuration files
        //    For security purposes, your implementation should take steps to encrypt the password as necessary
        //*******************************************************************************************************
        private static readonly string ServiceAccountDomain = WebConfigurationManager.AppSettings["ServiceAccountDomain"];
        private static readonly string ServiceAccountUsername = WebConfigurationManager.AppSettings["ServiceAccountUsername"];
        private static readonly string ServiceAccountPassword = WebConfigurationManager.AppSettings["ServiceAccountPassword"];

        #region Testing

        // remove all members in this section once testing is completed.

        private static readonly string TestSiteCollectionUrl = WebConfigurationManager.AppSettings["TestSiteCollectionUrl"];
        private static readonly string TestUsername = WebConfigurationManager.AppSettings["TestUsername"];

        [WebMethod]
        public string HelloWorld()
        {
            return "Hello World";
        }

        [WebMethod]
        public string ServiceAccountTest()
        {
            return GetGroupCollectionsOfUserViaServiceAccount(TestSiteCollectionUrl, TestUsername, ClientValidationKey);
        }
        [WebMethod]
        public string AppRegistrationTest()
        {
            return GetGroupCollectionsOfUserViaAppRegistration(TestSiteCollectionUrl, TestUsername, ClientValidationKey);
        }
        #endregion

        #region Public API
        /// <summary>
        /// Examines the authorization model of the specified site collection and returns a list of names of the SharePoint Security Groups to which the specified user belongs.
        /// </summary>
        /// <param name="siteCollectionURL">a Fully-qualified URL that identifies the site collection to process</param>
        /// <param name="username">a username that identifies the user to process
        ///     When targeting v15 farms (SPO-D/on-prem), use the "domain\username" format 
        ///     When targeting v16 farms (SPO-MT/SPO-vNext), use the "username@domain.com" format 
        /// </param>
        /// <param name="clientValidationKey">a character sequence used to validate the caller</param>
        /// <returns>A comma-separated list of SharePoint Security Group names</returns>
        /// <remarks>
        ///     Sample method that demonstrates the use of a Service Account
        ///     Intended for use by an InfoPath form.
        ///     Consider renaming this method to remove the implementation hint (i.e., "ViaServiceAccount")
        /// </remarks>
        [WebMethod]
        public string GetGroupCollectionsOfUserViaServiceAccount(string siteCollectionURL, string username, string clientValidationKey)
        {
            try
            {
                ValidateStringArgument(siteCollectionURL, "siteCollectionURL");
                ValidateStringArgument(username, "username");

                ValidateClient(clientValidationKey);

                using (var clientContext = CreateAuthenticatedUserContext(siteCollectionURL, ServiceAccountDomain, ServiceAccountUsername, ServiceAccountPassword))
                {
                    return GetGroupCollectionsOfUser(clientContext, username);
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        /// <summary>
        /// Examines the authorization model of the specified site collection and returns a list of names of the SharePoint Security Groups to which the specified user belongs.
        /// </summary>
        /// <param name="siteCollectionURL">a Fully-qualified URL that identifies the site collection to process</param>
        /// <param name="username">a username that identifies the user to process
        ///     When targeting v15 farms (SPO-D/on-prem), use the "domain\username" format 
        ///     When targeting v16 farms (SPO-MT/SPO-vNext), use the "username@domain.com" format 
        /// </param>
        /// <param name="clientValidationKey">a character sequence used to validate the caller</param>
        /// <returns>A comma-separated list of SharePoint Security Group names</returns>
        /// <remarks>
        ///     Sample method that demonstrates the use of an App Registration
        ///     Intended for use by an InfoPath form.
        ///     Consider renaming this method to remove the implementation hint (i.e., "ViaAppRegistration")
        /// </remarks>
        [WebMethod]
        public string GetGroupCollectionsOfUserViaAppRegistration(string siteCollectionURL, string username, string clientValidationKey)
        {
            try
            {
                ValidateStringArgument(siteCollectionURL, "siteCollectionURL");
                ValidateStringArgument(username, "username");

                ValidateClient(clientValidationKey);

                using (var clientContext = CreateAppOnlyClientContext(siteCollectionURL))
                {
                    return GetGroupCollectionsOfUser(clientContext, username);
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        #endregion

        #region Implementation

        private void ValidateStringArgument(string siteCollectionURL, string argumentName)
        {
            if (String.IsNullOrEmpty(argumentName))
            {
                throw new ArgumentNullException(argumentName);
            }
        }

        private static void ValidateClient(string clientValidationKey)
        {
            // This sample assumes that all forms share the same client validation key.

            // this method could be enhanced/improved to implement site-specific validation keys.
            // - register the site and its key in a central registration list
            // - form passes its specific site url and validation key
            // - this method queries the registration list and validates the siteUrl/Key pair.

            if (String.IsNullOrEmpty(ClientValidationKey))
            {
                throw new ApplicationException(ClientConfigurationError);
            }

            if (clientValidationKey.Equals(ClientValidationKey, StringComparison.InvariantCulture) == false)
            {
                throw new ApplicationException(ClientValidationError);
            }
        }

        private static string GetGroupCollectionsOfUser(ClientContext clientContext, string username)
        {
            string encodedLoginName = GetEncodedLoginName(username);
            var output = String.Empty;

            var user = clientContext.Web.SiteUsers.GetByLoginName(encodedLoginName);
            clientContext.Load(user, l => l.Groups, l => l.Id, l => l.LoginName);
            clientContext.ExecuteQuery();

            foreach (var group in user.Groups)
            {
                output += group.Title + ",";
            }

            if (!String.IsNullOrEmpty(output))
            {
                output = output.Substring(0, output.Length - 1);
            }

            return output;
        }

        private static string GetEncodedLoginName(string username)
        {
            //v15 (SPO-D/on-prem):    encodedLoginName = "i:0#.w|domain\user";
            //v16 (SPO-MT/vNext):     encodedLoginName = "i:0#.f|membership|user@contoso.onMicrosoft.com";
            return String.Format("{0}{1}", (TargetFarmIsSPOD ? AccountPrefix15 : AccountPrefix16), username);
        }

        private static ClientContext CreateAuthenticatedUserContext(string siteUrl, string domain, string username, string password)
        {
            ClientContext userContext = new ClientContext(siteUrl);

            SecureString securePassword = new SecureString();
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }

            if (TargetFarmIsSPOD)
            {
                // use Windows authentication for v15 farms (SPO-D or On-Prem) 
                userContext.Credentials = new NetworkCredential(username, securePassword, domain);
            }
            else
            {
                // use o365 authentication for v16 farms (SPO-MT or vNext)
                userContext.Credentials = new SharePointOnlineCredentials(username, securePassword);
            }

            return userContext;
        }

        private static ClientContext CreateAppOnlyClientContext(string siteUrl)
        {
            var parentSiteUri = new Uri(siteUrl);
            string realm = TokenHelper.GetRealmFromTargetUrl(parentSiteUri);
            var token = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, parentSiteUri.Authority, realm).AccessToken;
            var clientContext = TokenHelper.GetClientContextWithAccessToken(parentSiteUri.ToString(), token);

            return clientContext;
        }
        #endregion
    }
}