using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;
using System.Security;
using System.Net;

namespace Transformation.PowerShell.Common
{
    public class AuthenticationHelper
    {
        public SecureString GetSecureString(string input)
        {
            if (string.IsNullOrEmpty(input))
                throw new ArgumentException("Input string is empty and cannot be made into a SecureString", input);

            var secureString = new SecureString();
            foreach (char c in input.ToCharArray())
                secureString.AppendChar(c);

            return secureString;
        } 

        /// <summary>
        /// Returns a SharePointOnline ClientContext object 
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="tenantUser">User to be used to instantiate the ClientContext object</param>
        /// <param name="tenantUserPassword">Password of the user used to instantiate the ClientContext object</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetSharePointOnlineAuthenticatedContextTenant(string siteUrl, string tenantUser, string tenantUserPassword)
        {
            var spoPassword = GetSecureString(tenantUserPassword);
            SharePointOnlineCredentials sharepointOnlineCredentials = new SharePointOnlineCredentials(tenantUser, spoPassword);

            var ctx = new ClientContext(siteUrl);
            ctx.Credentials = sharepointOnlineCredentials;

            return ctx;
        }

        /// <summary>
        /// Returns a SharePoint on-premises / SharePoint Online Dedicated ClientContext object
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="user">User to be used to instantiate the ClientContext object</param>
        /// <param name="password">Password of the user used to instantiate the ClientContext object</param>
        /// <param name="domain">Domain of the user used to instantiate the ClientContext object</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetNetworkCredentialAuthenticatedContext(string siteUrl, string user, string password, string domain)
        {
            ClientContext clientContext = new ClientContext(siteUrl);
            clientContext.Credentials = new NetworkCredential(user, password, domain);
            return clientContext;
        }
    }
}
