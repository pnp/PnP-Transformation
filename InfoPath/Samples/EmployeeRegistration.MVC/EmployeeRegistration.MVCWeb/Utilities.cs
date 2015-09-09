using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Hosting;

namespace EmployeeRegistration.MVCWeb
{
    public class Utilities
    {
        public static string serverRelativeUrl = null;
        public static void InitServerRelativeUrl(ClientContext ctx)
        {
            if (serverRelativeUrl == null)
            {
                ctx.Load(ctx.Web);
                ctx.ExecuteQuery();
                serverRelativeUrl = ctx.Web.ServerRelativeUrl;
            }
        }
        public static string ReplaceTokens(ClientContext ctx, string input)
        {
            InitServerRelativeUrl(ctx);
            string output = input.Replace("~sitecollection", serverRelativeUrl);
            return output;
        }
        public static string ReplaceTokensInAssetFile(ClientContext ctx, string filePath, string clientId, string redirectURI)
        {
            string fileContent = System.IO.File.ReadAllText(HostingEnvironment.MapPath(String.Format("~/{0}", filePath)));
            fileContent = ReplaceTokens(ctx, fileContent);
            fileContent = fileContent.Replace("%clientId%", clientId);
            fileContent = fileContent.Replace("%redirectURI%", redirectURI);
            string newFilePath = HostingEnvironment.MapPath(String.Format("~/{0}", (filePath + Guid.NewGuid().ToString("D"))));
            System.IO.File.WriteAllText(newFilePath, fileContent);
            return newFilePath;
        }
    }
}