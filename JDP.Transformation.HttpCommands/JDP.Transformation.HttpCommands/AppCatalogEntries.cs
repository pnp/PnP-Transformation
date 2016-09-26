using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace JDP.Transformation.HttpCommands
{
    /// <summary>
    /// Get a list of available apps via addanapp.aspx
    /// </summary>
    public class AppCatalogEntries : RemoteOperation
    {
        #region CONSTRUCTORS

        public AppCatalogEntries(string TargetUrl, AuthenticationType authType, string User, string Password, string Domain = "")
            : base(TargetUrl, authType, User, Password, Domain)
        {
        }

        #endregion

        #region PROPERTIES

        public override string OperationPageUrl
        {
            get
            {
                return "/_layouts/15/addanapp.aspx?task=GetMyApps";
            }
        }

        #endregion

        #region METHODS

        public override void SetPostVariables()
        {
        }

        public AppCatalogEntry[] GetAppCatalogEntries()
        {
            string page = GetRequest();
            DataContractJsonSerializer jsonSerializer = new DataContractJsonSerializer(typeof(AppCatalogEntry[]));
            MemoryStream ms = new MemoryStream(Encoding.ASCII.GetBytes(page));
            return jsonSerializer.ReadObject(ms) as AppCatalogEntry[];
        }

        public AppCatalogEntry GetAppCatalogEntry(Guid productId)
        {
            Guid entryId;
            foreach (var entry in GetAppCatalogEntries())
                if (Guid.TryParse(entry.ProductId, out entryId) && (entryId.Equals(productId)))
                    return entry;
            return null;
        }

        public AppCatalogEntry GetAppCatalogEntry(string title)
        {
            foreach (var entry in GetAppCatalogEntries())
                if (entry.Title.Equals(title, StringComparison.InvariantCulture))
                    return entry;
            return null;
        }

        #endregion
    }

    [DataContract]
    public class AppCatalogEntry
    {
        [DataMember(Name = "Catalog")]
        public string Catalog { get; set; }
        [DataMember(Name = "ID")]
        public string ID { get; set; }
        [DataMember(Name = "Title")]
        public string Title { get; set; }
        [DataMember(Name = "Language")]
        public string Language { get; set; }
        [DataMember(Name = "Version")]
        public string Version { get; set; }
        [DataMember(Name = "ProductId")]
        public string ProductId { get; set; }
    }

}
