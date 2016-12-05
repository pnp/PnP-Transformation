using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;

namespace InfoPathScraper.Model.Feature
{
	class DataConnection : InfoPathFeature
	{
		#region Constants
		private const string query = @"query";
		private const string spListConnection = @"sharepointListAdapter";
		private const string spListConnectionRW = @"sharepointListAdapterRW";
		private const string soapConnection = @"webServiceAdapter";
		private const string xmlConnection = @"xmlFileAdapter"; // also used for REST!
		private const string adoConnection = @"adoAdapter";
		#endregion

		#region Public interface
		public string ConnectionType { get; private set; }
		public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
		{

			IEnumerable<XElement> allDataConnections = document.Descendants(xsfNamespace + query);
			foreach (XElement queryElement in allDataConnections)
			{
				yield return ParseDataConnection(queryElement);
			}

			// nothing left
			yield break;
		}

		public override string ToString()
		{
			return FeatureName + ": " + ConnectionType;
		}

		public override string ToCSV()
		{
			return ConnectionType;
		}
		#endregion

		#region Private helpers
		/// <summary>
		/// This should return DataConnection since every query element represents exactly one connection
		/// In special cases (SPList, Soap) we defer to a subclass to mine more data.
		/// </summary>
		/// <param name="queryElement"></param>
		/// <returns></returns>
		private static DataConnection ParseDataConnection(XElement queryElement)
		{
			XElement dataConnection = queryElement.Element(xsfNamespace + spListConnection);
			if (dataConnection != null)
				return SPListConnection.Parse(dataConnection);
			else if ((dataConnection = queryElement.Element(xsfNamespace + spListConnectionRW)) != null)
				return SPListConnection.Parse(dataConnection);
			else if ((dataConnection = queryElement.Element(xsfNamespace + soapConnection)) != null)
				return SoapConnection.Parse(dataConnection);
			else if ((dataConnection = queryElement.Element(xsfNamespace + xmlConnection)) != null)
				return XmlConnection.Parse(dataConnection);
			else if ((dataConnection = queryElement.Element(xsfNamespace + adoConnection)) != null)
				return AdoConnection.Parse(dataConnection);

			// else just grab the type and log that. Nothing else to do here.
			foreach (XElement x in queryElement.Elements())
			{
				if (dataConnection != null) throw new ArgumentException("More than one adapter found under a query node");
				dataConnection = x;
			}

			if (dataConnection == null) throw new ArgumentException("No adapter found under query node");
			DataConnection dc = new DataConnection();
			dc.ConnectionType = dataConnection.Name.LocalName;
			return dc;
		}
		#endregion
	}

	/// <summary>
	/// Subclass specifically for mining SP List connections
	/// </summary>
	class SPListConnection : DataConnection
	{
		#region Constants
		private const string siteUrlAttribute = @"siteUrl";
		private const string siteURLAttribute = @"siteURL";
		private const string listGuidAttribute = @"sharePointListID";
		private const string sharepointGuidAttribute = @"sharepointGuid";
		private const string submitAllowedAttribute = @"submitAllowed";
		private const string relativeListUrlAttribute = @"relativeListUrl";
		private const string field = @"field";
		private const string typeAttribute = @"type";
		private static string[] validTypes =
		{
			"Counter",
			"Integer",
			"Number",
			"Currency",
			"Text",
			"Choice",
			"Plain",
			"Compatible",
			"FullHTML",
			"DateTime",
			"Boolean",
			"Lookup",
			"LookupMulti",
			"MultiChoice",
			"URL",
			"User",
			"UserMulti",
			"Calculated",
			"Attachments",
			"HybridUser",
		};
		private InfoPathScraper.Utilities.BucketCounter _bucketCounter;
		#endregion

		private SPListConnection()
		{
			_bucketCounter = new InfoPathScraper.Utilities.BucketCounter();
			foreach (string type in validTypes)
				_bucketCounter.DefineKey(type);
		}

		#region Public interface
		public string SiteUrl { get; private set; }
		public string ListGuid { get; private set; }
		public string SubmitAllowed { get; private set; }
		public string RelativeListUrl { get; private set; }
		public IEnumerable<KeyValuePair<string, int>> ColumnTypes
		{
			get
			{
				foreach (KeyValuePair<string, int> kvp in _bucketCounter.Buckets)
					yield return kvp;
				yield break;
			}
		}

		public override string ToString()
		{
			StringBuilder sb = new StringBuilder();
			sb.Append(FeatureName).Append(": ");
			sb.Append("{").Append(SiteUrl).Append(", ").Append(RelativeListUrl).Append("} ");
			if (IsV2)
			{
				sb.Append("Types: {");
				foreach (KeyValuePair<string, int> kvp in ColumnTypes)
				{
					// for human-readable, just emit non-zero counts
					if (kvp.Value > 0)
						sb.Append(kvp.Key).Append("=").Append(kvp.Value).Append(" ");
				}
				sb.Append("}");
			}
			return sb.ToString().Trim();
		}

		public override string ToCSV()
		{
			StringBuilder sb = new StringBuilder();
			sb.Append(SiteUrl).Append(",").Append(RelativeListUrl).Append(",");
			if (IsV2)
			{
				foreach (KeyValuePair<string, Int32> kvp in ColumnTypes)
				{
					sb.Append(kvp.Key).Append(",").Append(kvp.Value).Append(",");
				}
			}
			return sb.ToString();
		}

		private bool IsV2 { get; set; }

		/// <summary>
		/// Self-explanatory
		/// </summary>
		/// <param name="dataConnection"></param>
		/// <returns></returns>
		public static SPListConnection Parse(XElement dataConnection)
		{
			SPListConnection spl = new SPListConnection();
			spl.IsV2 = dataConnection.Name.LocalName.Equals("sharepointListAdapterRW");
			if (spl.IsV2)
			{
				spl.SiteUrl = dataConnection.Attribute(siteURLAttribute).Value;
				spl.ListGuid = dataConnection.Attribute(listGuidAttribute).Value;
				spl.SubmitAllowed = dataConnection.Attribute(submitAllowedAttribute).Value;
				spl.RelativeListUrl = dataConnection.Attribute(relativeListUrlAttribute).Value;

				// we should also scrape out the queried column types, this shows us what types of data we consume from lists
				foreach (XElement fieldElement in dataConnection.Elements(xsfNamespace + field))
				{
					string fieldType = fieldElement.Attribute(typeAttribute).Value;
					spl._bucketCounter.IncrementKey(fieldType);
				}
			}
			else // if (dataConnection.Name.LocalName.Equals("sharepointListAdapter"))
			{
				spl.SiteUrl = dataConnection.Attribute(siteUrlAttribute).Value;
				spl.ListGuid = dataConnection.Attribute(sharepointGuidAttribute).Value;
				spl.SubmitAllowed = dataConnection.Attribute(submitAllowedAttribute).Value;
			}

			return spl;
		}
		#endregion
	}

	/// <summary>
	/// Subclass of DataConnection specifically for Soap web service calls.
	/// </summary>
	class SoapConnection : DataConnection
	{
		#region Constants
		private const string wsdlUrlAttribute = @"wsdlUrl";
		private const string serviceUrlAttribute = @"serviceUrl";
		private const string nameAttribute = @"name";
		private const string operation = @"operation";
		private const string input = @"input";
		private const string sourceAttribute = @"source";

		private const string webServiceAdapterExtension = @"webServiceAdapterExtension";
		private const string refAttribute = @"ref";
		private const string connectoid = @"connectoid";
		private const string udcxExt = @".udcx";
		#endregion

		private SoapConnection() { }

		#region Public interface
		public string ServiceUrl { get; private set; }
		public string ServiceMethod { get; private set; }

		public static DataConnection Parse(XElement dataConnection)
		{
			XElement udcExtension = null;
			if (IsConnectionUDCX(dataConnection, out udcExtension) && udcExtension != null)
			{
				return UdcConnection.Parse(dataConnection, udcExtension);
			}
			return ParseInternal(dataConnection);
		}

		public override string ToString()
		{
			return FeatureName + ": " + ServiceUrl + "::" + ServiceMethod;
		}

		public override string ToCSV()
		{
			return ServiceUrl + "," + ServiceMethod;
		}
		#endregion

		#region Private helpers
		private static SoapConnection ParseInternal(XElement dataConnection)
		{
			SoapConnection sc = new SoapConnection();
			XElement op = dataConnection.Element(xsfNamespace + operation);
			if (op == null)
			{
				sc.ServiceUrl = dataConnection.Attribute(wsdlUrlAttribute).Value;
				sc.ServiceMethod = dataConnection.Attribute(nameAttribute).Value;
			}
			else
			{
				sc.ServiceUrl = op.Attribute(serviceUrlAttribute).Value;
				sc.ServiceMethod = op.Attribute(nameAttribute).Value;
				XElement inp = op.Element(xsfNamespace + input);
				if (inp != null && sc.ServiceUrl.Equals(""))
				{
					sc.ServiceUrl = inp.Attribute(sourceAttribute).Value;
					sc.ServiceMethod = "?";
				}
			}
			if (sc.ServiceUrl.Equals("")) Console.WriteLine(dataConnection.ToString());
			return sc;
		}


		/// <summary>
		/// Need to find the xsf2:webServiceAdapterExtension node elsewhere in the XDocument
		/// Need to find the one that has ref the same as our connection name 
		/// </summary>
		/// <param name="dataConnection"></param>
		/// <returns></returns>
		private static bool IsConnectionUDCX(XElement dataConnection, out XElement udcExtension)
		{
			udcExtension = null;
			XAttribute name = dataConnection.Attribute(nameAttribute);
			if (name == null) return false;

			string connectionName = name.Value;
			foreach (XElement webServiceExt in dataConnection.Document.Descendants(xsf2Namespace + webServiceAdapterExtension))
			{
				XAttribute refAtt = webServiceExt.Attribute(refAttribute);
				if (refAtt == null) continue; // No name = no match
				if (!refAtt.Value.Equals(connectionName)) continue; // These are not the extensions you are looking for ... *waves hand*

				XElement connect = webServiceExt.Element(xsf2Namespace + connectoid);
				if (connect == null) return false;
				XAttribute source = connect.Attribute(sourceAttribute);
				if (source == null) return false;

				if (Path.GetExtension(source.Value) != udcxExt) return false;

				udcExtension = webServiceExt;
				return true;
			}
			return false;
		}
		#endregion
	}

	/// <summary>
	/// Subclass of DataConnection specifically for UDCX Soap web service calls.
	/// </summary>
	class UdcConnection : DataConnection
	{
		#region Constants
		private const string connectoid = @"connectoid";
		private const string nameAttribute = @"name";
		private const string sourceAttribute = @"source";
		#endregion

		private UdcConnection() { }

		#region Public interface
		public string SourceUrl { get; private set; }
		public string MethodName { get; private set; }

		public new static UdcConnection Parse(XElement dataConnection, XElement udcExtension)
		{
			UdcConnection uc = new UdcConnection();

			XElement connect = udcExtension.Element(xsf2Namespace + connectoid);

			uc.MethodName = connect.Attribute(nameAttribute).Value;
			uc.SourceUrl = connect.Attribute(sourceAttribute).Value;

			return uc;
		}

		public override string ToString()
		{
			return FeatureName + ": " + SourceUrl + "::" + MethodName;
		}

		public override string ToCSV()
		{
			return SourceUrl + "," + MethodName;
		}
		#endregion
	}

	/// <summary>
	/// Identifies a connection to an Xml file
	/// </summary>
	class XmlConnection : DataConnection
	{
		#region Constants
		private const string fileUrlAttribute = @"fileUrl";
		private const string nameAttribute = @"name";
		private const string refAttribute = @"ref";
		private const string xmlFileAdapterExtension = @"xmlFileAdapterExtension";
		private const string isRestAttribute = @"isRest";
		#endregion

		protected XmlConnection() { }

		#region Public interface
		public string Url { get; private set; }

		public static DataConnection Parse(XElement dataConnection)
		{
			XmlConnection xc = null;
			bool isRest = IsConnectionRest(dataConnection);
			xc = isRest ? new RESTConnection() : new XmlConnection();

            string fileUrl = dataConnection.Attribute(fileUrlAttribute).Value;

            if (!String.IsNullOrEmpty(fileUrl))
            {
                // We have an embedded XmlConnection
                xc.Url = dataConnection.Attribute(fileUrlAttribute).Value;
                return xc;
            }
            else
            {
                // The XmlConnection is stored in a UDCX connection file
                XElement udcExtension = FindChild(dataConnection);
                return UdcConnection.Parse(dataConnection, udcExtension);
            }
		}

		public override string ToString()
		{
			return FeatureName + ": " + Url;
		}

		public override string ToCSV()
		{
			return Url;
		}
        #endregion

        #region Private helpers
        private static XElement FindChild(XElement dataConnection)
        {
            XAttribute name = dataConnection.Attribute(nameAttribute);
            if (name == null) return null;

            string connectionName = name.Value;
            foreach (XElement xmlExt in dataConnection.Document.Descendants(xsf2Namespace + xmlFileAdapterExtension))
            {
                XAttribute refAtt = xmlExt.Attribute(refAttribute);
                if (refAtt == null) continue; // No name = no match
                if (!refAtt.Value.Equals(connectionName)) continue; // These are not the extensions you are looking for ... *waves hand*
                return xmlExt;
            }
            return null;
        }

        /// <summary>
        /// Need to find the xsf2:xmlFileAdapterExtension node elsewhere in the XDocument
        /// Need to find the one that has ref the same as our connection name, then look for isRest="[bool]" in there
        /// </summary>
        /// <param name="dataConnection"></param>
        /// <returns></returns>
        private static bool IsConnectionRest(XElement dataConnection)
		{
            XElement xmlExt = FindChild(dataConnection);
            if (xmlExt == null) return false;

			XAttribute isRest = xmlExt.Attribute(isRestAttribute);
			if (isRest == null) return false;

            return isRest.Value.Equals("yes");
		}
        #endregion
    }

    class RESTConnection : XmlConnection
	{
	}

	class AdoConnection : DataConnection
	{
		#region Constants
		private const string connectionStringAttribute = @"connectionString";
		#endregion

		private AdoConnection() { }

		#region Public interface
		public string ConnectionString { get; private set; }

		public static AdoConnection Parse(XElement dataConnection)
		{
			AdoConnection ac = new AdoConnection();
			ac.ConnectionString = dataConnection.Attribute(connectionStringAttribute).Value;
			return ac;
		}

		public override string ToString()
		{
			return FeatureName + ": " + ConnectionString;
		}

		public override string ToCSV()
		{
			return ConnectionString;
		}
		#endregion
	}
}
