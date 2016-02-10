using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Microsoft.Deployment.Compression.Cab;

namespace InfoPathScraper.Model
{
	using InfoPathScraper.Model.Feature;
	/// <summary>
	/// Represents a parseable InfoPath file. 
	/// </summary>
	abstract class InfoPathFile
	{
		private XDocument _xDocument;
		private List<InfoPathFeature> _features;

		#region Public interface
		public CabFileInfo CabFileInfo { get; protected set; }
		public XDocument XDocument { get { InitializeXDocument(); return _xDocument; } }

		/// <summary>
		/// Enumerates the features found in this InfoPathFile.
		/// </summary>
		public IEnumerable<InfoPathFeature> Features
		{
			get
			{
				InitializeFeatures();
				foreach (InfoPathFeature feature in _features)
					yield return feature;
				yield break;
			}
		}
		#endregion

		#region Abstract methods
		protected abstract IEnumerable<Func<XDocument, IEnumerable<InfoPathFeature>>> FeatureDiscoverers { get; }
		#endregion

		#region Private helpers
		/// <summary>
		/// Loads the XDocument that is the file from the cab. All parseable InfoPath files are xml documents.
		/// </summary>
		private void InitializeXDocument()
		{
			if (_xDocument != null) return;

			Stream content = CabFileInfo.OpenRead();
			// use the 3.5-compatible Load API that takes an XmlReader, that way this will work when targeting either .NET3.5 or later
			System.Xml.XmlReader reader = System.Xml.XmlReader.Create(content);
			_xDocument = XDocument.Load(reader);
		}

		/// <summary>
		/// Run all the FeatureDiscoverers for this file type. Each deriving InfoPathFile type
		/// defines what features it *might* contain.
		/// </summary>
		private void InitializeFeatures()
		{
			if (_features != null) return;
			_features = new List<InfoPathFeature>();
			foreach (Func<XDocument, IEnumerable<InfoPathFeature>> discoverer in FeatureDiscoverers)
				foreach (InfoPathFeature feature in discoverer.Invoke(XDocument))
					_features.Add(feature);
		}
		#endregion
	}
}
