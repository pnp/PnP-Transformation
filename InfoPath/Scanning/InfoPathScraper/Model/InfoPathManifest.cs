using System;
using System.Collections.Generic;
using System.Xml.Linq;
using Microsoft.Deployment.Compression.Cab;

namespace InfoPathScraper.Model
{
	using InfoPathScraper.Model.Feature;
	/// <summary>
	/// More complex derivation of InfoPathFile. In addit
	/// </summary>
	class InfoPathManifest : InfoPathFile
	{
		#region Private members
		private static XNamespace xsfNamespace = @"http://schemas.microsoft.com/office/infopath/2003/solutionDefinition";
		private static XNamespace xsf2Namespace = @"http://schemas.microsoft.com/office/infopath/2006/solutionDefinition/extensions";
		private static XNamespace xsf3Namespace = @"http://schemas.microsoft.com/office/infopath/2009/solutionDefinition/extensions";
		private const string viewNode = @"mainpane";
		private const string viewNameAttribute = @"transform";
		private const string xDocumentClass = @"xDocumentClass";
		private const string uniqueUrn = @"name";

		private XElement _xDocumentNode;
		private List<string> _viewNames;
		private InfoPathManifest() { }
		#endregion

		#region Public stuff
		public List<string> ViewNames { get { InitializeViewNames(); return _viewNames; } }
        public string Name 
        { 
            get 
            { 
                InitializeXDocumentNode(); 

                //InfoPath 2003 forms don't have a Name attribute, so return an empty string instead of throwing an exception.
                return (_xDocumentNode.Attribute(uniqueUrn) == null) ? string.Empty : _xDocumentNode.Attribute(uniqueUrn).Value; 
            } 
        }
		public static InfoPathManifest Create(CabFileInfo cabFileInfo)
		{
			InfoPathManifest manifest = new InfoPathManifest();
			manifest.CabFileInfo = cabFileInfo;
			return manifest;
		}
		#endregion

		#region Override implementations
		protected override IEnumerable<Func<XDocument, IEnumerable<InfoPathFeature>>> FeatureDiscoverers
		{
			get
			{
				yield return Mode.ParseFeature;
				yield return PublishUrl.ParseFeature;
				yield return DataConnection.ParseFeature;
				yield return ManagedCode.ParseFeature;
                yield return ProductVersion.ParseFeature;
				yield return DataRule.ParseFeature;
				yield return DataValidation.ParseFeature;
				yield break;
			}
		}

		#endregion

		#region Private helpers
		/// <summary>
		/// This is the root node of the manifest.xsf file from which get a few interesting properties.
		/// </summary>
		private void InitializeXDocumentNode()
		{
			if (_xDocumentNode != null) return;
			IEnumerable<XElement> elements = XDocument.Descendants(xsfNamespace + xDocumentClass);
			foreach (XElement element in elements)
			{
				if (_xDocumentNode != null) throw new ArgumentException("Manifest has multiple xDocumentClass nodes");
				_xDocumentNode = element;
			}
		}

		/// <summary>
		/// An xsn can have many resource files in it. The only reliable way to know which ones are views is to 
		/// parse the manifest for those that are called out as such. We just need string names of them because 
		/// the Microsoft.Deployment.Compression.Cab code will happily find them by name.
		/// </summary>
		private void InitializeViewNames()
		{
			if (_viewNames != null) return;
			_viewNames = new List<string>();

			foreach (XElement mainpane in XDocument.Descendants(xsfNamespace + viewNode))
			{
				_viewNames.Add(mainpane.Attribute(viewNameAttribute).Value);
			}
		}
		#endregion

	}
}
