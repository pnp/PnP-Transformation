using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace InfoPathScraper.Model.Feature
{
	/// <summary>
	/// Finds the mode of the form
	/// </summary>
	class Mode : InfoPathFeature
	{
		private const string modeAttribute = @"mode";
		private const string solutionFormatVersionAttribute = @"solutionFormatVersion";
		private const string solutionMode = @"solutionMode";
		private const string solutionDefinition = @"solutionDefinition";
		private const string solutionPropertiesExtension = @"solutionPropertiesExtension";
		private const string branchAttribute = @"branch";
		private const string runtimeCompatibilityAttribute = @"runtimeCompatibility";
		private const string xDocumentClass = @"xDocumentClass";

		public string ModeName { get; private set; }
		public string Compatibility { get; private set; }

		private Mode() { }

		public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
		{

			Mode m = new Mode();
			string mode = null;

			// look for fancy new modes first (these were new for xsf3 / IP2010)
			IEnumerable<XElement> allModeElements = document.Descendants(xsf3Namespace + solutionMode);
			foreach (XElement element in allModeElements)
			{
				if (mode != null) throw new ArgumentException("Found more than one mode!");
				XAttribute name = element.Attribute(modeAttribute);
				mode = name.Value;
			}

			// and if we didn't find the above, fall back to client v server in xsf2:solutionDefinition
			if (mode == null)
			{
				IEnumerable<XElement> allSolutionDefs = document.Descendants(xsf2Namespace + solutionDefinition);
				foreach (XElement solutionDef in allSolutionDefs)
				{
					if (mode != null) throw new ArgumentException("Found more than one xsf2:solutionDefition!");
					XElement extension = solutionDef.Element(xsf2Namespace + solutionPropertiesExtension);
					if (extension != null && extension.Attribute(branchAttribute) != null && extension.Attribute(branchAttribute).Equals("contentType"))
					{
						mode = "Document Information Panel";
					}
					else
					{
						XAttribute compat = solutionDef.Attribute(runtimeCompatibilityAttribute);
						mode = compat.Value;
					}
				}
			}

			// and if we still found nothing, it's a 2003 form and must be client:
			if (mode == null)
				mode = "client";

			m.ModeName = mode;

			string compatibility = null;
			foreach (XElement xDoc in document.Descendants(xsfNamespace + xDocumentClass))
			{
				if (compatibility != null) throw new ArgumentException("Multiple xDocumentClass nodes found!");
				compatibility = xDoc.Attribute(solutionFormatVersionAttribute).Value;
			}
			m.Compatibility = compatibility;

			yield return m;
			yield break;
		}


		public override string ToString()
		{
			return FeatureName + ": " + ModeName + " " + Compatibility;
		}

		public override string ToCSV()
		{
			return ModeName + "," + Compatibility;
		}
	}
}
