using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace InfoPathScraper.Model.Feature
{
	/// <summary>
	/// Finds if any managed code is being used in a form.
	/// </summary>
	class ManagedCode : InfoPathFeature
	{
        private const string enabledAttribute = @"enabled";
        private const string languageAttribute = @"language";
		private const string versionAttribute = @"version";
		private const string managedCode = @"managedCode";

		public string Language { get; private set; }
		public string Version { get; private set; }

		public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
		{
			IEnumerable<XElement> allElements = document.Descendants(xsf2Namespace + managedCode);
			foreach (XElement element in allElements)
			{ 
                //The enabled attribute is typically null, so check for that. 
                //If the attribute exists, then we can also check the value to make sure there's not an enabled="no" scenario.                
                if(element.Attribute(enabledAttribute) == null 
                    || !string.Equals(element.Attribute(enabledAttribute).Value, "yes", StringComparison.InvariantCultureIgnoreCase))
                {
                    continue;
                }

                //Return an object if enabled="yes"
                ManagedCode mc = new ManagedCode();
                mc.Language = element.Attribute(languageAttribute).Value;
                mc.Version = element.Attribute(versionAttribute).Value;
                yield return mc;
			}

			// nothing left
			yield break;
		}

		public override string ToString()
		{
			return FeatureName + ": " + Language + " " + Version;
		}

		public override string ToCSV()
		{
			return Language + "," + Version;
		}
	}
}
