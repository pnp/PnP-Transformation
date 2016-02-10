using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace InfoPathScraper.Model.Feature
{
	/// <summary>
	/// Finds if any managed code is being used in a form.
	/// </summary>
	class ProductVersion : InfoPathFeature
	{
        private const string productVersionAttribute = @"productVersion";
        		
		public string Version { get; private set; }

		public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
		{
            //ProductVersion is off the root node, so grab that value and return the value.
            XAttribute attribute = document.Root.Attribute(productVersionAttribute);

			if(attribute != null)
            {                
                ProductVersion pv = new ProductVersion();
                pv.Version = attribute.Value;
                yield return pv;
            }

			// nothing left
			yield break;
		}

		public override string ToString()
		{
			return FeatureName + ": " + Version;
		}

		public override string ToCSV()
		{
			return Version;
		}
	}
}
