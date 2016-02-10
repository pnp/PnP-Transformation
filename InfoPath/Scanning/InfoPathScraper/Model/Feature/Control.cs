using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace InfoPathScraper.Model.Feature
{
	/// <summary>
	/// Finds all controls in the InfoPathViews
	/// </summary>
	class Control : InfoPathFeature
	{
		#region Private stuff
		private const string xctName = @"xctname";

		private Control() { }
		#endregion

		#region Public interface
		public string Name { get; private set; }
		public int Count { get; private set; }

		/// <summary>
		/// Instead of logging on feature per control, I do 1 feature per control type along with the number of occurrences
		/// </summary>
		/// <param name="document"></param>
		/// <returns></returns>
		public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
		{
			IEnumerable<XElement> allElements = document.Descendants();
			Utilities.BucketCounter counter = new Utilities.BucketCounter();
			// collect the control counts
			foreach (XElement element in allElements)
			{
				XAttribute xctAttribute = element.Attribute(xdNamespace + xctName);
				if (xctAttribute != null)
				{
					counter.IncrementKey(xctAttribute.Value);
				}
			}

			// then create Control objects for each control
			foreach (KeyValuePair<string, int> kvp in counter.Buckets)
			{
				Control c = new Control();
				c.Name = kvp.Key;
				c.Count = kvp.Value;
				yield return c;
			}
			// nothing left
			yield break;
		}

		public override string ToString()
		{
			return FeatureName + ": " + Name + "[" + Count + "]";
		}

		public override string ToCSV()
		{
			return Name + "," + Count;
		}
		#endregion
	}
}
