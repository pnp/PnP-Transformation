using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;

namespace InfoPathScraper.Model.Feature
{
	/// <summary>
	/// Finds view-level formatting rules. These are a little tricky because they are 
	/// implemented by sprinkling xsl logic into our views. As such, a particular formatting rule
	/// could show up in multiple places, or merge-ish with other rules.
	/// </summary>
	class FormattingRule : InfoPathFeature
	{
		#region Constants
		private const string xslAttribute = @"attribute";
		private const string nameAttribute = @"name";
		private const string contentEditable = @"contentEditable"; // read-only
		private const string style = @"style"; // adjusting the look and feel (colors, hiding, etc ...)
		private const string when = @"when";
		#endregion

		#region Public interface
		public string FormatType { get; private set; }
		public string SubDetails { get; private set; }

		public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
		{
			IEnumerable<XElement> allElements = document.Descendants(xslNamespace + xslAttribute);
			foreach (XElement element in allElements)
			{
				XAttribute name = element.Attribute(nameAttribute);
				// these are the html attributes that we try to set. 
				// specifically for conditional hide we have to look under a style for an xsl:when with .Text() contains "DISPLAY: none"
				if (name.Value.Equals(contentEditable))
				{
					FormattingRule rule = new FormattingRule();
					rule.FormatType = "Readonly";
					yield return rule;
				}
				else if (name.Value.Equals(style))
				{
					Utilities.BucketCounter counter = new Utilities.BucketCounter();
					FormattingRule rule = new FormattingRule();
					rule.FormatType = "Style";
					// now let's count all the things we're affecting. 
					// Overloading BucketCounter to filter the noise of multiple touches to same style
					foreach (XElement xslWhen in element.Descendants(xslNamespace + when))
					{
						string[] styles = xslWhen.Value.Split(new char[] { ';' });
						foreach (string s in styles)
						{
							if (s.Trim().StartsWith("caption:")) continue;
							string affectedStyle = s.Split(':')[0].Trim().ToUpper();
							counter.IncrementKey(affectedStyle);
						}
					}
					StringBuilder sb = new StringBuilder();
					foreach (KeyValuePair<string, int> kvp in counter.Buckets)
					{
						sb.Append(kvp.Key).Append(" ");
					}
					rule.SubDetails = sb.ToString().Trim();
					yield return rule;
				}
			}

			// nothing left
			yield break;
		}

		public override string ToString()
		{
			return FeatureName + ": " + FormatType + (SubDetails == null ? "" : " " + SubDetails);
		}

		public override string ToCSV()
		{
			return FormatType + (SubDetails == null ? "" : "," + SubDetails);
		}
		#endregion
	}
}
