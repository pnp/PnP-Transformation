using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace InfoPathScraper.Model.Feature
{
	/// <summary>
	/// Finds all controls in the InfoPathViews
	/// </summary>
	class PublishUrl : InfoPathFeature
	{
		#region Private stuff
		private const string baseUrl = @"baseUrl";
		private const string relativeUrlBaseAttribute = @"relativeUrlBase";
		private const string publishUrlAttribute = @"publishUrl";
		private const string xDocumentClass = @"xDocumentClass";

		private PublishUrl() { }
		#endregion

		#region Public interface
		public string Publish { get; private set; }
		public string RelativeBase { get; private set; }

		/// <summary>
		/// Instead of logging on feature per control, I do 1 feature per control type along with the number of occurrences
		/// </summary>
		/// <param name="document"></param>
		/// <returns></returns>
		public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
		{
			PublishUrl pubRule = new PublishUrl();
			IEnumerable<XElement> allElements = document.Descendants(xsf3Namespace + baseUrl);
			// collect the control counts
			foreach (XElement element in allElements)
			{
				if (pubRule.RelativeBase != null) throw new ArgumentException("Should only see one xsf3:baseUrl node");
				XAttribute pathAttribute = element.Attribute(relativeUrlBaseAttribute);
				if (pathAttribute != null) // this attribute is technically optional per xsf3 spec
				{
					pubRule.RelativeBase = pathAttribute.Value;
				}
			}

			allElements = document.Descendants(xsfNamespace + xDocumentClass);
			foreach (XElement element in allElements)
			{
				if (pubRule.Publish != null) throw new ArgumentException("Should only see one xsf:xDocumentClass node");
				XAttribute pubUrl = element.Attribute(publishUrlAttribute);
				if (pubUrl != null)
				{
					pubRule.Publish = pubUrl.Value; 
				}
			}

			yield return pubRule;
			// nothing left
			yield break;
		}

		public override string ToString()
		{
			return FeatureName + ": RelativeBase=" + RelativeBase + ", PublishUrl=" + Publish;
		}

		public override string ToCSV()
		{
			return RelativeBase + "," + Publish;
		}
		#endregion
	}
}
