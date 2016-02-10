using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace InfoPathScraper.Model.Feature
{
	/// <summary>
	/// Finds all rules acting on data in an InfoPath form
	/// </summary>
	class DataRule : InfoPathFeature
	{
		#region Constants
		private const string rule = @"rule";
		#endregion

		#region Public interface
		public string ActionType { get; private set; }

		public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
		{
			IEnumerable<XElement> allRules = document.Descendants(xsfNamespace + rule);
			foreach (XElement ruleElement in allRules)
			{
				foreach (DataRule feature in ParseRuleElement(ruleElement))
					yield return feature;
			}
			yield break;
		}

		public override string ToString()
		{
			return FeatureName + ": " + ActionType;
		}

		public override string ToCSV()
		{
			return ActionType;
		}
		#endregion

		#region Private helpers
		private static IEnumerable<DataRule> ParseRuleElement(XElement ruleElement)
		{
			foreach (XElement ruleAction in ruleElement.Elements())
			{
				DataRule feature = new DataRule();
				feature.ActionType = ruleAction.Name.LocalName;
				// we can be any one of many types of rules: dialogbox, assignment, query, submit, switch view 
				// we could parse further if that turns out to be interesting
				yield return feature;
			}
		}
		#endregion
	}
}
