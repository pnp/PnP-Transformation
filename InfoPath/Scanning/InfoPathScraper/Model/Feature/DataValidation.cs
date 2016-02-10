using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace InfoPathScraper.Model.Feature
{
	/// <summary>
	/// Finds all data validations in an InfoPath form
	/// </summary>
	class DataValidation : InfoPathFeature
	{
		#region Constants
		private const string customValidation = @"customValidation";
		#endregion

		#region Public interface
		public string ValidationType { get; private set; }

		public static IEnumerable<InfoPathFeature> ParseFeature(XDocument document)
		{
			// we don't care about the condition details, just "any custom validation" vs "native cbb"
			IEnumerable<XElement> allValidations = document.Descendants(xsfNamespace + customValidation);
			foreach (XElement validationElement in allValidations)
			{
				DataValidation validation = new DataValidation();
				validation.ValidationType = "Custom validation";
				yield return validation;
			}

			allValidations = document.Descendants(xsf3Namespace + customValidation);
			foreach (XElement validationElement in allValidations)
			{
				DataValidation validation = new DataValidation();
				validation.ValidationType = "Cannot be blank";
				yield return validation;
			}

			yield break;
		}

		public override string ToString()
		{
			return FeatureName + ": " + ValidationType;
		}

		public override string ToCSV()
		{
			return ValidationType;
		}
		#endregion
	}
}
