using System;
using System.Text;

namespace InfoPathScraper.Reporting
{
	/// <summary>
	/// Outputs a comma-separated-value format, one line per feature:
	///    urn,feature,{feature details}
	/// The idea is to be able to consume this directly in Excel and do whatever filtering, reporting, etc on it
	/// </summary>
	class CSVReport : Report
	{
		protected override string OutputTemplate(Model.InfoPathTemplate template)
		{
			StringBuilder sb = new StringBuilder();

            //Skip reporting on InfoPath 2003 forms that don't have a name.
            if (string.IsNullOrEmpty(template.InfoPathManifest.Name))
            {
                return sb.ToString();
            }

			sb.Append(template.InfoPathManifest.Name).Append(",").Append("CabPath").Append(",").Append(template.CabInfo.FullName).Append("\r\n");
			foreach (Model.Feature.InfoPathFeature feature in template.Features)
			{
				sb.Append(template.InfoPathManifest.Name).Append(",");
				sb.Append(feature.FeatureName).Append(",");
				sb.Append(feature.ToCSV());
				sb.Append("\r\n");
			}
			return sb.ToString();
		}
	}
}
