using System;
using System.Text;

namespace InfoPathScraper.Reporting
{
	/// <summary>
	/// Report formatted for better human readability. 
	/// This is the more likely version you'd want to use to see what a single xsn has in it
	/// </summary>
	class TextReport : Report
	{
		protected override string OutputTemplate(Model.InfoPathTemplate template)
		{
			StringBuilder sb = new StringBuilder();
            
            //Skip reporting on InfoPath 2003 forms that don't have a name.
            if (string.IsNullOrEmpty(template.InfoPathManifest.Name))
            {
                return sb.ToString();
            }

			sb.Append(template.InfoPathManifest.Name).Append("\r\n");
			foreach (Model.Feature.InfoPathFeature feature in template.Features)
			{
				sb.Append(feature.ToString());
				sb.Append("\r\n");
			}
			return sb.ToString();
		}
	}
}
