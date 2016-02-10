using System;
using System.Collections.Generic;
using System.Xml.Linq;
using Microsoft.Deployment.Compression.Cab;

namespace InfoPathScraper.Model
{
	using InfoPathScraper.Model.Feature;
	/// <summary>
	/// Self-explanatory implementation
	/// </summary>
	class InfoPathView : InfoPathFile
	{
		#region Private stuff
		private InfoPathView() { }
		#endregion

		#region Public stuff
		public static InfoPathView Create(CabFileInfo cabFileInfo)
		{
			InfoPathView view = new InfoPathView();
			view.CabFileInfo = cabFileInfo;
			return view;
		}
		#endregion

		#region Override implementations
		protected override IEnumerable<Func<XDocument, IEnumerable<InfoPathFeature>>> FeatureDiscoverers
		{
			get
			{
				yield return Control.ParseFeature;
				yield return FormattingRule.ParseFeature;
				yield break;
			}
		}
		#endregion

	}
}
