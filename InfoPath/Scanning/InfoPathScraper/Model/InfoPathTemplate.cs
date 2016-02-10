using System;
using System.Collections.Generic;
using System.Xml.Linq;
using Microsoft.Deployment.Compression.Cab;

namespace InfoPathScraper.Model
{
	using InfoPathScraper.Model.Feature;
	class InfoPathTemplate
	{
		#region Members and Basics
		public CabInfo CabInfo { get; private set; }
		public List<InfoPathView> _infoPathViews;
		private InfoPathManifest _infoPathManifest;

		private InfoPathTemplate() { }

		public static InfoPathTemplate CreateTemplate(string path)
		{
			// Lazy Init
			CabInfo cabInfo = new CabInfo(path);
			InfoPathTemplate template = new InfoPathTemplate();
			template.CabInfo = cabInfo;
			return template;
		}
		#endregion


		#region Various Properties we should support
		public List<InfoPathView> InfoPathViews { get { InitializeViews(); return _infoPathViews; } }
		public InfoPathManifest InfoPathManifest { get { InitializeManifest(); return _infoPathManifest; } }
		public IEnumerable<InfoPathFile> FeaturedFiles { get { yield return InfoPathManifest; foreach (InfoPathView view in InfoPathViews) yield return view; yield break; } }
		public IEnumerable<InfoPathFeature> Features { get { foreach (InfoPathFile file in FeaturedFiles) { foreach (InfoPathFeature feature in file.Features) yield return feature; } yield break; } }
		#endregion

		#region Private helpers to compute things
		private void InitializeManifest()
		{
			if (_infoPathManifest != null) return;

			// get the files named manifest.xsf (there should be one)
			IList<CabFileInfo> cbInfos = CabInfo.GetFiles("manifest.xsf");
			if (cbInfos.Count != 1) throw new ArgumentException("Invalid InfoPath xsn");
			_infoPathManifest = InfoPathManifest.Create(cbInfos[0]);
		}

		private void InitializeViews()
		{
			if (_infoPathViews != null) return;
			_infoPathViews = new List<InfoPathView>();

			foreach (string name in InfoPathManifest.ViewNames)
			{
				IList<CabFileInfo> cbInfos = CabInfo.GetFiles(name);
				if (cbInfos.Count != 1) throw new ArgumentException(String.Format("Malformed template file: view {0} not found", name));
				InfoPathView viewFile = InfoPathView.Create(cbInfos[0]);
				_infoPathViews.Add(viewFile);
			}
		}
		#endregion
	}
}
