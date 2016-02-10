using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace InfoPathScraper.Model.Feature
{
	/// <summary>
	/// This is the baseclass for all of our features that we look for. 
	/// </summary>
	abstract class InfoPathFeature
	{
		#region Useful XNamespace values
		protected static XNamespace xdNamespace = @"http://schemas.microsoft.com/office/infopath/2003";
		protected static XNamespace xsfNamespace = @"http://schemas.microsoft.com/office/infopath/2003/solutionDefinition";
		protected static XNamespace xsf2Namespace = @"http://schemas.microsoft.com/office/infopath/2006/solutionDefinition/extensions";
		protected static XNamespace xsf3Namespace = @"http://schemas.microsoft.com/office/infopath/2009/solutionDefinition/extensions";
		protected static XNamespace xslNamespace = @"http://www.w3.org/1999/XSL/Transform";
		#endregion

		#region Public interface
		public string FeatureName { get { return this.GetType().Name; } }
		// need virtuals for formatting as a string, and an xml, or something ...
		public override string ToString()
		{
			return FeatureName;
		}
		#endregion

		#region Abstract method(s)
		/// <summary>
		/// This method returns a comma-separated list of the interesting values a particular feature has collected.
		/// </summary>
		/// <returns></returns>
		public abstract string ToCSV();
		#endregion
	}
}
