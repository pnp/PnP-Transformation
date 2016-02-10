using System;
using System.Collections.Generic;

namespace InfoPathScraper.Reporting
{
	// QaD implementation that is non-concurrent and 3.5 compat.
	partial class Report
	{
		private Queue<string> _processed = new Queue<string>();
		/// <summary>
		/// Simplified version of Run that doesn't do anything concurrent
		/// </summary>
		public void Run()
		{
			// while we have elements, process them
			while (Remaining > 0)
			{
				TaskBody(_unprocessed.Dequeue());
			}
		}
	}
}
