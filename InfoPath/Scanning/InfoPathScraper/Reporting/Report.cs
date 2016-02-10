using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace InfoPathScraper.Reporting
{
	abstract partial class Report
	{
		private Object _writeLock;
		private Queue<string> _unprocessed;

		#region Public stuff
		public Report()
		{
			_unprocessed = new Queue<string>();
			_writeLock = new Object();
		}

		// these could be made public if we ever get on a UI and we need progress info
		private int Remaining { get { return _unprocessed.Count; } }
		// no need for public void Pause/Unpause, that'd only make sense for UI, which we won't have

		/// <summary>
		/// Assumes all AddPath are called before Run(). This is enforced by CommandLineProcessor. 
		/// If we ever get on a UI where additional files can be added after Run() is called, there 
		/// will need to be a mechanism to ensure that we don't add after Run() has completed and 
		/// not run any extra paths.
		/// </summary>
		/// <param name="path"></param>
		public void AddPath(string path)
		{
			_unprocessed.Enqueue(path);
		}

		public FileStream Output { get; set; }
		#endregion

		#region Abstract method(s)
		/// <summary>
		/// A string representation of the report for a single template. 
		/// Deriving classes will collate template.Features in a way that makes sense for the type of report
		/// </summary>
		/// <param name="template"></param>
		/// <returns></returns>
		protected abstract string OutputTemplate(Model.InfoPathTemplate template);
		#endregion

		#region Private helpers
		/// <summary>
		/// Create a Task to load and run based on the path
		/// </summary>
		/// <param name="path"></param>
		/// <returns></returns>
		private void TaskBody(string path)
		{
			try
			{
				// Create the template and attempt to get the report for it
				// If anything goes wrong, we'll catch it below and push the Exception to the command line output
				Model.InfoPathTemplate template = Model.InfoPathTemplate.CreateTemplate(path);
				string output = OutputTemplate(template);
				// write to Console if no output file given
				if (Output == null)
				{
					Console.Write(output);
				}
				else
				{
					lock (_writeLock)
					{
						StreamWriter writer = new StreamWriter(Output);
						writer.Write(output);
						writer.Flush();
					}
				}
			}
			catch (Exception e)
			{
				Console.Error.WriteLine("Error processing " + path + ": " + e.Message);
			}
			finally
			{
				_processed.Enqueue(path);
			}
		}
		#endregion
	}
}
