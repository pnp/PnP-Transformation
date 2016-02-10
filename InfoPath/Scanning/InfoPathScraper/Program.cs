using System;


namespace InfoPathScraper
{
	using InfoPathScraper.Model;
	using InfoPathScraper.Model.Feature;
	class Program
	{
		/// <summary>
		/// Entrypoint: takes command line args. See CommandLineProcessor.Usage for details.
		/// </summary>
		/// <param name="args"></param>
		static void Main(string[] args)
		{
			Reporting.CommandLineProcessor processor = new Reporting.CommandLineProcessor();
			try
			{
				processor.ProcessArguments(args);
				processor.Report.Run();
			}
			catch (Exception e)
			{
				// if any Exception is thrown, show that to the user and also show usage.
				Console.Error.WriteLine(e.Message);
				if (e.InnerException != null)
					Console.Error.WriteLine(e.InnerException.Message);
				Console.Error.WriteLine();
				foreach (string s in processor.Usage)
					Console.Error.WriteLine(s);
			}

		}
	}
}
