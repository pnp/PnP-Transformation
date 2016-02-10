using System;
using System.Collections.Generic;
using System.IO;

namespace InfoPathScraper.Reporting
{
	// takes the args from Main. Processes, stores, runs. 
	class CommandLineProcessor
	{
		#region Definitions and constants
		private enum ProcessorState
		{
			Default, // ~= File
			File,
			InputFile,
			WasInputFile,
			OutputFile,
		}

		private const string textFile = @"/text";
		private const string csvFile = @"/csv";
		private const string outFile = @"/outfile";
		private const string fileList = @"/filelist";
		private const string append = @"/append";
		private const string file = @"/file";
		#endregion

		private List<Func<string, bool>> _handlers;

		#region Properties - note these are all private on purpose!
		private ProcessorState ParseState { get; set; }
		private bool Append { get; set; }
		private List<string> Paths { get; set; }
		private List<string> PathFiles { get; set; }
		private string OutFile { get; set; }
		#endregion

		#region Public interface - not much here. Process and get_Report
		public Report Report { get; set; }

		/// <summary>
		/// Initialize collections
		/// </summary>
		public CommandLineProcessor()
		{
			ParseState = ProcessorState.Default;
			Paths = new List<string>();
			PathFiles = new List<string>();

			_handlers = new List<Func<string, bool>>();
			_handlers.Add(HandleAppend);
			_handlers.Add(HandleArgument);
			_handlers.Add(HandleCsvFile);
			_handlers.Add(HandleFile);
			_handlers.Add(HandleFileList);
			_handlers.Add(HandleOutFile);
			_handlers.Add(HandleTextFile);
		}

		/// <summary>
		/// Process each arg and maintain of state machine in ParseState.
		/// Once all are processed, check if we could possibly succeed, 
		/// then process all the info we have to fully set up our Report object.
		/// </summary>
		/// <param name="args"></param>
		public void ProcessArguments(string[] args)
		{
			foreach (string argument in args)
				ProcessArgument(argument.ToLower());

			VerifyArgState();
			ProcessArgState();
		}

		/// <summary>
		/// Self-explanatory. No strongly defensible reason to use yield here, it just looked better to my eyes.
		/// </summary>
		public IEnumerable<string> Usage
		{
			get
			{
				yield return @"CommandLineProcessor processes input arguments for handling InfoPath reports.";
				yield return @"Sample: /csv /file template1.xsn template2.xsn /filelist templatefiles.txt /outfile report.csv /append";
				yield return @"Usage: [/csv | /text] [/file file1 [file2 ...]]* [/filelist list1 [list2, ...]]* [/outfile targetfile] [/append]";
				yield return @"Params:";
				yield return @"csv:       Emit the report in comma-separated-value format";
				yield return @"text:      Emit the report in human-friendly text format (default if neither /csv nor /text is given";
				yield return @"file:      Start of a list of input InfoPath template paths (default if no other switch is active)";
				yield return @"filelist:  Start of a list of files containing input InfoPath template paths";
				yield return @"outfile:   Specify the output file (default is command line)";
				yield return @"append:    Append values to existing output file";
				yield break;
			}
		}
		#endregion

		#region Private helpers.
		/// <summary>
		/// Use the data we've collected and initialize all the info for our Report
		/// </summary>
		private void ProcessArgState()
		{
			if (OutFile != null)
			{
				// will throw an IOException if we try to open a file for !Append that already exists. This is the desired behavior.
				Report.Output = new FileStream(OutFile, Append ? FileMode.Append : FileMode.CreateNew, FileAccess.Write);
			}
			foreach (string s in Paths)
				Report.AddPath(s);
			foreach (string s in PathFiles)
			{
				// this will throw an IOException if the file doesn't exist
				StreamReader reader = new StreamReader(new FileStream(s, FileMode.Open, FileAccess.Read));
				string line = null;
				while ((line = reader.ReadLine()) != null)
				{
					// we support !... comments in the filelist files.
					line = line.Trim();
					if (line.Length > 0 && !line.StartsWith("!"))
						Report.AddPath(line.Trim());
				}
			}
		}

		/// <summary>
		/// Self-explanatory
		/// </summary>
		private void VerifyArgState()
		{
			if (Report == null)
				Report = new TextReport();
			if (Paths.Count == 0 && PathFiles.Count == 0)
				throw new ArgumentException("No input files specified.");
			if (ParseState != ProcessorState.Default && ParseState != ProcessorState.WasInputFile)
				throw new ArgumentException("Expecting more arguments based on last switch");
		}

		/// <summary>
		/// Per-argument step function to our state machine. 
		/// Each handler communicates whether or not it's suitable to handle the argument in question. 
		/// For example HandleFileList looks only for "/filelist", and returns false for all others.
		/// The assumption is that each handler knows what state to leave ParseState in.
		/// </summary>
		/// <param name="argument"></param>
		private void ProcessArgument(string argument)
		{
			try
			{
				foreach (Func<string, bool> handler in _handlers)
					if (handler(argument))
						return;
			}
			catch (Exception e)
			{
				// We want to communicate that the command line args were bad, AND what was bad about them (that's where 'e' comes in).
				throw new ArgumentException(String.Format("Invalid argument {0} encountered.", argument), e);
			}
		}

		/// <summary>
		/// Self-explanatory
		/// </summary>
		/// <param name="arg"></param>
		/// <returns></returns>
		private bool HandleTextFile(string arg)
		{
			if (!arg.Equals(textFile))
				return false;
			if (Report != null)
				throw new ArgumentException("Multiple reports requested, this is not allowed.");
			ExpectReadyForNewState(arg);
			Report = new TextReport();
			ParseState = ProcessorState.Default;
			return true;
		}

		/// <summary>
		/// Self-explanatory
		/// </summary>
		/// <param name="arg"></param>
		/// <returns></returns>
		private bool HandleCsvFile(string arg)
		{
			if (!arg.Equals(csvFile))
				return false;
			if (Report != null)
				throw new ArgumentException("Multiple reports requested, this is not allowed.");
			ExpectReadyForNewState(arg);
			Report = new CSVReport();
			ParseState = ProcessorState.Default;
			return true;
		}

		/// <summary>
		/// Self-explanatory
		/// </summary>
		/// <param name="arg"></param>
		/// <returns></returns>
		private bool HandleAppend(string arg)
		{
			if (!arg.Equals(append))
				return false;
			ExpectReadyForNewState(arg);
			Append = true;
			ParseState = ProcessorState.Default;
			return true;
		}

		/// <summary>
		/// Self-explanatory
		/// </summary>
		/// <param name="arg"></param>
		/// <returns></returns>
		private bool HandleOutFile(string arg)
		{
			if (!arg.Equals(outFile))
				return false;
			if (OutFile != null)
				throw new ArgumentException(String.Format("Already specified output file {0}.", OutFile));
			ExpectReadyForNewState(arg);
			ParseState = ProcessorState.OutputFile;
			return true;
		}

		/// <summary>
		/// Self-explanatory
		/// </summary>
		/// <param name="arg"></param>
		/// <returns></returns>
		private bool HandleFile(string arg)
		{
			if (!arg.Equals(file))
				return false;
			ExpectReadyForNewState(arg);
			ParseState = ProcessorState.File;
			return true;
		}

		/// <summary>
		/// Self-explanatory
		/// </summary>
		/// <param name="arg"></param>
		/// <returns></returns>
		private bool HandleFileList(string arg)
		{
			if (!arg.Equals(fileList))
				return false;
			ExpectReadyForNewState(arg);
			ParseState = ProcessorState.InputFile;
			return true;
		}

		/// <summary>
		/// Helper for switch processors. A switch is unexpected right after /file or /output
		/// </summary>
		/// <param name="arg"></param>
		private void ExpectReadyForNewState(string arg)
		{
			if (ParseState != ProcessorState.Default && ParseState != ProcessorState.WasInputFile)
				throw new ArgumentException(String.Format("Not expecting switch {0} at this time.", arg));
		}

		/// <summary>
		/// This is the trickiest one. It handles all non-switch arguments and must use the value of ParseState to
		/// decide what to do in each case. For example, in the default state it interprets an arg as a /file name. 
		/// In each case it needs to know if the switch can take a list or a single parameter. 
		/// Ex: '/output outfile' we need to flip back to Default state after a single arg
		/// Ex: '/file f1 f2 f3 ...' we flip back to Default state after a single arg because Default and File are handled the same
		/// Ex: '/filelist fl1 fl2 fl3 ...' we flip to WasInputFile to keep track that fl2, etc need to be treated as such, but to also be ready for other switches
		/// </summary>
		/// <param name="arg"></param>
		/// <returns></returns>
		private bool HandleArgument(string arg)
		{
			if (arg.StartsWith("/"))
				return false;
			// depending on our state, we pick what to do with the argument.
			if (ParseState == ProcessorState.InputFile || ParseState == ProcessorState.WasInputFile)
			{
				ParseState = ProcessorState.WasInputFile;
				PathFiles.Add(arg);
			}
			else if (ParseState == ProcessorState.OutputFile)
			{
				// this case should be prevented by the logic in HandleOutFile, but it's a harmless check
				if (OutFile != null)
					throw new ArgumentException(String.Format("Already specified output file: {0}."), OutFile);
				OutFile = arg;
				ParseState = ProcessorState.Default;
			}
			else if (ParseState == ProcessorState.Default || ParseState == ProcessorState.File)
			{
				ParseState = ProcessorState.Default;
				Paths.Add(arg);
			}
			return true;
		}
		#endregion
	}
}
