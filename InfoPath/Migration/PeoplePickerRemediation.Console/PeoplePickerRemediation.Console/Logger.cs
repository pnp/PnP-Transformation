using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PeoplePickerRemediation.Console
{
    public static class Logger
    {
        private static System.IO.StreamWriter logFile = null;

        static Logger()
        {
        }

        public static void OpenLog(string operationName)
        {
            if (logFile == null)
            {
                if (String.IsNullOrEmpty(operationName))
                {
                    ConsoleMessage("Task-specific Log File not present; using default Log File instead...");
                    operationName = "TransformationConsole";
                }
                string logFileSpecFormat = "{0}-" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".log";
                string logFileSpec = String.Format(logFileSpecFormat, operationName);
                logFile = new System.IO.StreamWriter(logFileSpec);
                logFile.AutoFlush = true;
            }
        }

        public static void CloseLog()
        {
            if (logFile != null)
            {
                logFile.Close();
                logFile.Dispose();
                logFile = null;
            }
        }

        public static void LogErrorMessage(string msg, bool toConsole = true)
        {
            LogMessage("ERROR: " + msg, toConsole);
        }
        public static void LogWarningMessage(string msg, bool toConsole = true)
        {
            LogMessage("WARNING: " + msg, toConsole);
        }
        public static void LogInfoMessage(string msg, bool toConsole = true)
        {
            LogMessage("INFO: " + msg, toConsole);
        }

        public static void LogSuccessMessage(string msg, bool toConsole = true)
        {
            LogMessage("SUCCESS: " + msg, toConsole);
        }

        public static void LogMessage(string msg, bool toConsole = true)
        {
            if (toConsole)
            {
                ConsoleMessage(msg);
            }

            if (logFile == null)
            {
                OpenLog(String.Empty);
            }

            try
            {
                logFile.WriteLine(msg);
            }
            catch
            {
                // Echo to Console if any file issues occur...
                ConsoleMessage(msg);
            }
        }

        public static void ConsoleMessage(string msg)
        {
            System.Console.WriteLine(msg);
        }

    }
}
