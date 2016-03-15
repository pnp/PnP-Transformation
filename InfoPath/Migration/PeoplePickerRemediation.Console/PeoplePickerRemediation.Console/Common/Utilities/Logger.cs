using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace PeoplePickerRemediation.Console.Common.Utilities
{
    public class Logger
    {
        public static string LoggerFileName { get; set; }

        #region CreateLogFile Method

        /// <summary>
        ///     This method is used to create Log File.
        ///     It reterives the default path from Constants file and suffice with current date.
        ///     The TraceLog output file will be in TEXT format
        /// </summary>
        public string CreateLogFile(string folderPath)
        {
            string strLogFolderPath = folderPath;

            if (!Directory.Exists(strLogFolderPath))
            {
                Directory.CreateDirectory(strLogFolderPath);
            }

            string strLogFileName = Constants.TraceLogFileSuffix + "_" + DateTime.Now.ToString("dd_MM_yyyy_hh_mm_ss") + ".txt";
            string strFileNameWithPath = strLogFolderPath + @"\" + strLogFileName;

            if (!(File.Exists(strFileNameWithPath)))
            {
                using (FileStream fStream = new FileStream(strFileNameWithPath, FileMode.Create))
                {
                    //fStream.Close();
                }
            }
            //Setting Log File Name and Path
            LoggerFileName = strFileNameWithPath;

            return strFileNameWithPath;
        }

        public string CreateLogFile(string folderPath, string fileName)
        {
            string strLogFolderPath = folderPath;

            if (!Directory.Exists(strLogFolderPath))
            {
                Directory.CreateDirectory(strLogFolderPath);
            }

            string strFileNameWithPath = strLogFolderPath + @"\" + fileName;

            if (!(File.Exists(strFileNameWithPath)))
            {
                using (FileStream fStream = new FileStream(strFileNameWithPath, FileMode.Create))
                {
                    //fStream.Close();
                }
            }
            //Setting Log File Name and Path
            LoggerFileName = strFileNameWithPath;

            return strFileNameWithPath;
        }

        #endregion

        public static void AddMessageToTraceLogFile(bool logging, string traceLogMessage = "N/A")
        {
            if (logging)
            {
                StringBuilder strbErrMsg = new StringBuilder();
                try
                {
                    using (StreamWriter sw = new StreamWriter(LoggerFileName, true, System.Text.Encoding.UTF8))
                    {
                        strbErrMsg.Append(traceLogMessage + Environment.NewLine);
                        sw.Write(strbErrMsg.ToString());
                        //sw.Close();
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        public static string CurrentDateTime()
        {
            return DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
        }
        public static void Write_TraceLog_AND_ConsoleMessage(string Message, bool IsDateTime, bool loggingOnConsole, ConsoleColor foregroundColor)
        {
            Logger.AddMessageToTraceLogFile(Constants.Logging, Message);

            string DateTime = Logger.CurrentDateTime();
            if (IsDateTime)
            {
                Logger.AddMessageToTraceLogFile(Constants.Logging, "[DATE TIME] " + DateTime);
            }

            if (loggingOnConsole)
            {
                System.Console.ForegroundColor = foregroundColor;

                System.Console.WriteLine(Message);
                
                if (IsDateTime)
                {
                    System.Console.WriteLine("[DATE TIME] " + DateTime);
                }

                System.Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        #region --------------------------------- Singleton implementation ---------------------------------

        private static Logger _logger;

        public static Logger CurrentInstance
        {
            get
            {
                if (_logger == null)
                {
                    _logger = new Logger();
                }

                return _logger;
            }
        }

        #endregion
    }
}
