using System;
using System.IO;
using System.Text;

namespace Transformation.PowerShell.Common.Utilities
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

            string strLogFileName = Constants.TraceLogFileSuffix + "_" + DateTime.Now.ToString("dd_MM_yyyy") + ".txt";
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

        #endregion

        public static void AddMessageToTraceLogFile(bool logging, string traceLogMessage = "N/A")
        {
            if (logging)
            {
                StringBuilder strbErrMsg = new StringBuilder();
                try
                {
                    using (StreamWriter sw = new StreamWriter(LoggerFileName, true))
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