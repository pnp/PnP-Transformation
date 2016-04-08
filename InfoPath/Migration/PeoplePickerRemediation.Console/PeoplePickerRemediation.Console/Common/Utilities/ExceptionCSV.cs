using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;

namespace PeoplePickerRemediation.Console.Common.Utilities
{
    public class ExceptionCsv
    {
        private static ExceptionCsv _exceptionHandler;

        public static ExceptionCsv CurrentInstance
        {
            get
            {
                if (_exceptionHandler == null)
                {
                    _exceptionHandler = new ExceptionCsv();
                }

                return _exceptionHandler;
            }
        }

        public static string WebApplication { get; set; }
        public static string ContentDatabase { get; set; }
        public static string SiteCollection { get; set; }
        public static string WebUrl { get; set; }
        public static string Exceptionfilename { get; set; }
        public string MethodName { get; set; }
        public string ElementType { get; set; }
        public string ExceptionType { get; set; }
        public string ExceptionMessage { get; set; }
        public string ExceptionDetail { get; set; }
        public string ExceptionComments { get; set; }

        public bool CreateLogFile(string strLogFolderPath)
        {
            FileUtility.DeleteSpecificFileExtension(strLogFolderPath, "*archive*.csv");
            return CreateExceptionFile(strLogFolderPath);
        }

        public static void WriteException(string webapplication = "N/A", string ContentDatabase = "N/A", string sitecollection = "N/A",
            string weburl = "N/A", string elementType = "N/A", string exceptionmessage = "N/A",
            string exceptiondetail = "N/A", string methodname = "N/A", string exceptionType = "N/A",
            string exceptionComments = "N/A")
        {
            string strCsv = StringUtility.Escape(webapplication) + Constants.CsvDelimeter;
            strCsv = strCsv + StringUtility.Escape(ContentDatabase) + Constants.CsvDelimeter;
            strCsv = strCsv + StringUtility.Escape(sitecollection) + Constants.CsvDelimeter;
            strCsv = strCsv + StringUtility.Escape(weburl) + Constants.CsvDelimeter;
            strCsv = strCsv + StringUtility.Escape(elementType) + Constants.CsvDelimeter;
            strCsv = strCsv + StringUtility.Escape(exceptionmessage) + Constants.CsvDelimeter;
            strCsv = strCsv + StringUtility.Escape(exceptiondetail) + Constants.CsvDelimeter;
            strCsv = strCsv + StringUtility.Escape(methodname) + Constants.CsvDelimeter;
            strCsv = strCsv + StringUtility.Escape(exceptionType) + Constants.CsvDelimeter;
            strCsv = strCsv + StringUtility.Escape(exceptionComments);

            strCsv = strCsv.Replace("\r\n", string.Empty);


            DoPeriodicFlush(Exceptionfilename);
            FileUtility.AppendTextintoFile(Exceptionfilename, strCsv);
        }

        internal static bool IsFatal(Exception exception)
        {
            while (exception != null)
            {
                if (exception is OutOfMemoryException || exception is SEHException ||
                    exception is AccessViolationException || exception is ThreadAbortException)
                {
                    return true;
                }
                exception = exception.InnerException;
            }
            return false;
        }

        private static void DoPeriodicFlush(string sourceFileName)
        {
            try
            {
                if (FileUtility.FileSizeinKb(Exceptionfilename) > Constants.ExceptionFileSizeinKb)
                {
                    string directoryname = Path.GetDirectoryName(sourceFileName);
                    string fileextension = Path.GetExtension(sourceFileName);
                    string newFilename = directoryname + @"\" + Path.GetFileNameWithoutExtension(sourceFileName) +
                                      "Archive" + DateTime.Now.ToString("MMddyyyy_hhmm") + fileextension;

                    File.Move(sourceFileName, newFilename);
                    CreateExceptionFile(directoryname);
                }
            }
            catch
            {
            }
        }

        private static bool CreateExceptionFile(string folderPath)
        {
            try
            {
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                string strFileNameWithPath = folderPath + @"\" + Constants.Exception;
                if (!(File.Exists(strFileNameWithPath)))
                {
                    using (FileStream fStream = new FileStream(strFileNameWithPath, FileMode.Create))
                    {
                        fStream.Close();
                    }
                    Exceptionfilename = strFileNameWithPath;
                }
                else
                {
                    File.Delete(strFileNameWithPath);

                    using (FileStream fStream = new FileStream(strFileNameWithPath, FileMode.Create))
                    {
                        fStream.Close();
                    }
                    Exceptionfilename = strFileNameWithPath;
                }
                WriteException("WebApplication","ContentDatabase", "SiteCollection", "WebURL", "ElementType", "ExceptionMessage",
                    "ExceptionDetail", "MethodName", "ExceptionType", "ExceptionComments");
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
