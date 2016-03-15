using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using PeoplePickerRemediation.Console.Common.CSV;

namespace PeoplePickerRemediation.Console.Common.Utilities
{
    public class FileUtility
    {
        public static void AppendTextintoFile(string filePath, string content)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    // Create a file to write to. 
                    using (StreamWriter sw = new StreamWriter(filePath, false, System.Text.Encoding.UTF8))
                    {
                        sw.WriteLine(content);
                    }
                }
                else
                {
                    // This text is always added, appending the text into file 
                    if (File.Exists(filePath))
                    {
                        using (StreamWriter sw = new StreamWriter(filePath, true, System.Text.Encoding.UTF8))
                        {
                            sw.WriteLine(content);
                        }
                    }
                }
            }
            catch
            {
            }
        }

        public static void WriteCsVintoFile<T>(string filepath, ref List<T> list)
        {
            if (list.Count > 0)
            {
                string strCsv = ExportCsv.ToCsv(",", list);

                AppendTextintoFile(filepath, strCsv);
            }

            list.TrimExcess();
            //list = null;
        }

        public static void WriteCsVintoFile<T>(string filepath, ref List<T> list, ref bool blHeader)
        {
            if (list.Count > 0)
            {
                string strCsv = ExportCsv.ToCsv(",", list, ref blHeader);

                AppendTextintoFile(filepath, strCsv);
            }

            list.TrimExcess();
            //list = null;
        }

        public static void WriteCsVintoFile<T>(string filepath, T list, ref bool blHeader)
        {
            if (list != null)
            {
                string strCsv = ExportCsv.ToCsv(",", list, ref blHeader);

                AppendTextintoFile(filepath, strCsv);
            }
        }

        public static void DeleteFiles(string[] filePaths)
        {
            if (filePaths != null)
            {
                foreach (string filepath in filePaths)
                {
                    try
                    {
                        if (File.Exists(filepath))
                        {
                            File.Delete(filepath);
                        }
                    }
                    catch
                    {
                    }
                }
            }
        }

        public static void DeleteFiles(string filePaths)
        {
            if (filePaths != null)
            {
                try
                {
                    if (File.Exists(filePaths))
                    {
                        File.Delete(filePaths);
                        Logger.Write_TraceLog_AND_ConsoleMessage("[DeleteFiles] ::: The file [" + filePaths + "] has been deleted", false, false, ConsoleColor.Gray);
                    }
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }
        public static void DeleteFiles(string sourceDirectory, string searchPattern = "*.*")
        {
            try
            {
                if (sourceDirectory != "")
                {
                    string[] files = FileUtility.FindAllFilewithSearchPattern(sourceDirectory, searchPattern);
                    if (files == null) throw new ArgumentNullException(sourceDirectory + " not exist");

                    foreach (var file in files)
                    {
                        File.Delete(file);
                        Logger.Write_TraceLog_AND_ConsoleMessage("[DeleteFiles] ::: The file [" + file.ToString() + "] has been deleted from path: [" + sourceDirectory+"]", false, false, ConsoleColor.Gray);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static string GetSearchPattern(string sourceDirectory, string FileName)
        {
            string searchpattern = String.Empty;
            try
            {
                searchpattern = System.IO.Path.GetFileNameWithoutExtension(sourceDirectory + "\\" + FileName);
                searchpattern = searchpattern + "*.csv";
            }
            catch (Exception)
            {
                throw;
            }

            return searchpattern;
        }

        public static void DeleteEmptyDirectories(string startLocation)
        {
            foreach (string directory in Directory.GetDirectories(startLocation))
            {
                DeleteEmptyDirectories(directory);
                if (Directory.GetFiles(directory).Length == 0 &&
                    Directory.GetDirectories(directory).Length == 0)
                {
                    Directory.Delete(directory, false);
                     System.Console.WriteLine(directory);
                }
            }
        }

        public static void CreateDirectory(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
        public static void ValidateDirectory(ref string path)
        {            
            if (path.EndsWith(@"\"))
            {
                path = path.Remove(path.Length - 1);
            }
        }
        public static void DeleteSpecificFileExtension(string sourceDirectory, string fileExtenstion)
        {
            if (Directory.Exists(sourceDirectory))
            {
                string[] files = Directory.GetFiles(sourceDirectory, fileExtenstion, SearchOption.AllDirectories);

                foreach (string s in files)
                {
                    try
                    {
                        // Use static Path methods to extract only the file name from the path.

                        string fullpath = Path.GetFullPath(s);


                        File.Delete(fullpath);
                    }
                    catch (Exception ex)
                    {
                        string exception = ex.Message;
                    }
                }
            }
        }

        public static void FindAllFileExtension(string sourceDirectory)
        {
            if (Directory.Exists(sourceDirectory))
            {
                string[] files = Directory.GetFiles(sourceDirectory, "*.*", SearchOption.AllDirectories);
                ArrayList alfileextension = new ArrayList();
                foreach (string s in files)
                {
                    try
                    {
                        // Use static Path methods to extract only the file name from the path.

                        string fileName = Path.GetFileName(s);
                        string fullpath = Path.GetFullPath(s);
                        string extension = Path.GetExtension(s);

                        if (extension != null && !alfileextension.Contains(extension))
                        {
                            alfileextension.Add(extension);
                        }
                    }
                    catch (Exception ex)
                    {
                        string exception = ex.Message;
                    }
                }
            }
        }

        public static string[] FindAllFilewithSearchPattern(string sourceDirectory, string searchPattern)
        {
            if (Directory.Exists(sourceDirectory))
            {
                return Directory.GetFiles(sourceDirectory, searchPattern, SearchOption.AllDirectories);

            }
            return null;
        }

        public static void MoveAllFileswithSearchPattern(string sourceDirectory, string searchPattern)
        {
            string[] files = Directory.GetFiles(sourceDirectory, searchPattern, SearchOption.AllDirectories);
            foreach (string file in files)
            {
                string newfilename = Path.GetFileName(file);
                if (newfilename != null)
                {
                    string newlocation = Path.Combine(sourceDirectory, newfilename);

                    if (file != newlocation)
                    {
                        if (File.Exists(newlocation))
                        {
                            File.Delete(newlocation);
                        }
                        File.Move(file, newlocation);
                    }
                }
            }
        }

        public static void FindFilewithSpecificExtension(string sourceDirectory, string fileExtension)
        {
            if (Directory.Exists(sourceDirectory))
            {
                string[] files = Directory.GetFiles(sourceDirectory, fileExtension, SearchOption.AllDirectories);
                ArrayList alfileextension = new ArrayList();
                foreach (string s in files)
                {
                    try
                    {
                        // Use static Path methods to extract only the file name from the path.

                        string fileName = Path.GetFileName(s);
                        string fullpath = Path.GetFullPath(s);
                        string extension = Path.GetExtension(s);
                    }
                    catch (Exception ex)
                    {
                        string exception = ex.Message;
                    }
                }
            }
        }

        public static long FileSizeinKb(string fileName)
        {
            FileInfo fi = new FileInfo(fileName);
            return fi.Length / 1024;
        }

        /*
        public static bool DoPeriodicFlush(string filePath, long filesize)
        {
            try
            {
                if (filesize < 1)
                {
                    filesize = Constants.ExceptionFileSizeinKb;
                }
                if (FileSizeinKb(filePath) > filesize)
                {
                    string directoryname = Path.GetDirectoryName(filePath);
                    string fileextension = Path.GetExtension(filePath);
                    string newFilename = directoryname + @"\" + Path.GetFileNameWithoutExtension(filePath) +
                                      "Archive" + DateTime.Now.ToString("MMddyyyy_hhmmss") + fileextension;

                    File.Move(filePath, newFilename);
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }
        */

        public static void DoPeriodicFlush(string sourceFileName, ref bool blHeader)
        {
            try
            {
                //10 MB File=10240 KB
                if (!File.Exists(sourceFileName) || FileUtility.FileSizeinKb(sourceFileName) <= Constants.OutputFileSizeinKb) return;

                string directoryname = Path.GetDirectoryName(sourceFileName);
                string fileextension = Path.GetExtension(sourceFileName);
                string newFilename = directoryname + @"\" + Path.GetFileNameWithoutExtension(sourceFileName) +
                                     "_" + DateTime.Now.ToString("MMddyyyy_hhmmss") + fileextension;

                if (File.Exists(newFilename)) return;

                File.Move(sourceFileName, newFilename);
                //Making Header False to write Header in CSV
                blHeader = false;
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(ExceptionCsv.WebApplication, ExceptionCsv.SiteCollection, ExceptionCsv.WebUrl, "DoPeriodicFlush", ex.Message, ex.ToString(), "DoPeriodicFlush", ex.GetType().ToString(), Constants.NotApplicable);
            }
        }
        public static void DoPeriodicFlushOfListObject<T>(ref List<T> lstObjBase, string sourceFileName, ref bool blHeader)
        {
            int numberOfRecordsExported = 0;
            try
            {
                //186229
                if (lstObjBase.Count > Constants.MaxListRecordsToExportCountForPeriodic)
                {
                    int counter = (lstObjBase.Count / Constants.MaxListRecordsToExportCountForPeriodic);
                    int startIndex = 0;
                    for (int i = 0; i < counter; i++)
                    {
                        if (i != 0)
                        {
                            startIndex += Constants.MaxListRecordsToExportCountForPeriodic;
                        }
                        List<T> lstObjBaseChunk = lstObjBase.GetRange(startIndex, Constants.MaxListRecordsToExportCountForPeriodic);
                        numberOfRecordsExported += lstObjBaseChunk.Count;
                        FileUtility.WriteCsVintoFile(sourceFileName, ref lstObjBaseChunk, ref blHeader);
                        blHeader = false;
                    }
                    FileUtility.DoPeriodicFlush(sourceFileName, ref blHeader);
                    blHeader = true;
                    if ((lstObjBase.Count - numberOfRecordsExported) > 0)
                    {
                        List<T> lstObjBaseChunk = lstObjBase.GetRange(numberOfRecordsExported, (lstObjBase.Count - numberOfRecordsExported));
                        FileUtility.WriteCsVintoFile(sourceFileName, ref lstObjBaseChunk, ref blHeader);
                    }
                    FileUtility.DoPeriodicFlush(sourceFileName, ref blHeader);

                    lstObjBase = null;
                }

            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
