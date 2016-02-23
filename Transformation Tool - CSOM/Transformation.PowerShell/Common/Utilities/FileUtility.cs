using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Transformation.PowerShell.Common.CSV;

namespace Transformation.PowerShell.Common.Utilities
{
    public static class FileUtility
    {
        public static void AppendTextintoFile(string filePath, string content)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    // Create a file to write to. 
                    using (StreamWriter sw = File.CreateText(filePath))
                    {
                        sw.WriteLine(content);
                    }
                }
                else
                {
                    // This text is always added, appending the text into file 
                    if (File.Exists(filePath))
                    {
                        using (StreamWriter sw = File.AppendText(filePath))
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

        public static void WriteCsVintoFile<T>(string filepath, ref List<T> list, ref bool blHeader)
        {
            if (list.Count > 0)
            {
                //MasterPageBase[] objaray=  listmasterpagebase.ToArray();
                string strCsv = ExportCsv.ToCsv(",", list, ref blHeader);

                AppendTextintoFile(filepath, strCsv);
            }

            list.TrimExcess();
            list = null;
        }

        public static void WriteCsVintoFile<T>(string filepath, T obj, ref bool blHeader)
        {            
            if (obj != null)
            {
                string strCsv = ExportCsv.ToCsv(",", obj, ref blHeader);

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
                    }
                }
                catch
                {
                }
            }
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
                    Console.WriteLine(directory);
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
            return fi.Length/1024;
        }
    }
}