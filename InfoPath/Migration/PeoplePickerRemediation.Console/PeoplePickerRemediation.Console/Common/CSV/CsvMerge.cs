
using PeoplePickerRemediation.Console.Common.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace PeoplePickerRemediation.Console.Common.CSV
{
    public static class CsvMerge
    {
        public static void CSV_Merge(string sourceDirectory, string mergeFileName,bool deleteSplitFiles,string searchPattern="*.*")
        {
            try
            {
                if (sourceDirectory != "")
                {
                    int counter = 0;

                    string csvfile = sourceDirectory + "\\" + mergeFileName;

                    if (File.Exists(csvfile))
                    {
                        File.Delete(csvfile);
                    }

                    string[] files = FileUtility.FindAllFilewithSearchPattern(sourceDirectory, searchPattern);
                    if (files == null) throw new ArgumentNullException(sourceDirectory + " not exist");
                    foreach (var file in files)
                    {
                        //StringBuilder sb = new StringBuilder();
                        string filename = Path.GetFileNameWithoutExtension(file);
                        if (file.EndsWith(".csv"))
                        {
                            using (StreamWriter writer = new StreamWriter(csvfile, true, System.Text.Encoding.UTF8))
                            {
                                string[] rows = File.ReadAllLines(file);
                                for (int i = 0; i < rows.Length; i++)
                                {
                                    if (i == 0)
                                    {
                                        if (counter == 0)
                                        {
                                            //sb.Append(rows[i] + "\n");
                                            writer.WriteLine(rows[i]);
                                            counter++;
                                        }
                                    }
                                    else
                                    {
                                        //sb.Append(rows[i] + "\n");
                                        writer.WriteLine(rows[i]);
                                    }
                                }
                                writer.Flush();
                                writer.Close();
                            }
                        }
                    }
                    counter = 0;

                    //Delete All the Split Files
                    if (deleteSplitFiles)
                    {
                        foreach (var file in files)
                        {
                            File.Delete(file);
                        }
                    }
                }
            }
            catch (Exception)
            {
                
                throw;
            }
        }
    }
}
