using System;

namespace PeoplePickerRemediation.Console.Common.CSV
{
   public static class CsvSplit
    {
        /// <summary>
        /// Split CSV files on line breaks before a certain size in bytes.
        /// </summary>
        public static void SplitCsv_bySize(string file, string prefix, int size)
        {
            // Read lines from source file
            string strDir = System.IO.Path.GetDirectoryName(file);
            string[] arr = System.IO.File.ReadAllLines(file);

            
            string headerrow = string.Empty;

            if (arr.LongLength>0)
            {
                headerrow = arr[0];
            }

            int total = 0;
            int num = 0;
            var writer = new System.IO.StreamWriter(System.IO.Path.Combine(strDir, GetFileName(prefix, num)),true, System.Text.Encoding.UTF8);
            writer.WriteLine(headerrow);

            //Having Content in Array since first row is header
            if (arr.LongLength>1)
            {
                for (int i = 1; i < arr.Length; i++)
                {

                    // Current line
                    string line = arr[i];
                    // Length of current line
                    int length = line.Length;

                    // See if adding this line would exceed the size threshold
                    if (total + length >= size)
                    {
                        // Create a new file
                        num++;
                        total = 0;
                        writer.Dispose();
                        writer = new System.IO.StreamWriter(System.IO.Path.Combine(strDir, GetFileName(prefix, num)),true, System.Text.Encoding.UTF8);
                        writer.WriteLine(headerrow);
                    }
                    // Write the line to the current file
                    writer.WriteLine(line);

                    // Add length of line in bytes to running size
                    total += length;

                    // Add size of newlines
                    total += Environment.NewLine.Length;
                }

            }
            // Loop through all source lines
            
            writer.Dispose();
        }


        public static void SplitCsv_byitemcount(string file, string prefix, int itemcount)
        {
            // Read lines from source file
            string strDir = System.IO.Path.GetDirectoryName(file);
            string[] arr = System.IO.File.ReadAllLines(file);

            
            string headerrow = string.Empty;

            if (arr.LongLength > 0)
            {
                headerrow = arr[0];
            }

            int total = 0;
            int num = 0;
            var writer = new System.IO.StreamWriter(System.IO.Path.Combine(strDir, GetFileName(prefix, num)),true, System.Text.Encoding.UTF8);
            writer.WriteLine(headerrow);

            //Having Content in Array since first row is header
            if (arr.LongLength > 1)
            {
                for (int i = 1; i < arr.Length; i++)
                {

                    // Current line
                    string line = arr[i];
                    // Length of current line
                    int currentItemCount = i;

                    // See if adding this line would exceed the size threshold
                    if (total >= itemcount)
                    {
                        // Create a new file
                        num++;
                        total = 0;
                        writer.Dispose();
                        writer = new System.IO.StreamWriter(System.IO.Path.Combine(strDir, GetFileName(prefix, num)), true, System.Text.Encoding.UTF8);
                        writer.WriteLine(headerrow);
                    }
                    // Write the line to the current file
                    writer.WriteLine(line);

                    // Add item count 
                    total += 1;
                    
                }

            }
            // Loop through all source lines

            writer.Dispose();
        }
        /// <summary>
        /// Get an output file name based on a number.
        /// </summary>
        static string GetFileName(string prefix, int num)
        {
            return prefix + "_" + num.ToString("00") + ".csv";
        }
    }
}
