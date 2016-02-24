using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Transformation.PowerShell.Common.Utilities;

namespace Transformation.PowerShell.Common.CSV
{
    public class ImportCsv
    {
        private static readonly Regex RexCsvSplitter = new Regex(@",(?=(?:[^""]*""[^""]*"")*(?![^""]*""))");

        private static object ReadCSV(string filePath, object classType)
        {
            PropertyInfo[] properties = classType.GetType().GetProperties();
            ArrayUtility.ArraySort(properties);

            return classType;
        }

        private static void ReadCSV()
        {
        }

        private static IEnumerable<T> ReadData_as_IEnumerable<T>(string fileName, string delimeter)
            where T : class, new()
        {
            CsvReading<T> csvr = new CsvReading<T>();

            Type t = typeof (T);
            //FieldInfo[] fields = t.GetFields();
            PropertyInfo[] properties = t.GetProperties();

            properties = ArrayUtility.ArraySort(properties);


            bool firstrow = true;
            string[] headerrow = null;
            Hashtable ht = new Hashtable();

            using (StreamReader r = new StreamReader(fileName))
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    if (firstrow)
                    {
                        string[] delimiterChars = {delimeter};
                        headerrow = line.Split(delimiterChars, StringSplitOptions.RemoveEmptyEntries);
                        if (!CsVisComaptiblewithClass(properties, headerrow))
                        {
                            break;
                        }
                    }
                    else
                    {
                        T obj = default(T);
                        string[] delimiterChars = {delimeter};
                        string[] values = RexCsvSplitter.Split(line);
                        //string[] Contentline = line.Split(delimiterChars, StringSplitOptions.None);
                        obj = csvr.ReadObject(values, properties, headerrow);

                        yield return obj;
                    }
                    firstrow = false;
                }
            }
        }

        private static Collection<T> ReadData<T>(string fileName, string delimeter) where T : class, new()
        {
            Collection<T> classCollection = new Collection<T>();

            CsvReading<T> csvr = new CsvReading<T>();

            Type t = typeof (T);
            //FieldInfo[] fields = t.GetFields();
            PropertyInfo[] properties = t.GetProperties();

            properties = ArrayUtility.ArraySort(properties);


            bool firstrow = true;
            string[] headerrow = null;
            Hashtable ht = new Hashtable();

            using (StreamReader r = new StreamReader(fileName))
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    if (firstrow)
                    {
                        string[] delimiterChars = {delimeter};
                        headerrow = line.Split(delimiterChars, StringSplitOptions.RemoveEmptyEntries);
                        if (!CsVisComaptiblewithClass(properties, headerrow))
                        {
                            break;
                        }
                    }
                    else
                    {
                        T obj = default(T);
                        string[] delimiterChars = {delimeter};
                        string[] values = RexCsvSplitter.Split(line);
                        //string[] Contentline = line.Split(delimiterChars, StringSplitOptions.None);
                        obj = csvr.ReadObject(values, properties, headerrow);

                        classCollection.Add(obj);
                    }
                    firstrow = false;
                }
            }

            return classCollection;
        }



        public static IEnumerable<T> Read<T>(string fileName, string Delimeter) where T : class, new()
        {
            IEnumerable<T> ie = null;
            try
            {
                ie = ReadData<T>(fileName, Delimeter).Distinct();
                return ie;
            }
            catch
            {
                return ie;
            }
        }

        public static IEnumerable<T> ReadMatchingColumns<T>(string fileName, string Delimeter) where T : class, new()
        {
            IEnumerable<T> ie = null;
            try
            {
                ie = ReadMatchingColumnsData<T>(fileName, Delimeter).Distinct();
                return ie;
            }
            catch
            {
                throw;
            }
        }

        private static Collection<T> ReadMatchingColumnsData<T>(string fileName, string delimeter) where T : class, new()
        {
            Collection<T> classCollection = new Collection<T>();

            CsvReading<T> csvr = new CsvReading<T>();

            Type t = typeof(T);
            PropertyInfo[] properties = t.GetProperties();

            properties = ArrayUtility.ArraySort(properties);


            bool firstrow = true;
            string[] headerrow = null;
            Hashtable ht = new Hashtable();

            using (StreamReader r = new StreamReader(fileName))
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    if (firstrow)
                    {
                        string[] delimiterChars = { delimeter };
                        headerrow = line.Split(delimiterChars, StringSplitOptions.RemoveEmptyEntries);
                        if (!CsVContainsMatchingPropertiesOfClass(properties, headerrow))
                        {
                            break;
                        }
                    }
                    else
                    {
                        T obj = default(T);
                        string[] delimiterChars = { delimeter };
                        string[] values = RexCsvSplitter.Split(line);
                        obj = csvr.ReadObject(values, properties, headerrow);

                        classCollection.Add(obj);
                    }
                    firstrow = false;
                }
            }

            return classCollection;
        }

        private static bool CsVContainsMatchingPropertiesOfClass(PropertyInfo[] properties, string[] firstrowofCsv)
        {
            try
            {
                if (ArrayUtility.CompareTwoArrayForCommonElements(firstrowofCsv, properties))
                {
                    return true;
                }
                return false;
            }
            catch
            {
                throw;
            }
        }

        private static bool CsVisComaptiblewithClass(PropertyInfo[] properties, string[] firstrowofCsv)
        {
            try
            {
                if (firstrowofCsv.Count() == properties.Count() &&
                    ArrayUtility.CompareTwoArray(firstrowofCsv, properties))
                {
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        public static void OutputData<T>(IEnumerable<T> dataRows, string title)
        {
            foreach (T row in dataRows)
            {
            }
        }

        public static DataTable ConvertCsVtoDataTable(string strFilePath)
        {
            DataTable dt = new DataTable();
            using (StreamReader sr = new StreamReader(strFilePath))
            {
                string readLine = sr.ReadLine();
                if (readLine != null)
                {
                    string[] headers = readLine.Split(',');

                    foreach (string header in headers)
                    {
                        dt.Columns.Add(header);
                    }
                    while (!sr.EndOfStream)
                    {
                        string[] rows = readLine.Split(',');
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = rows[i];
                        }
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }
    }


    public class FieldMappers<T>
    {
    }

    public class CsvReading<T> : FieldMappers<T> where T : new()
    {
        public T ReadObject(string line, PropertyInfo[] properties, string delimeter)
        {
            T obj = new T();

            if (properties == null)
            {
                Type t = typeof (T);
                properties = t.GetProperties();
            }

            string[] delimiterChars = {delimeter};
            string[] csvline = line.Split(delimiterChars, StringSplitOptions.RemoveEmptyEntries);

            int counter = 0;
            foreach (string csvvalue in csvline)
            {
                properties[counter].SetValue(obj, csvvalue);
                counter = counter + 1;
            }

            return obj;
        }

        public T ReadObject(string[] contentline, PropertyInfo[] properties, string[] headerLine)
        {
            T obj = new T();

            if (properties == null)
            {
                Type t = typeof (T);
                properties = t.GetProperties();
                properties = ArrayUtility.ArraySort(properties);
            }

            Dictionary<string, string> contentRows = ArrayUtility.MergeTwoArray(headerLine, contentline);
            foreach (KeyValuePair<string, string> pair in contentRows)
            {
                foreach (PropertyInfo property in properties)
                {
                    if (property.Name.ToLower() == pair.Key.ToLower())
                    {
                        property.SetValue(obj, pair.Value.Trim());
                        break;
                    }
                }
            }
            return obj;
        }
    }
}