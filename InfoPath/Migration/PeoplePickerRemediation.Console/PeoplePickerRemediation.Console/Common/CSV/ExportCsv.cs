using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Text;
using PeoplePickerRemediation.Console.Common.Utilities;

namespace PeoplePickerRemediation.Console.Common.CSV
{
    class ExportCsv
    {
        public static string ToCsv<T>(string separator, List<T> objectlist)
        {
            Type t = typeof(T);
            PropertyInfo[] properties = t.GetProperties();
            StringBuilder csvdata = new StringBuilder();
            
            string header = String.Join(separator, properties.Select(f => f.Name).ToArray());
            csvdata.AppendLine(header);
               
            foreach (T o in objectlist)
                csvdata.AppendLine(ToCsvFields(separator, properties, o));

            return csvdata.ToString().Trim();
        }

        public static string ToCsv<T>(string separator, List<T> objectlist, ref bool blheader)
        {
            Type t = typeof(T);
            PropertyInfo[] properties = t.GetProperties();
            //properties = ArrayUtility.ArraySort(properties);
            StringBuilder csvdata = new StringBuilder();
            if (blheader)
            {
                string header = String.Join(separator, properties.Select(f => f.Name).ToArray());
                csvdata.AppendLine(header);
                blheader = true;
            }
            foreach (T o in objectlist)
                csvdata.AppendLine(ToCsvFields(separator, properties, o));

            return csvdata.ToString().Trim();
        }

        public static string ToCsv<T>(string separator, T obj, ref bool blheader)
        {
            Type t = typeof(T);
            PropertyInfo[] properties = t.GetProperties();
            properties = ArrayUtility.ArraySort(properties);
            StringBuilder csvdata = new StringBuilder();
            if (!blheader)
            {
                string header = String.Join(separator, properties.Select(f => f.Name).ToArray());
                csvdata.AppendLine(header);
                blheader = true;
            }
            csvdata.AppendLine(ToCsvFields(separator, properties, obj));
            return csvdata.ToString().Trim();
        }

        private static string ToCsvFields(string separator, FieldInfo[] fields, object o)
        {
            StringBuilder linie = new StringBuilder();

            foreach (FieldInfo f in fields)
            {
                if (linie.Length > 0)
                    linie.Append(separator);

                object x = f.GetValue(o);

                if (x != null)
                {
                    if (x.ToString().Contains(",") || x.ToString().Contains("\n"))
                    {
                        x = String.Format("\"{0}\"", x);
                        linie.Append(x);
                    }
                    else
                    {
                        linie.Append(x);
                    }
                }
            }

            return linie.ToString();
        }

        private static string ToCsvFields(string separator, PropertyInfo[] properties, object o)
        {
            StringBuilder linie = new StringBuilder();

            foreach (PropertyInfo f in properties)
            {
                if (linie.Length > 0)
                    linie.Append(separator);

                object x = f.GetValue(o);

                if (x != null)
                {
                    if (x.ToString().Contains(",") || x.ToString().Contains("\n") || x.ToString().Contains("\""))
                    {
                        if (x.ToString().Contains("\""))
                        {
                            x = x.ToString().Replace("\"", "'");
                        }
                        x = String.Format("\"{0}\"", x);
                        linie.Append(x);
                    }
                    else
                    {
                        linie.Append(x);
                    }
                }
            }

            return linie.ToString();
        }
    }
}
