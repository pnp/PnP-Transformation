using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using Transformation.PowerShell.Common.Utilities;

namespace Transformation.PowerShell.Common.CSV
{
    public class ExportCsv
    {
        public static string ToCsv<T>(string separator, List<T> objectlist, ref bool blheader)
        {
            Type t = typeof (T);
            //FieldInfo[] fields = t.GetFields();
            PropertyInfo[] properties = t.GetProperties();
            properties = ArrayUtility.ArraySort(properties);
            StringBuilder csvdata = new StringBuilder();
            if (!blheader)
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
            //FieldInfo[] fields = t.GetFields();
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

                /*if (x != null)
                    linie.Append(x.ToString());*/
                if (x != null)
                {
                    if (x.ToString().Contains(","))
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
                /*
                if (x != null)
                    linie.Append(x.ToString());*/
                if (x != null)
                {
                    if (x.ToString().Contains(","))
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
    }
}