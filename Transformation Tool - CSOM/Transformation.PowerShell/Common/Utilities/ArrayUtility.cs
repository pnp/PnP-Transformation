using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Transformation.PowerShell.Common.Utilities
{
    public static class ArrayUtility
    {
        public static PropertyInfo[] ArraySort(PropertyInfo[] properties)
        {
            properties = properties.Select(x => new {Property = x})
                .OrderBy(x => x.Property.Name != null ? x.Property.Name : Constants.NotApplicable)
                .Select(x => x.Property)
                .ToArray();
            return properties;
        }

        public static string[] ArraySort(string[] stringarr)
        {
            Array.Sort(stringarr);
            return stringarr;
        }

        //public static dynamic MergeTwoArray(string[] Headers,string[] rows)
        //{             
        //    var returnrows= rows.Zip(Headers, (number, word) => new KeyValuePair<string, string>(word, number));
        //    return returnrows;
        //}

        public static Dictionary<string, string> MergeTwoArray(string[] headers, string[] rows)
        {
            Dictionary<string, string> dictionary = new Dictionary<string, string>();

            for (int index = 0; index < headers.Length; index++)
            {
                dictionary.Add(headers[index], StringUtility.Unescape(rows[index]));
            }

            return dictionary;
        }

        public static bool CompareTwoArray(string[] strarray, PropertyInfo[] propertyarray)
        {
            try
            {
                for (int i = 0; i < strarray.Length; i++)
                {
                    string tmp = strarray[i];
                    bool tmp1 = false;
                    for (int index = 0; index < propertyarray.Length; index++)
                    {
                        PropertyInfo tmp2 = propertyarray[index];
                        if (tmp.ToLower().Trim() == tmp2.Name.ToLower().Trim())
                        {
                            tmp1 = true;
                            break;
                        }
                    }

                    if (!tmp1)
                    {
                        return false;
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool CompareTwoArrayForCommonElements(string[] strarray, PropertyInfo[] propertyarray)
        {
            try
            {
                //for (int i = 0; i < strarray.Length; i++)
                for (int i = 0; i < propertyarray.Length; i++)
                {
                    //string tmp = strarray[i];
                    PropertyInfo tmp2 = propertyarray[i];
                    bool tmp1 = false;
                    //for (int index = 0; index < propertyarray.Length; index++)
                    for (int index = 0; index < strarray.Length; index++)
                    {
                        //PropertyInfo tmp2 = propertyarray[index];
                        string tmp = strarray[index];
                        if (tmp.ToLower().Trim() == tmp2.Name.ToLower().Trim())
                        {
                            tmp1 = true;
                            break;
                        }
                    }

                    if (!tmp1)
                    {
                        throw new Exception("The property of class file : " + tmp2 + " does not match with any of the header column of the file");
                    }
                }
                return true;
            }
            catch
            {
                throw;
            }
        }
    }
}