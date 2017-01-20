using PeoplePickerRemediation.Console.Common.Base;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;

namespace PeoplePickerRemediation.Console
{
    public class ContentIterator
    {
        public readonly ClientContext _context;
        public Dictionary<string, string> dictUserUpns = new Dictionary<string, string>();

        public ContentIterator(ClientContext context)
        {
            if (context == null) throw new ArgumentNullException("context");
            _context = context;
        }

        public delegate void ItemsProcessor(ListItemCollection items);

        public delegate bool ItemsProcessorErrorCallout(ListItemCollection items, System.Exception e);

        public delegate void ItemProcessor(ListItem item, ClientContext context, ref List<PeoplePickerListOutput> lstPeoplepickeroutput);

        public delegate bool ItemProcessorErrorCallout(ListItem item, System.Exception e);

        private const string itemEnumerationOrderByID = "<OrderBy Override='TRUE'><FieldRef Name='ID' /></OrderBy>";

        private const string itemEnumerationOrderByIDDesc = "<OrderBy Override='TRUE' ><FieldRef Name='ID' Ascending='FALSE' /></OrderBy>";

        private const string itemEnumerationOrderByPath = "<OrderBy Override='TRUE'><FieldRef Name='FileDirRef' /><FieldRef Name='FileLeafRef' /></OrderBy>";

        private const string itemEnumerationOrderByNVPField = "<OrderBy UseIndexForOrderBy='TRUE' Override='TRUE' />";

        private const string overrideQueryThrottleMode = "<QueryThrottleMode>Override</QueryThrottleMode>";

        /// <summary>
        /// Process ListItem one by one
        /// </summary>
        /// <param name="listName">ListName</param>
        /// <param name="camlQuery">CamlQuery</param>
        /// <param name="itemProcessor">itemprocessor delegate</param>
        /// <param name="errorCallout">error delegate</param>
        public void ProcessListItem(string listName, CamlQuery camlQuery, ItemProcessor itemProcessor,ref List<PeoplePickerListOutput> lstPeoplepickeroutput, ItemProcessorErrorCallout errorCallout)
        {
            List list = _context.Web.Lists.GetByTitle(listName);
            CamlQuery query = camlQuery;            

            //EventReceiverDefinitionCollection erCollection = list.EventReceivers;
            //foreach(EventReceiverDefinition erDefinition in erCollection)
            //{
            //    erDefinition.
            //}
            ListItemCollectionPosition position = null;
            query.ListItemCollectionPosition = position;            

            while (true)
            {
                ListItemCollection listItems = list.GetItems(query);
                _context.Load(listItems, items => items.ListItemCollectionPosition);
                _context.ExecuteQueryRetry();

                for (int i = 0; i < listItems.Count; i++)
                {
                    try
                    {
                        itemProcessor(listItems[i], _context, ref lstPeoplepickeroutput);

                    }
                    catch (System.Exception ex)
                    {
                        if (errorCallout == null || errorCallout(listItems[i], ex))
                        {
                            throw;
                        }
                    }

                }

                if (listItems.ListItemCollectionPosition == null)
                {
                    return;
                }
                else
                {
                    /*if query contains lookup column filter last batch returns null 
                     by removing the lookup column in paginginfo query will return next records
                     */
                    string pagingInfo = listItems.ListItemCollectionPosition.PagingInfo;
                    string[] parameters = pagingInfo.Split(new char[] { '&' }, StringSplitOptions.RemoveEmptyEntries);
                    List<string> requiredParameters = new List<string>();
                    foreach (string str in parameters)
                    {
                        if (str.Contains("Paged=") || str.Contains("p_ID="))
                            requiredParameters.Add(str);
                    }

                    pagingInfo = string.Join("&", requiredParameters.ToArray());
                    listItems.ListItemCollectionPosition.PagingInfo = pagingInfo;
                    query.ListItemCollectionPosition = listItems.ListItemCollectionPosition;
                }

            }

        }
        
    }
}
