// --------------------------------------------------------------------------------------------------------------------
// <copyright file="JDPDisableFeatureActivation.ascx.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// <summary>
//   Adding custom page header to restrict feature activation and deactivation for set of features as part of DvNext migration.
// </summary>
// ---------------------------------------------------------------------------------------------------------------------

using System;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Collections.Generic;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint;

namespace JDP.Transformation.DisableFeatureActivation.CONTROLTEMPLATES
{

    public partial class JDPDisableFeatureActivation : UserControl
    {
        #region variables used to hold data
        private List<string> featureIds = new List<string>();
        #endregion

        #region Protected methods
        protected bool IsRendered { get; set; }
        protected void Page_Load(object sender, EventArgs e)
        {
            IsRendered = false;

            string path = Context.Request.Url.LocalPath;
            if (path.ToLower().Contains("_layouts/15/managefeatures.aspx"))
            {
                // current page is feature management page
                IsRendered = true;
            }
        }

        protected override void OnPreRender(EventArgs e)
        {
            if (IsRendered)
            {
                GetFeatureIDs();
                ProcessControls(Page.Controls);
            }

            base.OnPreRender(e);
        }

        /// <summary>
        /// This method displays custom feature disabled message and also disbles set of features activation and deactivation
        /// </summary>
        /// <param name="writer"></param>
        protected override void Render(HtmlTextWriter writer)
        {
            if (IsRendered)
            {
                                
                SPPageStatusSetter ctlStatusBar = new SPPageStatusSetter();
                ctlStatusBar.AddStatus(
                    Settings.customDisableMsgTitle,
                    Settings.customDisableMgsHtml,
                    SPPageStatusColor.Yellow);
                this.Controls.Add(ctlStatusBar);
            }
            base.Render(writer);
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// This method reads all features displaying on the page and blocks the features activation and deactivation for specified set of features
        /// </summary>
        /// <param name="controls"></param>
        private void ProcessControls(ControlCollection controls)
        {
            if (controls == null)
                return;

            foreach (Control c in controls)
            {
                Microsoft.SharePoint.WebControls.FeatureActivatorItem featureActivator = c as Microsoft.SharePoint.WebControls.FeatureActivatorItem;
                if (featureActivator != null)
                {
                    Type t = featureActivator.GetType();
                    string featureId = t.GetProperty("FeatureId", System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance).GetValue(featureActivator, null)+"";
                    if (featureIds.Contains(featureId.ToLower()))
                    {
                        Button btnActivate = t.GetField("btnActivate", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(featureActivator) as Button;
                        btnActivate.Enabled = false;
                    }
                }

                ProcessControls(c.Controls);
            }
        }

        /// <summary>
        /// This method reads all the features which needs to be disabled from feature activation and deactivation
        /// </summary>
        /// <returns></returns>
        private void ReadFeatureConfigDetails()
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (SPWeb web = new SPSite(SPContext.Current.Site.WebApplication.Sites[0].ID).OpenWeb())
                {
                    if (web != null)
                    {
                        #region Reading custom messages to display on feature management page
                        SPList list = web.Lists.TryGetList(Settings.custMessageListName);
                        if (list != null && list.ItemCount > 0)
                        {
                            SPQuery query = new SPQuery();
                            query.ViewFields = "<FieldRef Name='Title'/><FieldRef Name='Description'/>";
                            query.RowLimit = 1;
                            SPListItemCollection items = list.GetItems(query);
                            if (items != null && items.Count == 1)
                            {
                                SPListItem item = items[0];
                                //Reading custom disable features message title
                                if (item["Title"] != null)
                                {
                                    Settings.customDisableMsgTitle = item["Title"].ToString();
                                }
                                //Reading custom disable features message
                                if (item["Description"] != null)
                                {
                                    Settings.customDisableMgsHtml = item["Description"].ToString();
                                }
                            }
                        }
                        #endregion

                        #region Reading featureids to disable activaiton/deactivation on feature management page

                        list = web.Lists.TryGetList(Settings.featureIdsListName);
                        if (list != null)
                        {
                            featureIds = new List<string>();

                            string[] fields = { "Title", "FeatureID" };
                            SPListItemCollection items = list.GetItems(fields);
                            foreach (SPListItem item in items)
                            {
                                if (item["FeatureID"] != null)
                                {
                                    featureIds.Add(item["FeatureID"].ToString().ToLower().Trim());
                                }
                            }

                        }
                        #endregion
                    }
                }
            });
        }

        private string getWebAppName()
        {
            string strName = "";
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                strName = SPContext.Current.Site.WebApplication.Name;
            });
            return strName.Replace(" ","");
        }

        /// <summary>
        /// This method reads cache for list of disabled features, if the cache is not present then builds the cache for that
        /// </summary>
        /// <returns></returns>
        private void GetFeatureIDs()
        {
            //adding cache Impenentation
            string strName = getWebAppName();

            if ((featureIds = (List<string>)System.Web.HttpContext.Current.Cache.Get(strName + "_" + Settings.featuresToBlock_cacheKey)) == null ||
                (Settings.customDisableMsgTitle = (string)System.Web.HttpContext.Current.Cache.Get(strName + "_" + Settings.customDisableMsgTitle_cacheKey)) == null ||
                (Settings.customDisableMgsHtml = (string)System.Web.HttpContext.Current.Cache.Get(strName + "_" + Settings.customDisableMgsHtml_cacheKey)) == null)
            {
                ReadFeatureConfigDetails();
                if (featureIds != null && featureIds.Count > 0)
                {
                    System.Web.HttpContext.Current.Cache.Insert(strName + "_" + Settings.featuresToBlock_cacheKey, featureIds, null, DateTime.Now.AddMinutes(Settings.BatchCachCacheDuration), System.Web.Caching.Cache.NoSlidingExpiration);
                    System.Web.HttpContext.Current.Cache.Insert(strName + "_" + Settings.customDisableMsgTitle_cacheKey, Settings.customDisableMsgTitle, null, DateTime.Now.AddMinutes(Settings.BatchCachCacheDuration), System.Web.Caching.Cache.NoSlidingExpiration);
                    System.Web.HttpContext.Current.Cache.Insert(strName + "_" + Settings.customDisableMgsHtml_cacheKey, Settings.customDisableMgsHtml, null, DateTime.Now.AddMinutes(Settings.BatchCachCacheDuration), System.Web.Caching.Cache.NoSlidingExpiration);
                }
            }
        }
        #endregion
    }
}
