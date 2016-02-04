// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Control.EventReceiver.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// <summary>
//   Adding custom page header to restrict feature activation and deactivation for set of features as part of DvNext migration.
// </summary>
// ---------------------------------------------------------------------------------------------------------------------


using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace JDP.Transformation.DisableFeatureActivation.Features.Control
{
    /// <summary>
    /// This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
    /// </summary>
    /// <remarks>
    /// The GUID attached to this class may be used during packaging and should not be modified.
    /// </remarks>

    [Guid("5037b1e8-28fd-4fbd-9216-eaa369943177")]
    public class ControlEventReceiver : SPFeatureReceiver
    {
        #region Feature activation/deactivation logic
        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            try
            {
                SPWebApplication wa = (SPWebApplication)properties.Feature.Parent;
                createLists(wa.Sites[0].Url);
            }
            catch
            {
                throw;
            }
        }

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            try
            {
                SPWebApplication wa = (SPWebApplication)properties.Feature.Parent;
                deletLists(wa.Sites[0].Url);
            }
            catch
            {
                throw;
            }
        }

        #endregion

        #region Private methods to create and delete lists
        /// <summary>
        /// This method used to create 'Disable Features List' list, which will be used to hold all the features to be disabled
        /// </summary>
        /// <param name="web"></param>
        private void featuresList(SPWeb web)
        {
            SPList list = web.Lists.TryGetList(Settings.featureIdsListName);
            if (list == null)
            {
                SPListTemplate template = web.ListTemplates["Custom List"];
                Guid listId = web.Lists.Add(Settings.featureIdsListName, Settings.featureIdsListName, template);
                list = web.GetList(web.Url.Trim() + "/Lists/" + Settings.featureIdsListName);
                list.Fields.Add("FeatureID", SPFieldType.Text, true);
                SPView view = list.Views["All Items"];
                view.ViewFields.Add("FeatureID");
                view.Update();
                list.Update();
            }
        }

        /// <summary>
        /// This method used to create 'Disable Feature Message' list, which will hold generic message i.e. to displayed on features page
        /// </summary>
        /// <param name="web"></param>
        private void disableCustomMsgList(SPWeb web)
        {
            SPList list = web.Lists.TryGetList(Settings.custMessageListName);
            if (list == null)
            {
                SPListTemplate template = web.ListTemplates["Custom List"];
                Guid listId = web.Lists.Add(Settings.custMessageListName, Settings.custMessageListName, template);
                list = web.GetList(web.Url.Trim() + "/Lists/" + Settings.custMessageListName);
                list.Fields.Add("Description", SPFieldType.Text, true);
                SPView view = list.Views["All Items"];
                view.ViewFields.Add("Description");
                view.Update();
                list.Update();
            }
        }

        /// <summary>
        /// This method used to create lists to support disable functionality
        /// </summary>
        /// <param name="url"></param>
        private void createLists(string url)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using(SPWeb web = new SPSite(url).OpenWeb())
                {
                    featuresList(web);
                    disableCustomMsgList(web);
                }
            });
        }

        /// <summary>
        /// This method used to delete lists that are created as part of feature activation
        /// </summary>
        /// <param name="url"></param>
        private void deletLists(string url)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (SPWeb web = new SPSite(url).OpenWeb())
                {
                    SPList list = web.Lists.TryGetList(Settings.featureIdsListName);
                    if (list != null)
                    {
                        list.Delete();
                    }

                    list = web.Lists.TryGetList(Settings.custMessageListName);
                    if (list != null)
                    {
                        list.Delete();
                    }
                }
            });
        }
        #endregion
    }
}
