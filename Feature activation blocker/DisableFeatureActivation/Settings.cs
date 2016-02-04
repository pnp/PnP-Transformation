// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Settings.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// <summary>
//   Adding custom page header to restrict feature activation and deactivation for set of features as part of DvNext migration.
// </summary>
// ---------------------------------------------------------------------------------------------------------------------


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDP.Transformation.DisableFeatureActivation
{
    public static class Settings
    {
        #region Generic variables
        public static string customDisableMsgTitle = "FTC retraction in progress";
        public static string customDisableMgsHtml = "Activation of some features is disabled.";
        public static readonly string featureIdsListName = "Disable Features List";
        public static readonly string custMessageListName = "Disable Feature Message";
        #endregion

        #region Cache Settings and Keys
        public static readonly string featuresToBlock_cacheKey = "BlockedFeatureIdsDetails";
        public static readonly string customDisableMsgTitle_cacheKey = "customDisableMsgTitle";
        public static readonly string customDisableMgsHtml_cacheKey = "customDisableMgsHtml";
        public static readonly int BatchCachCacheDuration = 10;
        #endregion
    }
}
