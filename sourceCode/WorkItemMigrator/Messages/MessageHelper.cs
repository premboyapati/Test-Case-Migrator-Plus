//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    /// <summary>
    /// Delegate Definition used by Message helper to show Messages
    /// </summary>
    /// <param name="args"></param>
    internal delegate void ShowMessageDelegate(MessageEventArgs args);


    internal static class MessageHelper
    {
        #region Properties

        /// <summary>
        /// Delegate for showing messges
        /// </summary>
        public static ShowMessageDelegate ShowMessageWindow
        {
            get;
            set;
        }

        #endregion

        #region Information Messages

        /// <summary>
        /// Shows information with title and information details with OK button
        /// </summary>
        /// <param name="title">Title of the information</param>
        /// <param name="likelyCause">Detail information</param>
        public static void ShowInfo(string title, string likelyCause)
        {
            if (ShowMessageWindow != null)
            {
                MessageEventArgs args = new MessageEventArgs(MessageCategory.Information,
                                                             title,
                                                             likelyCause,
                                                             null,
                                                             null,
                                                             null,
                                                             Resources.OKButtonLabel,
                                                             null);
                ShowMessageWindow(args);
            }
        }

        #endregion

    }
}
