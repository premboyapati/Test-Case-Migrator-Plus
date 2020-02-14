//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Windows.Input;

    /// <summary>
    /// This is a class to help display a busy cursor during long operations
    /// that will automatically revert to a "normal" cursor when it's through.
    internal class AutoWaitCursor : IDisposable
    {
        #region Constructor

        /// <summary>
        /// Loads Wait Cursor in UI Context
        /// </summary>
        public AutoWaitCursor()
        {
            App.CallMethodInUISynchronizationContext(ShowWaitCursor, null);
        }

        #endregion

        #region Public Methods

        // Flag used to don't show waiting cursor during console mode
        public static bool IsConsoleMode = false;

        void IDisposable.Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Overrides Mouse Cursor to show Wait Cursor
        /// </summary>
        /// <param name="value"></param>
        private static void ShowWaitCursor(object value)
        {
            try
            {
                // Override Mouse Cursor if the tool is not in console mode
                if (!IsConsoleMode)
                {
                    Mouse.OverrideCursor = Cursors.Wait;
                }
            }
            catch (InvalidOperationException)
            { }
        }

        /// <summary>
        /// Disposes the Wait Cursore and restore normal cursor in UI Synchronization Context
        /// </summary>
        /// <param name="disposing"></param>
        void Dispose(bool disposing)
        {
            App.CallMethodInUISynchronizationContext(RestoreNormalCursor, disposing);
        }

        /// <summary>
        /// Restores normal mouse cursor. This method should be call in UI 
        /// </summary>
        /// <param name="disposing"></param>
        private static void RestoreNormalCursor(object disposing)
        {
            try
            {
                if ((bool)disposing && !IsConsoleMode)
                {
                    Mouse.OverrideCursor = null;
                }
            }
            catch (InvalidOperationException)
            { }
        }

        #endregion
    }
}
