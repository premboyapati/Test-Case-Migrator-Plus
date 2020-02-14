//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    /// <summary>
    /// Used by Message Event Args to decide whether a message is a warning, information or error message.
    /// </summary>
    internal enum MessageCategory
    {
        Information,
        Warning,
        Error
    }
}
