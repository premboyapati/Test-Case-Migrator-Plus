//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    /// <summary>
    /// Enum for separating one message response from other.
    /// </summary>
    internal enum MessageResponseDefinition
    {
        FirstResponse = 1,
        SecondResponse = 2,
        CancelResponse = -1
    }
}
