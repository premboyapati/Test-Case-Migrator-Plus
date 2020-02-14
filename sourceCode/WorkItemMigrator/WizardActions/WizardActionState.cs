//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    /// <summary>
    /// Enum that represents the different state a Wizard Action can take during course of its life time
    /// </summary>
    internal enum WizardActionState
    {
        Pending,
        InProgress,
        Success,
        Warning,
        Failed,
        Stopped
    }
}
