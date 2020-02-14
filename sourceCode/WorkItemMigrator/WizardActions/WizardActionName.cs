//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    /// <summary>
    /// Enum for uniquly identifying each wizard action and distinguish them from each other
    /// </summary>
    internal enum WizardActionName
    {
        ParseDataSource,
        SaveSettings,
        MigrateWorkItems,
        Relationships,
        PublishReport
    }
}
