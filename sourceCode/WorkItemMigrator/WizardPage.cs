//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    /// <summary>
    /// Enumerator that shows the different WizardPages possible in the Wizard. 
    /// It is used by Wizard to distinguish one Wizard page from another.
    /// </summary>
    internal enum WizardPage
    {
        Welcome,
        SelectDataSource,
        SelectDestinationServer,
        SettingsFile,
        FieldsSelection,
        FieldMapping,
        DataMapping,
        Linking,
        MiscSettings,
        ConfirmSettings,
        Summary
    }
}
