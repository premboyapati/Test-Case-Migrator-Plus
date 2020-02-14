//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for WelcomeView.xaml
    /// </summary>
    internal partial class WelcomeView : ContentControl
    {
        public WelcomeView(WelcomePart part)
        {
            InitializeComponent();
            part.CanNext = true;
        }
    }
}
