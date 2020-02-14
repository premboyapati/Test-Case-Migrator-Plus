//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for ConfirmSettings.xaml
    /// </summary>
    internal partial class ConfirmSettingsView : ContentControl
    {
        #region Fields

        ConfirmSettingsPart m_part;

        #endregion

        #region Constructor

        public ConfirmSettingsView(ConfirmSettingsPart part)
        {
            m_part = part;
            InitializeComponent();
            DataContext = part.WizardInfo;
            part.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(Part_PropertyChanged);
        }

        #endregion

        #region Private methods

        private void Part_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "WizardInfo")
            {
                App.CallMethodInUISynchronizationContext(SetDataContext, null);
            }
        }

        private void SetDataContext(object value)
        {
            DataContext = m_part.WizardInfo;
        }

        private void Hyperlink_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Clipboard.SetText(m_part.WizardInfo.CommandUsed);
        }

        #endregion
    }
}
