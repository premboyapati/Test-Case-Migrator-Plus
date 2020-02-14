//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Windows;
    using System.Windows.Controls;
    using Microsoft.Win32;

    /// <summary>
    /// Interaction logic for SettingsFile.xaml
    /// </summary>
    internal partial class SettingsFileView : ContentControl
    {
        #region Fields

        SettingsFilePart m_part;

        #endregion

        #region Constructor

        public SettingsFileView(SettingsFilePart part)
        {
            InitializeComponent();
            m_part = part;
            DataContext = m_part;
        }

        #endregion

        #region Private Methods

        private void LoadSettingButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "XMl files (*.xml)|*.xml|All files(*.*)|*.*";

            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                m_part.SettingsFilePath = dlg.FileName;
            }
        }

        #endregion

        private void WITMigratorTextBox_TextChangeAction(object sender, MessageEventArgs e)
        {
            TextBox textbox = sender as TextBox;
            m_part.SettingsFilePath = textbox.Text;
            m_part.CanNext = m_part.ValidatePartState();

            e.SetValues(m_part.ValidateSettingsFilePath());
        }
    }
}
