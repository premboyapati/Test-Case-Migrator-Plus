//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.IO;
    using System.Windows;
    using System.Windows.Controls;
    using WizardResource = Microsoft.VisualStudio.TestTools.WorkItemMigrator.Resources;

    /// <summary>
    /// Interaction logic for MiscSettings.xaml
    /// </summary>
    internal partial class MiscSettingsView : ContentControl
    {
        #region Fields

        private MiscSettingsPart m_part;

        #endregion

        #region Constructor

        public MiscSettingsView(MiscSettingsPart part)
        {
            InitializeComponent();
            m_part = part;
            DataContext = m_part;
        }

        #endregion

        #region private Methods

        private void SaveFileBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "XML file|*.xml";
            saveFileDialog.Title = WizardResource.Save;
            saveFileDialog.ShowDialog();

            // If the file name is not an empty string open it for saving.
            if (!string.IsNullOrEmpty(saveFileDialog.FileName))
            {
                m_part.OutputMappingsFile = saveFileDialog.FileName;
            }
        }

        private void SaveReportFolderBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (!string.IsNullOrEmpty(m_part.ReportFolder) && Directory.Exists(m_part.ReportFolder))
            {
                dialog.SelectedPath = m_part.ReportFolder;
            }
            try
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    m_part.ReportFolder = dialog.SelectedPath;
                }

            }
            catch (AccessViolationException)
            {
                m_part.Warning = "Unable to load the save dialog. Please try to launh save dialog again";
            }
        }

        #endregion
    }
}
