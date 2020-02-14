//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.IO;
    using System.Windows;
    using System.Windows.Controls;
    using Microsoft.Win32;

    /// <summary>
    /// Interaction logic for SelectDataSource.xaml
    /// </summary>
    internal partial class SelectDataSourceView : ContentControl
    {
        #region Fields

        private SelectDataSourcePart m_part;

        #endregion

        #region Constructor

        public SelectDataSourceView(SelectDataSourcePart part)
        {
            m_part = part;
            InitializeComponent();
            DataContext = part;
        }

        #endregion

        #region private methods

        private void ExcelFileBrowse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            Nullable<bool> result;
            dlg.Filter = "Excel files (*.xlsx,*.xls)|*.xlsx;*.xls|All files(*.*)|*.*";
            // Show open file dialog box
            try
            {
                result = dlg.ShowDialog();
            }
            catch (AccessViolationException)
            {
                m_part.Warning = "Unable to load the open dialog. Please try to launh open dialog again";
                return;
            }

            // Process open file dialog box results
            if (result == true)
            {
                m_part.ExcelFilePath = dlg.FileName;
            }
        }

        private void ListOfMHTsFilePathBrowse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            Nullable<bool> result;
            dlg.Filter = "Text files (*.txt)|*.txt|All files(*.*)|*.*";
            // Show open file dialog box
            try
            {
                result = dlg.ShowDialog();
            }
            catch (AccessViolationException)
            {
                m_part.Warning = "Unable to load the open dialog. Please try to launh open dialog again";
                return;
            }

            // Process open file dialog box results
            if (result == true)
            {
                m_part.ListOfMHTsFilePath = dlg.FileName;
            }
        }

        private void MHTFolderBrowse_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (!string.IsNullOrEmpty(m_part.MHTFolderPath) && Directory.Exists(m_part.MHTFolderPath))
            {
                dialog.SelectedPath = m_part.MHTFolderPath;
            }
            try
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    m_part.MHTFolderPath = dialog.SelectedPath;
                }

            }
            catch (AccessViolationException)
            {
                m_part.Warning = "Unable to load the open dialog. Please try to launh open dialog again";
            }
        }

        private void Preview_Click(object sender, RoutedEventArgs e)
        {
            m_part.PreviewDataSourceFile();
        }

        #endregion

        private void ExcelFileName_TextChangeAction(object sender, MessageEventArgs e)
        {
            TextBox textbox = sender as TextBox;
            m_part.ExcelFilePath = textbox.Text;
            m_part.CanNext = m_part.ValidatePartState();

            e.SetValues(m_part.ValidateExcelFilePath());
        }

        private void MHTFolderPath_TextChangeAction(object sender, MessageEventArgs e)
        {
            TextBox textbox = sender as TextBox;
            m_part.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(Part_PropertyChanged);
            m_part.MHTFolderPath = textbox.Text;
            if (!m_part.IsLoadingFiles)
            {
                m_part.CanNext = false;
            }
        }

        void Part_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "IsLoadingFiles")
            {
                if (!m_part.IsLoadingFiles && m_part.MHTCount > 0)
                {
                    m_part.CanNext = true;
                    m_part.PropertyChanged -= Part_PropertyChanged;
                }
            }
        }

        private void ListOfMHTsFilePath_TextChangeAction(object sender, MessageEventArgs e)
        {
            TextBox textbox = sender as TextBox;
            m_part.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(Part_PropertyChanged);
            m_part.ListOfMHTsFilePath = textbox.Text;
            if (!m_part.IsLoadingFiles)
            {
                m_part.CanNext = false;
            }
        }

        private void ExcelHeaderRow_TextChangeAction(object sender, MessageEventArgs e)
        {
            TextBox textbox = sender as TextBox;
            m_part.ExcelHeaderRow = textbox.Text;
            m_part.CanNext = m_part.ValidatePartState();

            e.SetValues(m_part.ValidateExcelHeaderRow());

        }
    }
}
