//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Windows;
    using System.Windows.Controls;
    using Microsoft.Win32;

    /// <summary>
    /// Interaction logic for FieldsSelectionView.xaml
    /// </summary>
    internal partial class FieldsSelectionView : ContentControl
    {
        #region Fields

        private FieldsSelectionPart m_part;

        #endregion

        #region Constructor

        public FieldsSelectionView(FieldsSelectionPart fieldsSelectionPart)
        {
            m_part = fieldsSelectionPart;
            DataContext = fieldsSelectionPart;
            InitializeComponent();
        }

        #endregion

        #region private methods

        private void MHTFileBrowse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            Nullable<bool> result;
            dlg.Filter = "MHT files (*.doc,*.docx,*.mht)|*.doc;*.docx;*.mht|All files(*.*)|*.*";
            // Show open file dialog box
            try
            {
                result = dlg.ShowDialog();
            }
            catch (AccessViolationException)
            {
                m_part.Warning = "Unable to load the open dialog. Please try again.";
                return;
            }

            // Process open file dialog box results
            if (result == true)
            {
                m_part.SourcePath = dlg.FileName;
            }
        }

        private void AddNewFieldNameButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(NewFieldNameInputBox.Text))
            {
                string newField = NewFieldNameInputBox.Text.Trim();

                if (m_part.AddFieldName(newField))
                {
                    m_part.Warning = null;
                    NewFieldNameInputBox.Text = string.Empty;
                }
                else
                {
                    m_part.Warning = "Could not be found following field in the sample mht file: " + newField;
                }
            }
        }

        private void DeleteFieldNameButton_Click(object sender, RoutedEventArgs e)
        {
            Button button = sender as Button;
            if (button != null && button.DataContext != null)
            {
                SourceField field = (SourceField)button.DataContext;
                m_part.Fields.Remove(field);
            }
        }

        private void Preview_Click(object sender, RoutedEventArgs e)
        {
            m_part.Preview();
        }

        #endregion
    }
}
