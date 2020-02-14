//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for MapColumns.xaml
    /// </summary>
    internal partial class FieldMappingView : ContentControl
    {
        #region Fields
        FieldMappingPart m_part;
        #endregion
        #region Constructor

        public FieldMappingView(FieldMappingPart part)
        {
            m_part = part;
            InitializeComponent();
            this.IsVisibleChanged += new System.Windows.DependencyPropertyChangedEventHandler(FieldMappingView_IsVisibleChanged);
            DataContext = part;
        }

        private void FieldMappingView_IsVisibleChanged(object sender, System.Windows.DependencyPropertyChangedEventArgs e)
        {
            if (Visibility == System.Windows.Visibility.Visible)
            {
                m_part.Warning = m_part.Warning;
            }
        }

        #endregion

        #region private methods

        private void WIFieldComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox wiFieldBox = sender as ComboBox;
            FieldMappingRow row = wiFieldBox.DataContext as FieldMappingRow;
            if (wiFieldBox.SelectedItem != null)
            {
                row.TFSField = wiFieldBox.SelectedItem.ToString();
            }
        }

        #endregion
    }
}
