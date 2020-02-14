//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for RelationshipsMappingView.xaml
    /// </summary>
    internal partial class LinkingView : ContentControl
    {
        #region Fields

        LinkingPart m_part;

        #endregion

        public LinkingView(LinkingPart part)
        {
            m_part = part;
            DataContext = part;
            InitializeComponent();
        }

        #region private methods

        private void LinkedFieldComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            LinkingRow row = comboBox.DataContext as LinkingRow;
            if (comboBox.SelectedItem != null)
            {
                row.LinkedField = comboBox.SelectedItem.ToString();
            }
        }

        #endregion

    }
}
