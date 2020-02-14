//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for MapValues.xaml
    /// </summary>
    internal partial class DataMappingView : ContentControl
    {
        #region Fields

        private DataMappingPart m_part;

        #endregion

        #region Constructor

        public DataMappingView(DataMappingPart part)
        {
            InitializeComponent();
            DataContext = part;
            m_part = part;
        }

        #endregion

        #region Private Methods

        private void AddDataMappingRowButton_Click(object sender, RoutedEventArgs e)
        {
            m_part.AddEditableDataMappingRow();
        }

        #endregion
    }
}
