//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for MigrationProgress.xaml
    /// </summary>
    internal partial class SummaryView : ContentControl
    {
        #region Fields

        SummaryPart m_part;

        #endregion

        #region Constructor

        public SummaryView(SummaryPart part)
        {
            InitializeComponent();
            m_part = part;
            DataContext = part;
        }

        #endregion

        #region private methods

        private void Hyperlink_Click(object sender, RoutedEventArgs e)
        {
            string workSheetName = null;
            if (sender == ErrorLink)
            {
                workSheetName = ExcelReporter.ErrorSheetName;
            }
            else if (sender == WarningLink)
            {
                workSheetName = ExcelReporter.WarningSheetName;
            }
            else if (sender == PassedLink)
            {
                workSheetName = ExcelReporter.SuccessSheetName;
            }

            m_part.ShowReport(workSheetName);
        }

        private void ShowLinkingReport_Click(object sender, RoutedEventArgs e)
        {
            m_part.ShowLinkingReport();
        }

        private void StopButton_Click(object sender, RoutedEventArgs e)
        {
            m_part.StopMigration();
        }

        #endregion
    }
}
