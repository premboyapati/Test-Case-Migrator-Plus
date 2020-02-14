//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Threading;
    using System.Windows;
    using System.Windows.Controls;
    using WizardResources = Microsoft.VisualStudio.TestTools.WorkItemMigrator.Resources;

    /// <summary>
    /// Interaction logic for WorkItemImportWizard.xaml
    /// </summary>
    internal partial class WorkItemImportWizard : Window, IDisposable
    {
        #region Fields

        /// <summary>
        /// Wizard member variable responsible for loading different Wizard Pages and controlling the Wizard parts
        /// </summary>
        private WizardController m_wizard;

        private AutoResetEvent m_stopMigrationEvent;
        bool m_forceClose;

        #endregion

        #region Constructor/Destructor

        /// <summary>
        /// Constructor to intialize the UI and Wizard
        /// </summary>
        public WorkItemImportWizard()
        {
            InitializeComponent();
            m_wizard = new WizardController();
            DataContext = m_wizard;
            MessageHelper.ShowMessageWindow += ShowMessage;
            App.UISynchronizationContext = SynchronizationContext.Current;
            MHTParser.STAThreadContext = SynchronizationContext.Current;
        }

        #endregion

        #region MessageBox Display/Callback

        private void ShowMessage(MessageEventArgs args)
        {
            App.CallMethodInUISynchronizationContext(ShowMessageInUIContext, args);
        }

        private void ShowMessageInUIContext(object value)
        {
            MessageEventArgs args = value as MessageEventArgs;
            MessageBox messageBox = new MessageBox(args);
            messageBox.Owner = this;
            messageBox.ShowDialog();
        }

        private void WindowClosingCallback(MessageEventArgs args)
        {
            if (args.ResponseDefinition == MessageResponseDefinition.FirstResponse)
            {
                args.Data = false;
            }
            else
            {
                args.Data = true;
            }
        }

        #endregion

        #region Evenet handlers

        /// <summary>
        /// Opens the Help
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            m_wizard.ShowHelp();
        }

        /// <summary>
        /// Go to Previous Wizard Page
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            m_wizard.GoBack();
        }

        /// <summary>
        /// Go to next Wizard page
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            m_wizard.GoNext();
        }

        /// <summary>
        /// Loads the confirm page
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            m_wizard.LoadConfirmPage();
        }

        /// <summary>
        /// Cancel and close the Wizard
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void WizardNavigationButton_Click(object sender, RoutedEventArgs e)
        {
            Button button = sender as Button;
            m_wizard.LoadWizardPart(button.DataContext as IWizardPart);
        }

        private void SaveandExitButton_Click(object sender, RoutedEventArgs e)
        {
            m_wizard.Save();
            m_forceClose = true;
            Close();
        }

        private void SaveAndMigrateButton_Click(object sender, RoutedEventArgs e)
        {
            m_wizard.SaveAndMigrate();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Wizard_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (String.CompareOrdinal(e.PropertyName, WizardController.IsMigrationCompletedTag) == 0)
            {
                m_stopMigrationEvent.Set();
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (m_wizard != null &&
                 m_wizard.WizardInfo != null &&
                 m_wizard.WizardInfo.ProgressPart != null)
            {
                e.Cancel = true;
                BringProgressPartAtTop();
                return;
            }

            if (m_forceClose)
            {
                return;
            }

            if (m_wizard.WizardPart == null ||
                (m_wizard.WizardPart.WizardPage == WizardPage.Welcome && m_wizard.WizardInfo.DataSourceParser == null))
            {
                return;
            }

            if (!m_wizard.IsSummaryPage)
            {
                MessageEventArgs args = new MessageEventArgs(MessageCategory.Warning,
                                                             WizardResources.CancelWizard_MessageTitle,
                                                             "You have pending changes. Choose 'Yes' to exit or 'No' to continue with the migration",
                                                             null,
                                                             WizardResources.YesButtonLabel,
                                                             null,
                                                             WizardResources.NoButtonLabel,
                                                             WindowClosingCallback);
                MessageHelper.ShowMessageWindow(args);
                e.Cancel = args.Data == null || (bool)args.Data;
            }
            else if (!m_wizard.IsMigrationCompleted)
            {
                MessageEventArgs args = new MessageEventArgs(MessageCategory.Warning,
                                             WizardResources.CloseMigration_MessageTitle,
                                             "Migration is in progress. Choose 'Yes' to exit or 'No' to continue with the migration",
                                             null,
                                             WizardResources.YesButtonLabel,
                                             null,
                                             WizardResources.NoButtonLabel,
                                             WindowClosingCallback);
                MessageHelper.ShowMessageWindow(args);
                e.Cancel = args.Data == null || (bool)args.Data;
            }
        }

        private void BringProgressPartAtTop()
        {
            m_wizard.WizardInfo.ProgressPart.Activate();
        }


        private void StartAgainButton_Click(object sender, RoutedEventArgs e)
        {
            m_wizard.Reset();
        }

        private void Window_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (m_wizard != null &&
                m_wizard.WizardInfo != null &&
                m_wizard.WizardInfo.ProgressPart != null)
            {
                BringProgressPartAtTop();
                e.Handled = true;
            }
        }

        private void Window_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (m_wizard != null &&
                m_wizard.WizardInfo != null &&
                m_wizard.WizardInfo.ProgressPart != null)
            {
                BringProgressPartAtTop();
                e.Handled = true;
            }
        }

        private void Window_FocusableChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (m_wizard != null &&
                m_wizard.WizardInfo != null &&
                m_wizard.WizardInfo.ProgressPart != null)
            {
                BringProgressPartAtTop();
            }
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            if (m_wizard != null &&
                m_wizard.WizardInfo != null &&
                m_wizard.WizardInfo.ProgressPart != null)
            {
                if (WindowState == System.Windows.WindowState.Minimized)
                {
                    m_wizard.WizardInfo.ProgressPart.WindowState = System.Windows.WindowState.Minimized;
                }
                else
                {
                    m_wizard.WizardInfo.ProgressPart.WindowState = System.Windows.WindowState.Normal;
                    BringProgressPartAtTop();
                }
            }
        }
        #endregion

        #region Overriden methods

        protected override void OnClosed(EventArgs e)
        {
            MessageHelper.ShowMessageWindow -= ShowMessage;
            Dispose();
            base.OnClosed(e);
        }

        #endregion

        #region IDisposible Implementation

        /// <summary>
        /// Disposes  the Wizard Controller
        /// </summary>
        public void Dispose()
        {
            m_wizard.Dispose();
        }

        #endregion
    }
}