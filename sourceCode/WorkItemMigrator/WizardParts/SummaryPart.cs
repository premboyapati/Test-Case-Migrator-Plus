//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.IO;

    /// <summary>
    /// Wizard part Reponsible for all Wizard Actions Possible after the Wizard Configuration
    /// It can perform following Actions:-
    /// a) Parse Data Source
    /// b) Save Settings File
    /// c) Migrate Workitems
    /// d) public Migration Report
    /// </summary>
    internal class SummaryPart : BaseWizardPart
    {
        #region Fields

        // Member Variables needed for Data Binding
        private WizardActionState m_wizardActionsState;
        private int m_totalWorkItemsCount;
        private int m_failedWorkItemsCount;
        private int m_passedWorkItemsCount;
        private int m_warningWorkItemsCount;
        private int m_currentWizardActionNumber;
        private bool m_isReportPublished;
        private bool m_isProcessed;
        private bool m_isMigrationStopped = false;
        private bool m_isLinksReportPublished;

        #endregion

        #region Constants

        // Used for sending Notification to wizard Controller that all wizard actions are performed
        public const string MigrationStateTag = "MigrationState";

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        public SummaryPart()
        {
            Header = Resources.MigrationProgress_Header;
            Description = string.Empty;
            CanBack = false;
            CanNext = false;
            WizardPage = WizardPage.Summary;
            MigrationState = WizardActionState.Pending;
            WizardActions = new ObservableCollection<WizardAction>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Tells the State of Migration Part
        /// </summary>
        public WizardActionState MigrationState
        {
            get
            {
                return m_wizardActionsState;
            }
            set
            {
                m_wizardActionsState = value;
                NotifyPropertyChanged(MigrationStateTag);
                NotifyPropertyChanged("MigrationStatus");
            }
        }

        /// <summary>
        /// The string representation of current state of Wizard Part
        /// </summary>
        public string MigrationStatus
        {
            get
            {
                switch (MigrationState)
                {
                    case WizardActionState.InProgress:
                        return Resources.InProgress;

                    case WizardActionState.Success:
                        return Resources.Success;

                    case WizardActionState.Failed:
                        return Resources.Failed;

                    case WizardActionState.Warning:
                        return Resources.Warning;

                    default:
                        return string.Empty;
                }
            }
        }

        /// <summary>
        /// Total number of workitems which are going to migrate
        /// </summary>
        public int TotalWorkItems
        {
            get
            {
                return m_totalWorkItemsCount;
            }
            private set
            {
                m_totalWorkItemsCount = value;
                NotifyPropertyChanged("TotalWorkItems");
            }
        }

        /// <summary>
        /// Total number of Failed Workitems
        /// </summary>
        public int FailedWorkItemsCount
        {
            get
            {
                return m_failedWorkItemsCount;
            }
            set
            {
                m_failedWorkItemsCount = value;
                NotifyPropertyChanged("FailedWorkItemsCount");
            }
        }

        /// <summary>
        /// Total number of Passed Workitems
        /// </summary>
        public int PassedWorkItemsCount
        {
            get
            {
                return m_passedWorkItemsCount;
            }
            set
            {
                m_passedWorkItemsCount = value;
                NotifyPropertyChanged("PassedWorkItemsCount");
            }
        }

        /// <summary>
        /// Total number of warning workitems
        /// </summary>
        public int WarningWorkItemsCount
        {
            get
            {
                return m_warningWorkItemsCount;
            }
            set
            {
                m_warningWorkItemsCount = value;
                NotifyPropertyChanged("WarningWorkItemsCount");
            }
        }

        /// <summary>
        /// Current Wizard Action performing
        /// </summary>
        public WizardAction CurrentAction
        {
            get
            {
                if (m_currentWizardActionNumber >= 0 &&
                    WizardActions != null &&
                    m_currentWizardActionNumber < WizardActions.Count)
                {
                    return WizardActions[m_currentWizardActionNumber];
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// List of Wizard Action which have to perform
        /// </summary>
        public ObservableCollection<WizardAction> WizardActions
        {
            get;
            private set;
        }

        /// <summary>
        /// Is Report successfully published
        /// </summary>
        public bool IsReportPublished
        {
            get
            {
                return m_isReportPublished;
            }
            private set
            {
                m_isReportPublished = value;
                NotifyPropertyChanged("IsReportPublished");
            }
        }

        public bool IsLinksReportPublished
        {
            get
            {
                return m_isLinksReportPublished;
            }
            set
            {
                m_isLinksReportPublished = false;
                if (value && m_wizardInfo.LinksManager != null)
                {
                    string linksReportFilePath = Path.Combine(Path.GetDirectoryName(WizardInfo.LinksManager.LinksFilePath),
                                                              LinksManager.XLReportFileName);

                    if (File.Exists(linksReportFilePath))
                    {
                        m_isLinksReportPublished = true;
                    }
                }
                NotifyPropertyChanged("IsLinksReportPublished");
            }
        }

        /// <summary>
        /// is Migration in progress
        /// </summary>
        public bool IsMigrating
        {
            get
            {
                return CurrentAction != null &&
                       CurrentAction.ActionName == WizardActionName.MigrateWorkItems;
            }
        }

        #endregion

        #region public methods

        public override void Clear()
        {
            base.Clear();
            m_isProcessed = false;
            ClearWizardActions();
            MigrationState = WizardActionState.Pending;
            TotalWorkItems = 0;
            FailedWorkItemsCount = 0;
            WarningWorkItemsCount = 0;
            PassedWorkItemsCount = 0;
            IsReportPublished = false;
            IsLinksReportPublished = false;
        }

        public void ShowLinkingReport()
        {
            if (WizardInfo.LinksManager != null &&
                !string.IsNullOrEmpty(WizardInfo.LinksManager.LinksFilePath))
            {
                string linksReportFilePath = Path.Combine(Path.GetDirectoryName(WizardInfo.LinksManager.LinksFilePath),
                                                LinksManager.XLReportFileName);
                try
                {
                    if (File.Exists(linksReportFilePath))
                    {
                        ExcelParser.OpenWorkSheet(linksReportFilePath, "Migration-Summary");
                    }
                }
                catch (WorkItemMigratorException we)
                {
                    Warning = "Unable to open Linking Report File" + we.Args.Title;
                }
            }
        }

        /// <summary>
        /// Resets the Migration State
        /// </summary>
        public override void Reset()
        {
            // Only start the migration if it is active Wizard part and it is not already migrated
            if (IsActiveWizardPart && !m_isProcessed)
            {
                MigrationState = WizardActionState.InProgress;

                m_isProcessed = true;

                // Initialize All Wizard Actions which have to perform
                InitializeActions();

                // Start the First Action if it exists and sets the Description of the Wizard Part to be in progress
                if (WizardActions.Count > 0)
                {
                    CurrentAction.Start();
                    Description = Resources.MigrationInProgressText;
                }
            }
        }

        /// <summary>
        /// Stops the Migration Wizard Action
        /// </summary>
        public void StopMigration()
        {
            // If current action is Migrate action then stop it and start the next one. Also send the notifications
            if (!m_isMigrationStopped && CurrentAction.ActionName == WizardActionName.MigrateWorkItems)
            {
                m_isMigrationStopped = true;
                WizardActions[m_currentWizardActionNumber].Stop();
            }
        }

        /// <summary>
        /// Stops all the Wizard Actions that are in queue or in progress
        /// </summary>
        public void StopAll()
        {
            for (int i = m_currentWizardActionNumber; i < WizardActions.Count; i++)
            {
                WizardActions[i].Stop();
            }
            MigrationState = WizardActionState.Stopped;
        }

        /// <summary>
        ///  There is no need of Updating the Wizard part's State so just return true
        /// </summary>
        /// <returns></returns>
        public override bool UpdateWizardPart()
        {
            return true;
        }

        /// <summary>
        /// Shows Published Report
        /// </summary>
        /// <param name="workSheetName"></param>
        public void ShowReport(string workSheetName)
        {
            try
            {
                switch (m_wizardInfo.DataSourceType)
                {
                    case DataSourceType.Excel:
                        ExcelParser.OpenWorkSheet(m_wizardInfo.Reporter.ReportFile, workSheetName);
                        break;

                    case DataSourceType.MHT:
                        Process.Start(m_wizardInfo.Reporter.ReportFile);
                        break;

                    default:
                        throw new InvalidEnumArgumentException("Invalid Enum Value");
                }
            }
            catch (WorkItemMigratorException te)
            {
                // Display Message if some error occured while trying to show the report
                MessageHelper.ShowMessageWindow(te.Args);
            }
        }

        #endregion

        #region protected/private methods

        protected override bool IsInitializationRequired(WizardInfo state)
        {
            return true;
        }

        public override bool ValidatePartState()
        {
            return true;
        }

        /// <summary>
        /// Can Controller show Migration Wizard part
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
        protected override bool CanInitializeWizardPage(WizardInfo info)
        {
            m_canShow = true;

            foreach (IWorkItemField field in info.WorkItemGenerator.TfsNameToFieldMapping.Values)
            {
                if (field.IsMandatory && string.IsNullOrEmpty(field.SourceName))
                {
                    Warning = Resources.MandatoryFieldsNotMappedErrorTitle;
                    m_canShow = false;
                    break;
                }
            }
            return m_canShow;
        }

        /// <summary>
        /// Initializes all Wizard Actions which are possible
        /// </summary>
        private void InitializeActions()
        {
            if (m_wizardInfo.DataSourceType == DataSourceType.Excel)
            {
                // Parse Data Source Action
                WizardAction parseDataSourceAction = new ParseDataSourceAction(m_wizardInfo);
                parseDataSourceAction.PropertyChanged += new PropertyChangedEventHandler(WizardAction_PropertyChanged);
                App.CallMethodInUISynchronizationContext(AddWizardAction, parseDataSourceAction);
            }

            // Save Settings Action : Only possible if save mapping file location is provided
            if (!string.IsNullOrEmpty(m_wizardInfo.OutputSettingsFilePath))
            {
                WizardAction saveSettingsAction = new SaveSettingsAction(m_wizardInfo);
                saveSettingsAction.PropertyChanged += new PropertyChangedEventHandler(WizardAction_PropertyChanged);
                App.CallMethodInUISynchronizationContext(AddWizardAction, saveSettingsAction);
            }

            // Migrate Workitems Action
            WizardAction migrateWorkItemsAction = new MigrateWorkItemsAction(m_wizardInfo);
            migrateWorkItemsAction.PropertyChanged += new PropertyChangedEventHandler(WizardAction_PropertyChanged);
            App.CallMethodInUISynchronizationContext(AddWizardAction, migrateWorkItemsAction);

            if (m_wizardInfo.RelationshipsInfo != null &&
                m_wizardInfo.IsLinking &&
                !string.IsNullOrEmpty(m_wizardInfo.RelationshipsInfo.SourceIdField) &&
                String.CompareOrdinal(m_wizardInfo.RelationshipsInfo.SourceIdField, Resources.SelectPlaceholder) != 0)
            {
                WizardAction relationshipsAction = new ProcessLinksAction(m_wizardInfo);
                relationshipsAction.PropertyChanged += new PropertyChangedEventHandler(WizardAction_PropertyChanged);
                App.CallMethodInUISynchronizationContext(AddWizardAction, relationshipsAction);
            }


            // Publish Report Action
            if (m_wizardInfo.Reporter != null && !string.IsNullOrEmpty(m_wizardInfo.Reporter.ReportFile))
            {
                WizardAction publishReportAction = new PublishReportAction(m_wizardInfo);
                publishReportAction.PropertyChanged += new PropertyChangedEventHandler(WizardAction_PropertyChanged);
                App.CallMethodInUISynchronizationContext(AddWizardAction, publishReportAction);
            }

            m_currentWizardActionNumber = 0;
        }

        private void AddWizardAction(object obj)
        {
            WizardAction action = obj as WizardAction;
            if (action != null)
            {
                WizardActions.Add(action);
            }
        }

        /// <summary>
        /// Migration Wizard Part listens to the Wizard Action's Notifications. 
        /// This is handler for those notifications
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void WizardAction_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            // If Notification is of Migration of a singleWorkitem
            if (e.PropertyName == MigrateWorkItemsAction.Passed ||
                e.PropertyName == MigrateWorkItemsAction.Failed ||
                e.PropertyName == MigrateWorkItemsAction.Warning)
            {
                // Then update teh counters
                UpdateCounters(e.PropertyName);
            }
            // else if Wizard Action state is changed and it is changed to either Success/Warning/Failed
            else if (String.CompareOrdinal(e.PropertyName, "State") == 0 &&
                    (CurrentAction.State == WizardActionState.Success ||
                    CurrentAction.State == WizardActionState.Failed ||
                    CurrentAction.State == WizardActionState.Stopped ||
                    CurrentAction.State == WizardActionState.Warning))
            {
                // then Update Status of MigrationPart(Counters/Flags etc)
                UpdateStatus(CurrentAction.State);

                CurrentAction.Dispose();

                // Change current Action to next action in queue
                m_currentWizardActionNumber++;
                NotifyPropertyChanged("IsMigrating");

                // If there is any action pending then start it else update the Migration Part's satte and update its description
                if (m_currentWizardActionNumber < WizardActions.Count)
                {
                    CurrentAction.Start();
                }
                else
                {
                    UpdateWizardActionState();
                    switch (MigrationState)
                    {
                        case WizardActionState.Success:

                            Description = Resources.MigrationSuccessfulText;
                            break;

                        case WizardActionState.Warning:
                            Description = "Work items were migrated successfully but with warnings.";
                            break;
                        case WizardActionState.Failed:
                            Description = "Some error occured during migration of work items.";
                            break;

                        default:
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Updates Passed/Warning/Failed Counters 
        /// </summary>
        /// <param name="propertyName"></param>
        private void UpdateCounters(string propertyName)
        {
            switch (propertyName)
            {
                case MigrateWorkItemsAction.Passed:
                    PassedWorkItemsCount++;
                    break;

                case MigrateWorkItemsAction.Failed:
                    FailedWorkItemsCount++;
                    break;

                case MigrateWorkItemsAction.Warning:
                    WarningWorkItemsCount++;
                    break;

                default:
                    break;
            }
        }

        /// <summary>
        /// Updates Migration Part's State
        /// </summary>
        private void UpdateWizardActionState()
        {
            WizardActionState state = WizardActionState.Success;
            foreach (WizardAction action in WizardActions)
            {
                if (action.State == WizardActionState.Failed || action.State == WizardActionState.Stopped)
                {
                    state = WizardActionState.Failed;
                    break;
                }
                if (action.State == WizardActionState.Warning)
                {
                    state = WizardActionState.Warning;
                }
            }
            MigrationState = state;
        }

        /// <summary>
        /// Updates Flags
        /// </summary>
        /// <param name="wizardActionState"></param>
        private void UpdateStatus(WizardActionState wizardActionState)
        {
            // If Publish report action is performed successfully
            if (CurrentAction.ActionName == WizardActionName.PublishReport &&
                wizardActionState == WizardActionState.Success)
            {
                // then sets IsReportPublished flag as true
                IsReportPublished = true;
                IsLinksReportPublished = true;

            }
            // else if Data Source is parsed then update the Counter for Total number of workitems
            else if (CurrentAction.ActionName == WizardActionName.ParseDataSource)
            {
                TotalWorkItems = m_wizardInfo.Migrator.SourceWorkItems.Count;
            }
        }

        private void ClearWizardActions()
        {
            App.CallMethodInUISynchronizationContext(ClearWizardActionsInUIContext, null);
        }

        private void ClearWizardActionsInUIContext(object state)
        {
            if (WizardActions != null)
            {
                WizardActions.Clear();
            }
        }

        #endregion

    }
}
