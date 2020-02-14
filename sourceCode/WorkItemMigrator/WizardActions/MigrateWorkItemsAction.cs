//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Globalization;
    using WorkItemMigrator = Microsoft.VisualStudio.TestTools.WorkItemMigrator;

    /// <summary>
    /// Wizard Action class Responsible for Migration of Data Source Workitems to TFS Workitems
    /// </summary>
    internal class MigrateWorkItemsAction : WizardAction
    {
        #region Fields

        // Counters needed for concluding the outcome of the Action
        private int m_count = 0;
        private int m_total;
        private int m_error;
        private int m_warning;

        #endregion

        #region Constants

        // Constants needed for sending Notifications to Wizard part listning to this action
        public const string Passed = "Passed";
        public const string Failed = "Failed";
        public const string Warning = "Warning";

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor to intialize the action. it also supports Cancellation in course of operation
        /// </summary>
        /// <param name="wizardInfo"></param>
        public MigrateWorkItemsAction(WizardInfo wizardInfo)
            : base(wizardInfo)
        {
            Description = Resources.MigrateWorkItemsAction_Description;
            ActionName = WizardActionName.MigrateWorkItems;
            m_worker.WorkerSupportsCancellation = true;
        }

        #endregion

        #region Overriden methods

        /// <summary>
        /// The Main Working Function
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        protected override WizardActionState DoWork(DoWorkEventArgs e)
        {
            // initializing the Counters and Data Structures
            m_total = m_wizardInfo.Migrator.SourceWorkItems.Count;

            m_error = 0;
            m_warning = 0;

            m_wizardInfo.Migrator.PreMigration = PreMigrationEvent;
            m_wizardInfo.Migrator.PostMigration = PostMigrationEvent;

            if (m_wizardInfo.IsLinking)
            {
                m_wizardInfo.LinksManager = new LinksManager(m_wizardInfo);
            }

            m_wizardInfo.ResultWorkItems = m_wizardInfo.Migrator.Migrate(m_wizardInfo.WorkItemGenerator);

            if (m_wizardInfo.IsLinking)
            {
                foreach (SourceWorkItem sourceWI in m_wizardInfo.ResultWorkItems)
                {
                    if (sourceWI is SkippedSourceWorkItem)
                    {
                        string category = m_wizardInfo.WorkItemGenerator.WorkItemTypeToCategoryMapping[m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName];
                        int tfsId = m_wizardInfo.LinksManager.WorkItemCategoryToIdMappings[category][sourceWI.SourceId].TfsId;

                        AddWorkItemInTestSuites(tfsId, sourceWI.TestSuites);
                    }
                }
            }

            // Calculate Final State Based on counters value and return it
            return CalculateState();
        }

        #endregion

        #region private methods

        private bool PreMigrationEvent(ISourceWorkItem sourceWorkItem, IWorkItem workItem)
        {
            string category = m_wizardInfo.WorkItemGenerator.WorkItemTypeToCategoryMapping[m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName];
            if (m_wizardInfo.IsLinking &&
                m_wizardInfo.LinksManager != null &&
                m_wizardInfo.LinksManager.WorkItemCategoryToIdMappings.ContainsKey(category) &&
                m_wizardInfo.LinksManager.WorkItemCategoryToIdMappings[category].ContainsKey(sourceWorkItem.SourceId) &&
                m_wizardInfo.LinksManager.WorkItemCategoryToIdMappings[category][sourceWorkItem.SourceId].TfsId != -1)
            {
                m_count++;
                UpdateCounters(typeof(WarningSourceWorkItem));
                return false;
            }
            return true;
        }

        private bool PostMigrationEvent(ISourceWorkItem sourceWorkItem)
        {
            // If this action is cancelled then update the message and return Stopped State
            if (m_worker.CancellationPending)
            {
                Message = Resources.UserInterruptionText;
                return false;
            }

            if (sourceWorkItem is SkippedSourceWorkItem)
            {
                return true;
            }

            // updating the counters
            UpdateCounters(sourceWorkItem.GetType());

            PassedSourceWorkItem passedWorkItem = sourceWorkItem as PassedSourceWorkItem;

            if (passedWorkItem != null)
            {
                AddWorkItemInTestSuites(passedWorkItem.TFSId, passedWorkItem.TestSuites);
            }

            if (m_wizardInfo.IsLinking)
            {
                var workItemStatus = new WorkItemMigrationStatus();
                workItemStatus.SourceId = sourceWorkItem.SourceId;
                workItemStatus.SessionId = m_wizardInfo.LinksManager.SessionId;
                workItemStatus.WorkItemType = m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName;

                if (passedWorkItem != null)
                {
                    workItemStatus.Status = WorkItemMigrator.Status.Passed;
                    workItemStatus.TfsId = passedWorkItem.TFSId;
                }
                else
                {
                    var failedWorkItem = sourceWorkItem as FailedSourceWorkItem;
                    if (failedWorkItem != null)
                    {
                        workItemStatus.Status = WorkItemMigrator.Status.Failed;
                        workItemStatus.TfsId = -1;
                        workItemStatus.Message = failedWorkItem.Error;
                    }
                }
                if (!string.IsNullOrEmpty(workItemStatus.SourceId))
                {
                    m_wizardInfo.LinksManager.UpdateIdMapping(workItemStatus.SourceId, workItemStatus);
                }
            }

            return true;
        }

        private void AddWorkItemInTestSuites(int tfsId, IList<string> testSuites)
        {
            foreach (string testSuite in testSuites)
            {
                m_wizardInfo.WorkItemGenerator.AddWorkItemToTestSuite(tfsId, testSuite);
            }
        }

        /// <summary>
        /// Calculate Wizard Action State after the action is completed based on counters value
        /// </summary>
        /// <returns></returns>
        private WizardActionState CalculateState()
        {
            if (m_worker.CancellationPending)
            {
                return WizardActionState.Stopped;
            }

            // if migration of all workitems is failed then set Failed Message and state as Failed
            if (m_error == m_count)
            {
                Message = "None of the test cases could be migrated";
                return WizardActionState.Failed;
            }
            else if (m_warning > 0 && m_error == 0)
            {
                Message = "Migrated successfully with warnings";
                return WizardActionState.Warning;
            }
            else if (m_error > 0)
            {
                Message = "Some of the test cases could not be migrated";
                return WizardActionState.Warning;
            }
            // else if all of them are successfully migrated then retrun state as Success
            else if (m_error == 0 && m_warning == 0)
            {
                return WizardActionState.Success;
            }
            // else Set warning Message and return state as Warning
            else
            {
                Message = Resources.MigrateWorkItemsAction_WarningText;
                return WizardActionState.Warning;
            }
        }

        /// <summary>
        /// Updates Counter & Progess Message and send Notifications after every attempt of migration
        /// </summary>
        /// <param name="type"></param>
        private void UpdateCounters(Type type)
        {
            if (type.Equals(typeof(PassedSourceWorkItem)))
            {
                NotifyPropertyChanged(Passed);
            }
            else if (type.Equals(typeof(WarningSourceWorkItem)))
            {
                NotifyPropertyChanged(Warning);
                m_warning++;
            }
            else if (type.Equals(typeof(FailedSourceWorkItem)))
            {
                NotifyPropertyChanged(Failed);
                m_error++;
            }

            m_count++;
            if (m_count < m_total)
            {
                Message = String.Format(CultureInfo.CurrentCulture,
                                        Resources.MigrateWorkItemsAction_InProgressText,
                                        m_count,
                                        m_total);
            }
            else
            {
                Message = string.Empty;
            }
        }

        #endregion
    }
}
