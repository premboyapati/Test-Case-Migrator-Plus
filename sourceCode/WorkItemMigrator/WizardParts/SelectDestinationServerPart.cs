//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.ObjectModel;
    using System.Configuration;
    using System.Security;
    using Microsoft.TeamFoundation;

    /// <summary>
    /// Implemantation of Wizard Part reponsible for getting TFS Details
    /// </summary>
    internal class SelectDestinationServerPart : BaseWizardPart
    {
        #region Fields

        // Fields for Data Binding
        private string m_server;
        private string m_project;
        private string m_selectedWorkItemType;
        private DataSourceType m_dataSourceType;

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor for initializing the Wizard Part
        /// </summary>
        public SelectDestinationServerPart()
        {
            Header = Resources.SelectDestinationServer_Header;
            Description = Resources.SelectDestinationServer_Description;
            CanBack = true;
            WizardPage = WizardPage.SelectDestinationServer;
            WorkItemTypes = new ObservableCollection<string>();
            CanShow = false;
        }

        #endregion

        #region Properties

        /// <summary>
        /// TFS Server to be connected
        /// </summary>
        public string Server
        {
            get
            {
                if (string.IsNullOrEmpty(m_server))
                {
                    return Resources.SelectServerWatermarkText;
                }
                return m_server;
            }
            set
            {
                m_server = value;
                NotifyPropertyChanged("Server");
            }
        }

        /// <summary>
        /// TFS Project to be connected
        /// </summary>
        public string Project
        {
            get
            {
                return m_project;
            }
            set
            {
                m_project = value;
                NotifyPropertyChanged("Project");

                // Connect with the Server/project Collection and Fill List of workitem types present in that server
                App.CallMethodInUISynchronizationContext(FillWorkItemTypeNames, null);
            }
        }

        /// <summary>
        /// List of workitem types in the current Server/project Collection
        /// </summary>
        public ObservableCollection<string> WorkItemTypes
        {
            get;
            private set;
        }

        /// <summary>
        /// Selected Workitem type
        /// </summary>
        public string SelectedWorkItemType
        {
            get
            {
                return m_selectedWorkItemType;
            }
            set
            {
                m_selectedWorkItemType = value;
                NotifyPropertyChanged("SelectedWorkItemType");
            }
        }

        #endregion

        #region public methods

        /// <summary>
        /// Resets the Server project to empty strings
        /// </summary>
        public override void Reset()
        {
            Server = string.Empty;
            Project = string.Empty;

            try
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                if (config == null || config.AppSettings == null || config.AppSettings.Settings.Count == 0)
                {
                    return;
                }
                if (config.AppSettings.Settings["CollectionURL"] != null && !string.IsNullOrEmpty(config.AppSettings.Settings["CollectionURL"].Value))
                {
                    Server = config.AppSettings.Settings["CollectionURL"].Value;
                }

                if (config.AppSettings.Settings["DefaultProject"] != null && !string.IsNullOrEmpty(config.AppSettings.Settings["DefaultProject"].Value))
                {
                    Project = config.AppSettings.Settings["DefaultProject"].Value;
                }
            }
            catch (TeamFoundationServerException)
            {
                Server = string.Empty;
                Project = string.Empty;
            }
        }

        /// <summary>
        /// Updates the Server/project details in the Wizard Info 
        /// </summary>
        /// <returns></returns>
        public override bool UpdateWizardPart()
        {
            // If Wizard Part State is not valid return false
            if (!ValidatePartState())
            {
                return false;
            }

            if (!IsUpdationRequired())
            {
                return true;
            }


            m_dataSourceType = m_wizardInfo.DataSourceType;
            if (m_wizardInfo.DataSourceType == DataSourceType.Excel)
            {
                m_wizardInfo.WorkItemGenerator.AddTestStepsField = true;
            }
            else
            {
                m_wizardInfo.WorkItemGenerator.AddTestStepsField = false;
            }
            try
            {
                m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName = SelectedWorkItemType;
            }
            catch (ArgumentException)
            {
                Warning = "Unexpected Action";
                return false;
            }

            if (m_wizardInfo.Migrator.SourceNameToFieldMapping != null)
            {
                m_wizardInfo.Migrator.SourceNameToFieldMapping.Clear();
            }

            //return true as Updation and TFS Server/project Connection was successful
            return true;
        }

        #endregion

        #region protected/private methods

        /// <summary>
        /// Validates Wizard Part State for Server/Project Values
        /// </summary>
        /// <returns></returns>
        public override bool ValidatePartState()
        {
            Warning = null;
            // Server/Project values can't be null
            if (string.IsNullOrEmpty(Server) || string.IsNullOrEmpty(Project) || string.IsNullOrEmpty(SelectedWorkItemType))
            {
                Warning = Resources.ServerNotSpecifiedErrorTitle;
                return false;
            }
            return true;
        }

        private void FillWorkItemTypeNames(object value)
        {
            WorkItemTypes.Clear();

            if (string.IsNullOrEmpty(Server) || string.IsNullOrEmpty(Project))
            {
                return;
            }

            try
            {
                using (new AutoWaitCursor())
                {
                    m_wizardInfo.WorkItemGenerator = new WorkItemGenerator(Server, Project);

                    foreach (string wiType in m_wizardInfo.WorkItemGenerator.WorkItemTypeNames)
                    {
                        WorkItemTypes.Add(wiType);
                    }
                    SelectedWorkItemType = m_wizardInfo.WorkItemGenerator.DefaultWorkItemTypeName;
                }
                CanNext = ValidatePartState();
            }
            catch (TeamFoundationServerException tfe)
            {
                Warning = tfe.Message;
            }
            catch (SecurityException se)
            {
                Warning = se.Message;
            }
            catch (System.Net.WebException webEx)
            {
                Warning = webEx.Message;
            }
        }
        /// <summary>
        /// We can always initialize it
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
        protected override bool CanInitializeWizardPage(WizardInfo info)
        {
            if (info.DataSourceParser == null)
            {
                m_canShow = false;
            }
            else
            {
                m_canShow = true;
            }
            return m_canShow;
        }

        protected override bool IsUpdationRequired()
        {
            if (m_dataSourceType != m_wizardInfo.DataSourceType ||
                m_prerequisite.IsServerConnectionModified(Server, Project, SelectedWorkItemType))
            {
                return true;
            }
            return false;
        }

        #endregion

    }
}
