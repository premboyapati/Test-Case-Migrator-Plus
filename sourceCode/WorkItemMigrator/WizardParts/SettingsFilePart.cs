//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.IO;
    using System.Xml;

    /// <summary>
    /// Wizard Part for Loading Settings Files
    /// </summary>
    internal class SettingsFilePart : BaseWizardPart
    {
        #region Fields

        /// <summary>
        /// Member Varaibles for Data binding
        /// </summary>
        private bool m_loadSettings;
        private string m_settingsFilePath;

        #endregion

        #region Constants

        private const string XMLExtension = ".xml";

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        public SettingsFilePart()
        {
            Header = Resources.MappingsFile_Header;
            Description = Resources.MappingsFile_Description;
            CanBack = true;
            WizardPage = WizardPage.SettingsFile;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Is settings File To be loaded
        /// </summary>
        public bool LoadSettings
        {
            get
            {
                return m_loadSettings;
            }
            set
            {
                m_loadSettings = value;
                EvaluateCanConfirm();
                NotifyPropertyChanged("LoadSettings");
            }
        }

        /// <summary>
        /// File Path of the Settings File
        /// </summary>
        public string SettingsFilePath
        {
            get
            {
                return m_settingsFilePath;
            }
            set
            {
                if (String.CompareOrdinal(m_settingsFilePath, value) != 0)
                {
                    m_settingsFilePath = value;
                    EvaluateCanConfirm();
                    NotifyPropertyChanged("SettingsFilePath");
                }
            }
        }

        public override bool CanNext
        {
            get
            {
                return m_canNext;
            }
            set
            {
                if (m_canNext != value && WizardInfo != null)
                {
                    m_canNext = value;
                    WizardInfo.CanConfirm = m_canNext;
                    NotifyPropertyChanged(BaseWizardPart.CanNextPropertyName);
                }
            }
        }

        #endregion

        #region public methods

        /// <summary>
        /// Resets the Settings Wizard Part to start from blank settings File
        /// </summary>
        public override void Reset()
        {
            CanNext = ValidatePartState();
            LoadSettings = false;
            SettingsFilePath = m_wizardInfo.InputSettingsFilePath;
            m_wizardInfo.InputSettingsFilePath = string.Empty;
            m_wizardInfo.Migrator.SourceNameToFieldMapping = null;
        }

        /// <summary>
        /// Updates the Wizard Info and Loads Settings if Settings File is specified
        /// </summary>
        /// <returns></returns>
        public override bool UpdateWizardPart()
        {
            // If Wizard Part State is not valid return false
            if (!ValidatePartState())
            {
                return false;
            }

            try
            {
                // Load Settings if Settings File is specified and it is changed from last time
                if (LoadSettings)
                {
                    if (String.CompareOrdinal(SettingsFilePath, m_wizardInfo.InputSettingsFilePath) != 0)
                    {
                        m_wizardInfo.LoadSettings(SettingsFilePath);
                        m_wizardInfo.InputSettingsFilePath = SettingsFilePath;
                    }
                }
                // else just updates the Input Settings FilePath in WizardInfo
                else if (!string.IsNullOrEmpty(m_wizardInfo.InputSettingsFilePath))
                {
                    m_wizardInfo.InputSettingsFilePath = null;
                    m_wizardInfo.DataSourceParser.ParseDataSourceFieldNames();
                    m_wizardInfo.Migrator.SourceNameToFieldMapping = null;
                }
                m_prerequisite.Save();
            }
            catch (XmlException)
            {
                m_wizardInfo.ProgressPart = null;
                Warning = Resources.InputSettingsFileError;
                return false;
            }
            catch (WorkItemMigratorException te)
            {
                m_wizardInfo.ProgressPart = null;
                Warning = te.Args.Title;
                return false;
            }
            return true;
        }

        public MessageEventArgs ValidateSettingsFilePath()
        {
            // If Setting File is Not Valid XML file return false
            if (LoadSettings && !WizardInfo.CanConfirm)
            {
                var args = new MessageEventArgs();
                args.Title = Resources.InputSettingsFileError;
                args.LikelyCause = Resources.InputSettingsFileErrorLikelyCause;
                args.PotentialSolution = Resources.InputSettingsFileErrorPotentialSolution;
                return args;
            }
            return null;
        }


        #endregion

        #region protected/private methods

        /// <summary>
        /// We can initialize it if Data Source and Server connection is provided
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
        protected override bool CanInitializeWizardPage(WizardInfo info)
        {
            string title = null;
            string likelyCause = null;
            string potentialSolution = null;
            m_canShow = true;

            if (info.DataSourceParser == null)
            {
                m_canShow = false;
                title = Resources.SettingsFilePart_CantShowTitle;
                likelyCause = Resources.DataSourceNotEnteredErrorLikelyCause;
                potentialSolution = Resources.DataSourceNotEnteredErrorPotentialSolution;
            }
            else if (info.WorkItemGenerator == null || info.WorkItemGenerator.TfsNameToFieldMapping == null || info.WorkItemGenerator.TfsNameToFieldMapping.Count == 0)
            {
                m_canShow = false;
                title = Resources.SettingsFilePart_CantShowTitle;
                likelyCause = Resources.ServerNotSpecifiedErrorLikelyCause;
                potentialSolution = Resources.ServerNotSpecifiedErrorPotentialSolution;
            }
            if (!m_canShow)
            {
                Warning = title;
            }
            return m_canShow;
        }

        /// <summary>
        /// Initialization is required if there is any change in Data Source or Server Connection
        /// </summary>
        /// <param name="state"></param>
        /// <returns></returns>
        protected override bool IsInitializationRequired(WizardInfo state)
        {
            if (m_wizardInfo == null || m_prerequisite.IsDataSourceChanged() || m_prerequisite.IsServerConnectionChanged())
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Validates part state for valid setting File
        /// </summary>
        /// <returns></returns>
        public override bool ValidatePartState()
        {
            Warning = null;

            // If Setting File is Not Valid XML file return false
            if (LoadSettings && !WizardInfo.CanConfirm)
            {
                Warning = Resources.InputSettingsFileError;
                return false;
            }
            if (m_wizardInfo == null)
            {
                Warning = "Unexpected Action";
                return false;
            }
            return true;
        }

        /// <summary>
        /// Checks whether Confirm button can be enabled or not based on Input Mapping file is valid or not.
        /// </summary>
        private void EvaluateCanConfirm()
        {
            CanNext = ValidatePartState();
            WizardInfo.CanConfirm = false;
            if (LoadSettings &&
                String.CompareOrdinal(Path.GetExtension(SettingsFilePath), XMLExtension) == 0 &&
                File.Exists(SettingsFilePath))
            {
                WizardInfo.CanConfirm = true;
            }
            else
            {
                WizardInfo.CanConfirm = false;
            }
        }

        #endregion
    }
}
