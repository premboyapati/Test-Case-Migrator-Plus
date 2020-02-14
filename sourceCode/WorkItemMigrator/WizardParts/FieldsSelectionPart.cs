//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.IO;

    internal class FieldsSelectionPart : BaseWizardPart
    {
        #region Fields

        private string m_sourcePath;

        #endregion

        #region Constructor

        public FieldsSelectionPart()
        {
            Header = "Fields";
            Description = "Confirm system generated fields or create user defined fields from sample mht/word file. Use the preview link to verify the boundaries (field name highlighted in yellow and value highlighted in gray) for field name-value pairs.";
            CanBack = true;
            WizardPage = WizardPage.FieldsSelection;
            Fields = new ObservableCollection<SourceField>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// FIle Path of Sample MHT file
        /// </summary>
        public string SourcePath
        {
            get
            {
                return m_sourcePath;
            }
            set
            {
                m_sourcePath = value;
                NotifyPropertyChanged("SourcePath");
                if (File.Exists(value))
                {
                    LoadMHTParser(SourcePath);
                }
                else
                {
                    App.CallMethodInUISynchronizationContext(ClearFields, null);
                }
                CanNext = ValidatePartState();
            }
        }

        /// <summary>
        /// List of MHT Field Names
        /// </summary>
        public ObservableCollection<SourceField> Fields
        {
            get;
            private set;
        }

        #endregion

        #region Public methods

        public override void Reset()
        {
            using (new AutoWaitCursor())
            {
                m_sourcePath = m_wizardInfo.DataSourceParser.StorageInfo.Source;
                App.CallMethodInUISynchronizationContext(ClearFields, null);

                if (m_wizardInfo.DataSourceParser.StorageInfo.FieldNames.Count == 0)
                {
                    m_wizardInfo.DataSourceParser.ParseDataSourceFieldNames();
                }
                foreach (SourceField field in m_wizardInfo.DataSourceParser.StorageInfo.FieldNames)
                {
                    App.CallMethodInUISynchronizationContext(AddField, field);
                }
                NotifyPropertyChanged("SourcePath");
            }
        }

        public override bool UpdateWizardPart()
        {
            if (!ValidatePartState())
            {
                return false;
            }

            if (!IsUpdationRequired())
            {
                return true;
            }

            m_wizardInfo.DataSourceParser.StorageInfo.FieldNames.Clear();
            foreach (var field in Fields)
            {
                m_wizardInfo.DataSourceParser.StorageInfo.FieldNames.Add(field);
            }
            m_wizardInfo.SampleFileForFields = SourcePath;
            return true;
        }

        public bool AddFieldName(string fieldName)
        {
            if (!string.IsNullOrEmpty(fieldName))
            {
                fieldName = fieldName.Trim();
                MHTStorageInfo info = m_wizardInfo.DataSourceParser.StorageInfo as MHTStorageInfo;

                foreach (var field in Fields)
                {
                    if (String.CompareOrdinal(field.FieldName, fieldName) == 0)
                    {
                        return false;
                    }
                }
                if (info.PossibleFieldNames.Contains(fieldName))
                {
                    App.CallMethodInUISynchronizationContext(AddField, new SourceField(fieldName, true));
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Open MHT file with field name-value pairs in different highlighting
        /// </summary>
        public void Preview()
        {
            Warning = null;
            if (IsMHTFile(SourcePath))
            {
                var fields = new List<string>();
                foreach (var field in Fields)
                {
                    fields.Add(field.FieldName);
                }
                try
                {
                    MHTParser.Preview(SourcePath, fields);
                }
                catch (WorkItemMigratorException ex)
                {
                    Warning = ex.Args.Title;
                }
            }
            else
            {
                Warning = "Invalid MHT file";
            }
        }

        #endregion

        #region Protected methods

        private bool IsMHTFile(string SampleMHTFilePath)
        {
            string extension = Path.GetExtension(SampleMHTFilePath);
            if (String.CompareOrdinal(extension, ".mht") == 0 ||
                String.CompareOrdinal(extension, ".mhtml") == 0 ||
                String.CompareOrdinal(extension, ".doc") == 0 ||
                String.CompareOrdinal(extension, ".docx") == 0)
            {
                return true;
            }
            return false;
        }


        public override bool ValidatePartState()
        {
            Warning = null;
            if (Fields.Count == 0)
            {
                Warning = "No field name is specified";
                return false;
            }
            return true;
        }

        protected override bool CanInitializeWizardPage(WizardInfo info)
        {
            string title = null;
            string likelyCause = null;
            string potentialSolution = null;
            m_canShow = true;

            // If Data Source is not initialized then raise error
            if (info.DataSourceParser == null)
            {
                m_canShow = false;
                title = Resources.ColumnMappingPart_CantShowTitle;
                likelyCause = Resources.DataSourceNotEnteredErrorLikelyCause;
                potentialSolution = Resources.DataSourceNotEnteredErrorPotentialSolution;
            }
            // else if Workitem Fields have not been populated by TFS Server 
            else if (info.WorkItemGenerator == null || info.WorkItemGenerator.TfsNameToFieldMapping == null || info.WorkItemGenerator.TfsNameToFieldMapping.Count == 0)
            {
                m_canShow = false;
                title = Resources.ColumnMappingPart_CantShowTitle;
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
        /// Initialization is required if this wizard is not initialized or Data Source is updated or 
        /// TFS Connection is Updated or Mapping File selection is updated
        /// </summary>
        /// <param name="state"></param>
        /// <returns></returns>
        protected override bool IsInitializationRequired(WizardInfo state)
        {
            if (m_wizardInfo == null ||
                m_prerequisite.IsDataSourceChanged() ||
                m_prerequisite.IsServerConnectionChanged() ||
                m_prerequisite.IsSettingsFilePathChanged())
            {
                return true;
            }
            return false;
        }

        #endregion

        #region Private methods

        private void ClearFields(object value)
        {
            Fields.Clear();
        }

        private void AddField(object value)
        {
            Fields.Add(value as SourceField);
        }

        private void LoadMHTParser(string source)
        {
            using (new AutoWaitCursor())
            {
                try
                {
                    MHTStorageInfo info = m_wizardInfo.DataSourceParser.StorageInfo as MHTStorageInfo;

                    m_wizardInfo.DataSourceParser.Dispose();

                    App.CallMethodInUISynchronizationContext(ClearFields, null);

                    MHTStorageInfo newInfo = new MHTStorageInfo(source);
                    newInfo.IsFirstLineTitle = info.IsFirstLineTitle;
                    newInfo.IsFileNameTitle = info.IsFileNameTitle;

                    m_wizardInfo.DataSourceParser = new MHTParser(newInfo);
                    m_wizardInfo.DataSourceParser.ParseDataSourceFieldNames();
                    foreach (SourceField field in m_wizardInfo.DataSourceParser.StorageInfo.FieldNames)
                    {
                        App.CallMethodInUISynchronizationContext(AddField, field);
                    }
                }
                catch (ArgumentException)
                { }
            }
        }

        #endregion
    }

}
