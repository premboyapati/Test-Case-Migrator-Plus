//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;

    /// <summary>
    /// Wizard part to Manage Data Mapping
    /// </summary>
    internal class DataMappingPart : BaseWizardPart
    {
        #region Fields

        // Representation of all values found in the Data Source for Fields which has allowed values
        IDictionary<string, IList<string>> m_dataValuesByFieldName;

        private bool m_createAreaIterationPath;
        private bool m_isCreateAreaIterationPathVisible;

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        public DataMappingPart()
        {
            Header = Resources.MapValues_Header;
            Description = Resources.MapValues_Description;
            CanBack = true;
            WizardPage = WizardPage.DataMapping;
            DataMappingRows = new ObservableCollection<DataMappingRow>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Data Mapping Rows
        /// </summary>
        public ObservableCollection<DataMappingRow> DataMappingRows
        {
            get;
            private set;
        }

        public bool CreateAreaIterationPath
        {
            get
            {
                return m_createAreaIterationPath;
            }
            set
            {
                m_createAreaIterationPath = value;
                ChangeEnableStateofDataMappingRowsForAreaIterationPathField();
                NotifyPropertyChanged("CreateAreaIterationPath");
            }
        }

        public bool IsCreateAreaIterationPathVisible
        {
            get
            {
                return m_isCreateAreaIterationPathVisible;
            }
            set
            {
                m_isCreateAreaIterationPathVisible = value;
                NotifyPropertyChanged("IsCreateAreaIterationPathVisible");
            }
        }

        #endregion

        #region public methods

        /// <summary>
        /// Initializes the Data Mapping part
        /// </summary>
        /// <param name="info"></param>
        public override void Initialize(WizardInfo info)
        {
            CanNext = true;

            // if we can't Initialize the Data mapping part then just returns
            if (!CanInitializeWizardPage(info))
            {
                return;
            }

            // If initialization is required
            if (IsInitializationRequired(m_wizardInfo))
            {
                m_wizardInfo = info;

                if (m_prerequisite == null)
                {
                    m_prerequisite = new WizardPartPrerequisite(m_wizardInfo);
                }
                try
                {
                    Reset();
                }
                catch (WorkItemMigratorException ex)
                {
                    Warning = ex.Args.Title;
                    CanShow = false;
                    return;
                }
            }
            else
            {
                // else Just Updates the Data Mapping for removed Column mappings
                UpdateDatamappingRowsOnRemovedColumnMappings();

                // Update Allowed Values of Data Mappings if only WorkItem fields are changed in Field Mappings
                foreach (DataMappingRow row in DataMappingRows)
                {
                    if (row.DataSourceField != null && m_wizardInfo.Migrator.SourceNameToFieldMapping.ContainsKey(row.DataSourceField))
                    {
                        row.UpdateAllowedValues(m_wizardInfo.Migrator.SourceNameToFieldMapping[row.DataSourceField]);
                    }
                }
            }

            m_prerequisite.Save();

            // Send Notifications
            FireStateNotifications();
        }

        /// <summary>
        /// Adds Editable Data Mapping Row at the end of List of rows
        /// </summary>
        public void AddEditableDataMappingRow()
        {
            DataMappingRows.Insert(DataMappingRows.Count - 1, new DataMappingRow(m_wizardInfo));
        }

        /// <summary>
        /// Resets the Data mapping Wizard Part 
        /// </summary>
        public override void Reset()
        {
            //Clears Data Mapping Rows

            App.CallMethodInUISynchronizationContext(ClearDataMappingRows, null);

            // Add Data Mapping Row's Header
            App.CallMethodInUISynchronizationContext(AddDataMappingRow, new HeaderDataMappingRow());

            // Parsing the DataSource. It will return list of allowed values

            m_dataValuesByFieldName = new Dictionary<string, IList<string>>();
            IsCreateAreaIterationPathVisible = false;
            foreach (var kvp in m_wizardInfo.Migrator.SourceNameToFieldMapping)
            {
                if (kvp.Value.HasAllowedValues && m_wizardInfo.Migrator.FieldToUniqueValues.ContainsKey(kvp.Key))
                {
                    m_dataValuesByFieldName.Add(kvp.Key, m_wizardInfo.Migrator.FieldToUniqueValues[kvp.Key]);
                }
                if (kvp.Value.IsAreaPath || kvp.Value.IsIterationPath)
                {
                    IsCreateAreaIterationPathVisible = true;
                }
            }


            // Filling the data mapping rows
            foreach (KeyValuePair<string, IList<string>> dataValuesbyFieldName in m_dataValuesByFieldName)
            {
                string dataSourceFieldName = dataValuesbyFieldName.Key;
                IWorkItemField workItemField = m_wizardInfo.Migrator.SourceNameToFieldMapping[dataSourceFieldName];
                if (workItemField.HasAllowedValues)
                {
                    foreach (string dataSourceValue in dataValuesbyFieldName.Value)
                    {
                        string newValue = string.Empty;
                        if (workItemField.ValueMapping.ContainsKey(dataSourceValue))
                        {
                            newValue = workItemField.ValueMapping[dataSourceValue];
                        }

                        DataMappingRow dataMappingRow = new DataMappingRow(m_wizardInfo, dataSourceFieldName, dataSourceValue, newValue);
                        App.CallMethodInUISynchronizationContext(AddDataMappingRow, dataMappingRow);
                    }
                }
            }
            App.CallMethodInUISynchronizationContext(AddDataMappingRow, new BlankDataMappingRow());
            CreateAreaIterationPath = m_wizardInfo.WorkItemGenerator.CreateAreaIterationPath;
        }

        /// <summary>
        /// Updates the Wizard Info with the Data Mappings selected in this Wizard Part
        /// </summary>
        /// <returns></returns>
        public override bool UpdateWizardPart()
        {
            foreach (DataMappingRow row in DataMappingRows)
            {
                if (row is HeaderDataMappingRow || row is BlankDataMappingRow || row.DataSourceField == null)
                {
                    continue;
                }
                IWorkItemField workItemField = m_wizardInfo.Migrator.SourceNameToFieldMapping[row.DataSourceField];
                if (!string.IsNullOrEmpty(row.DataSourceValue) && !string.IsNullOrEmpty(row.NewValue))
                {
                    if (workItemField.ValueMapping.ContainsKey(row.DataSourceValue))
                    {
                        workItemField.ValueMapping[row.DataSourceValue] = row.NewValue;
                    }
                    else
                    {
                        workItemField.ValueMapping.Add(row.DataSourceValue, row.NewValue);
                    }
                }
            }
            m_wizardInfo.WorkItemGenerator.CreateAreaIterationPath = CreateAreaIterationPath;
            return true;
        }

        #endregion

        #region protected/private methods

        private void ChangeEnableStateofDataMappingRowsForAreaIterationPathField()
        {
            foreach (var row in DataMappingRows)
            {
                row.IsEnabled = true;
                if (row.WorkItemField != null &&
                    (row.WorkItemField.IsAreaPath || row.WorkItemField.IsIterationPath))
                {
                    row.IsEnabled = !CreateAreaIterationPath;
                }
            }
        }

        /// <summary>
        /// Removes Data mapping Rows corresponding to those Data Fields which are removed in Field Mapping
        /// </summary>
        private void UpdateDatamappingRowsOnRemovedColumnMappings()
        {
            List<string> removedColumns = new List<string>();
            foreach (KeyValuePair<string, IList<string>> kvp in m_dataValuesByFieldName)
            {
                if (!m_wizardInfo.Migrator.SourceNameToFieldMapping.ContainsKey(kvp.Key))
                {
                    removedColumns.Add(kvp.Key);
                }
            }
            foreach (string removedColumn in removedColumns)
            {
                for (int i = 0; i < DataMappingRows.Count; i++)
                {
                    if (String.CompareOrdinal(DataMappingRows[i].DataSourceField, removedColumn) == 0)
                    {
                        DataMappingRows.Remove(DataMappingRows[i]);
                        i--;
                    }
                }
            }
        }

        /// <summary>
        /// Is Initialization of Data Mapping Wizard Part required?
        /// </summary>
        /// <param name="state"></param>
        /// <returns></returns>
        protected override bool IsInitializationRequired(WizardInfo state)
        {
            if (m_wizardInfo == null || m_prerequisite.IsDataSourceChanged() || m_prerequisite.IsFieldMappingChanged())
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// The Wizard Part's State will always be correct so just returns true
        /// </summary>
        /// <returns></returns>
        public override bool ValidatePartState()
        {
            Warning = null;
            return true;
        }

        /// <summary>
        /// Can Controller Initialize the Data Mapping Wizard part and can make it as active Wizard part
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
        protected override bool CanInitializeWizardPage(WizardInfo info)
        {
            m_canShow = true;

            if (info.Migrator.SourceNameToFieldMapping == null || info.Migrator.FieldToUniqueValues == null)
            {
                m_canShow = false;
                Warning = Resources.MandatoryFieldsNotMappedErrorTitle;
                m_canShow = false;
            }
            else
            {
                foreach (IWorkItemField field in info.WorkItemGenerator.TfsNameToFieldMapping.Values)
                {
                    if (field.IsMandatory && string.IsNullOrEmpty(field.SourceName))
                    {
                        Warning = Resources.MandatoryFieldsNotMappedErrorTitle;
                        m_canShow = false;
                        break;
                    }
                }
            }
            return m_canShow;
        }

        private void ClearDataMappingRows(object obj)
        {
            DataMappingRows.Clear();
        }

        private void AddDataMappingRow(object row)
        {
            DataMappingRows.Add(row as DataMappingRow);
        }

        private void RemoveDataMappingRow(object row)
        {
            DataMappingRows.Remove(row as DataMappingRow);
        }

        #endregion
    }

    /// <summary>
    /// Class that represent one Data Mapping Row.
    /// </summary>
    internal class DataMappingRow : NotifyPropertyChange
    {
        #region Fields

        // Wizard Info: needed for initialization and updation of the Mapping Row
        private WizardInfo m_wizardInfo;

        // The Mapped Workitem Field for the selected Data Source Field
        private IWorkItemField m_wiField;

        // member variables needed for Data Binding
        private string m_dataSourceField;
        private string m_dataSourceValue;
        private string m_newValue;
        private bool m_isEditable;
        private ObservableCollection<string> m_allowedValues;
        private bool m_isEnabled;

        #endregion

        #region Constructor

        /// <summary>
        /// Basic Constructor for Derived Data Mapping Rows
        /// </summary>
        protected DataMappingRow()
        { }

        /// <summary>
        /// Constructor - Creates Data Mapping row in Non-Editable Mode
        /// </summary>
        /// <param name="info"></param>
        /// <param name="field"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        public DataMappingRow(WizardInfo info, string field, string oldValue, string newValue)
        {
            m_wizardInfo = info;
            AllowedNewValues = new ObservableCollection<string>();
            DataSourceField = field;
            DataSourceValue = oldValue;
            FillAllowedValues();
            NewValue = newValue;
            IsEnabled = true;
        }

        /// <summary>
        /// Constructor - Creates Data Mapping row in editable mode
        /// </summary>
        /// <param name="info"></param>
        public DataMappingRow(WizardInfo info)
        {
            m_wizardInfo = info;
            AllowedNewValues = new ObservableCollection<string>();
            InitializeDataSourceFields();
            IsEditable = true;
            IsEnabled = true;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Selected Data Source Field
        /// </summary>
        public string DataSourceField
        {
            get
            {
                return m_dataSourceField;
            }
            set
            {
                m_dataSourceField = value;

                // Updates the list of corresponding allowed values at TFS for this new Data Source Field
                if (m_wizardInfo != null && m_wizardInfo.Migrator.SourceNameToFieldMapping.ContainsKey(m_dataSourceField))
                {
                    UpdateAllowedValues(m_wizardInfo.Migrator.SourceNameToFieldMapping[m_dataSourceField]);
                }
                NotifyPropertyChanged("DataSourceField");
                NotifyPropertyChanged("NewValue");
            }
        }

        /// <summary>
        /// List of Data Source Fields which are mapped in Field Mapping
        /// </summary>
        public List<string> DataSourceFields
        {
            get;
            private set;
        }

        /// <summary>
        /// The value at data Source which is going to be mapped
        /// </summary>
        public string DataSourceValue
        {
            get
            {
                return m_dataSourceValue;
            }
            set
            {
                m_dataSourceValue = value;
                NotifyPropertyChanged("DataSourceValue");
            }
        }

        /// <summary>
        /// The new Value of the Data mapping. This value will replace the Data Source Value during migration
        /// </summary>
        public virtual string NewValue
        {
            get
            {
                return m_newValue;
            }
            set
            {
                // Only update the new value if it is not null and present in list of allowed values
                if (string.IsNullOrEmpty(value) || AllowedNewValues == null || !AllowedNewValues.Contains(value))
                {
                    value = Resources.IgnoreLabel;
                }
                m_newValue = value;
                NotifyPropertyChanged("NewValue");
            }
        }

        /// <summary>
        /// List of allowed values for new Value
        /// </summary>
        public ObservableCollection<string> AllowedNewValues
        {
            get
            {
                return m_allowedValues;
            }
            set
            {
                m_allowedValues = value;
                NotifyPropertyChanged("AllowedNewValues");
            }
        }

        public IWorkItemField WorkItemField
        {
            get
            {
                return m_wiField;
            }
        }

        /// <summary>
        /// Is this Data Mapping Row in Editable Mode
        /// </summary>
        public bool IsEditable
        {
            get
            {
                return m_isEditable;
            }
            set
            {
                m_isEditable = value;
                NotifyPropertyChanged("IsEditable");
            }
        }

        public bool IsEnabled
        {
            get
            {
                return m_isEnabled;
            }
            set
            {
                m_isEnabled = value;
                NotifyPropertyChanged("IsEnabled");
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Update Allowed Values of Data Mapping Row if TFS connection is modified
        /// </summary>
        /// <param name="wiField"></param>
        public void UpdateAllowedValues(IWorkItemField wiField)
        {
            if (wiField != m_wiField)
            {
                m_wiField = wiField;
                FillAllowedValues();
                NewValue = null;
            }
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Initializes The List of Data Source Fields
        /// </summary>
        private void InitializeDataSourceFields()
        {
            DataSourceFields = new List<string>();
            foreach (KeyValuePair<string, IWorkItemField> kvp in m_wizardInfo.Migrator.SourceNameToFieldMapping)
            {
                if (kvp.Value.HasAllowedValues)
                {
                    DataSourceFields.Add(kvp.Key);
                }
            }
            if (DataSourceFields.Count > 0)
            {
                DataSourceField = DataSourceFields[0];
            }
        }

        private void FillAllowedValues()
        {
            AllowedNewValues.Clear();
            AllowedNewValues.Add(Resources.IgnoreLabel);

            if (m_wiField.HasAllowedValues)
            {
                foreach (string value in m_wiField.AllowedValues)
                {
                    AllowedNewValues.Add(value);
                }
            }
        }

        #endregion

    }

    /// <summary>
    /// Dummy Data Mapping Row: needed by view to create "Add" Data Mapping Functionality
    /// </summary>
    internal class BlankDataMappingRow : DataMappingRow { }


    /// <summary>
    /// Header of Data Mapping Rows
    /// </summary>
    internal class HeaderDataMappingRow : DataMappingRow
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public HeaderDataMappingRow()
        {
            DataSourceField = Resources.DataMapping_DataSourceFieldHeader;
            DataSourceValue = Resources.DataMapping_DataSourceValueHeader;
            NewValue = Resources.Datamapping_NewValueHeader;
        }

        public override string NewValue
        {
            get;
            set;
        }
    }
}
