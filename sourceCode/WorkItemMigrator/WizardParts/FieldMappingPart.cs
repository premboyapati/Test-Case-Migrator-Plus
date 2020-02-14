//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.IO;
    using System.Linq;

    /// <summary>
    /// Wizard Part Responsible for Field mappings
    /// </summary>
    internal class FieldMappingPart : BaseWizardPart
    {
        #region Fields

        // Prerequisite
        private bool m_isFileNameTitle;
        private bool m_isFirstLineTitle;
        private string m_testSuiteField;
        private List<string> m_mandatoryFields;

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        public FieldMappingPart()
        {
            Header = Resources.MapColumns_Header;
            Description = Resources.MapColumns_Description;
            CanBack = true;
            WizardPage = WizardPage.FieldMapping;
            CanShow = false;
            FieldMappingRows = new ObservableCollection<FieldMappingRow>();
            TestSuiteAvailableFields = new ObservableCollection<string>();
        }

        #endregion

        #region Properties

        public string TestSuiteField
        {
            get
            {
                return m_testSuiteField;
            }
            set
            {
                m_testSuiteField = value;
                NotifyPropertyChanged("TestSuiteField");
            }
        }

        public bool IsTestSuiteVisible
        {
            get
            {
                return WizardInfo != null &&
                       WizardInfo.DataSourceType == DataSourceType.Excel &&
                       WizardInfo.WorkItemGenerator != null &&
                       WizardInfo.WorkItemGenerator.WorkItemCategory == WorkItemGenerator.TestCaseCategory;
            }
        }


        public ObservableCollection<string> TestSuiteAvailableFields
        {
            get;
            private set;
        }


        /// <summary>
        /// List of Field Mapping Rows
        /// </summary>
        public ObservableCollection<FieldMappingRow> FieldMappingRows
        {
            get;
            private set;
        }

        /// <summary>
        /// MHT Flow: Is FIelst line of MHT title?
        /// </summary>
        public bool IsFirstLineTitle
        {
            get
            {
                return m_isFirstLineTitle;
            }
            set
            {
                m_isFirstLineTitle = value;
                if (m_isFirstLineTitle)
                {
                    m_isFileNameTitle = false;
                }
                UpdateTitleFieldInFieldMappingRows();
                NotifyPropertyChanged("IsFirstLineTitle");
                NotifyPropertyChanged("IsFileNameTitle");
                CanNext = ValidatePartState();
            }
        }

        /// <summary>
        /// MHT Flow: Is MHT File Name Title?
        /// </summary>
        public bool IsFileNameTitle
        {
            get
            {
                return m_isFileNameTitle;
            }
            set
            {
                m_isFileNameTitle = value;
                if (m_isFileNameTitle)
                {
                    m_isFirstLineTitle = false;
                }
                UpdateTitleFieldInFieldMappingRows();
                NotifyPropertyChanged("IsFirstLineTitle");
                NotifyPropertyChanged("IsFileNameTitle");
                CanNext = ValidatePartState();
            }
        }

        #endregion

        #region public methods

        /// <summary>
        /// Resets/ Updates the Field Mapping Wizard part
        /// </summary>
        public override void Reset()
        {
            m_mandatoryFields = new List<string>();
            foreach (var field in WizardInfo.WorkItemGenerator.TfsNameToFieldMapping.Values)
            {
                if (field.IsMandatory)
                {
                    m_mandatoryFields.Add(field.TfsName);
                }
            }

            IList<SourceField> fieldNames = m_wizardInfo.DataSourceParser.StorageInfo.FieldNames;

            // if Data Source is updated then clear all mapping rows and initialize them 
            if (m_prerequisite.IsDataSourceChanged())
            {
                IsFirstLineTitle = false;
                IsFileNameTitle = false;
                App.CallMethodInUISynchronizationContext(ClearFieldMappingRows, null);
                InitializeFieldMappingRows(m_wizardInfo.WorkItemGenerator.TfsNameToFieldMapping);
            }

            // if headers collection is not null
            if (fieldNames != null)
            {
                // then iterate throgh each header
                foreach (SourceField field in fieldNames)
                {
                    // getting the Work item field
                    IWorkItemField wiField = null;
                    if (m_wizardInfo.Migrator.SourceNameToFieldMapping != null &&
                        m_wizardInfo.Migrator.SourceNameToFieldMapping.ContainsKey(field.FieldName))
                    {
                        wiField = m_wizardInfo.Migrator.SourceNameToFieldMapping[field.FieldName];
                    }

                    // Getting Field Mapping Row having data source field equal to header
                    FieldMappingRow row = FieldMappingRows.FirstOrDefault((r) =>
                    {
                        return String.CompareOrdinal(r.DataSourceField, field.FieldName) == 0;
                    });

                    if (row != null)
                    {
                        // If TFS is Updated then Update the allowed values of WorkitemField in Mapping Row
                        if (m_prerequisite.IsServerConnectionChanged())
                        {
                            row.ResetAvailableFields(m_wizardInfo.WorkItemGenerator.TfsNameToFieldMapping);
                        }

                        // Setting the Workitemfield in colum mapping row
                        row.SetWorkItemField(wiField);
                    }

                    if (m_wizardInfo.DataSourceType == DataSourceType.MHT)
                    {
                        MHTStorageInfo info = m_wizardInfo.DataSourceParser.StorageInfo as MHTStorageInfo;
                        IsFileNameTitle = info.IsFileNameTitle;
                        IsFirstLineTitle = info.IsFirstLineTitle;
                    }
                }
            }

            ResetTestSuiteFields();
            NotifyPropertyChanged("IsTestSuiteVisible");
        }

        /// <summary>
        /// Updates the Wizard part with current Configuration
        /// </summary>
        /// <returns>Is updation successful?</returns>
        public override bool UpdateWizardPart()
        {
            m_wizardInfo.DataSourceParser.StorageInfo.TestSuiteFieldName = null;

            if (String.CompareOrdinal(TestSuiteField, Resources.SelectPlaceholder) != 0)
            {
                m_wizardInfo.DataSourceParser.StorageInfo.TestSuiteFieldName = TestSuiteField;
                m_wizardInfo.RelationshipsInfo.TestSuiteField = TestSuiteField;
            }

            if (!IsUpdationRequired())
            {
                return true;
            }
            IDictionary<string, IWorkItemField> sourceNameToFieldMapping = new Dictionary<string, IWorkItemField>();

            foreach (IWorkItemField field in m_wizardInfo.WorkItemGenerator.TfsNameToFieldMapping.Values)
            {
                field.SourceName = string.Empty;
            }

            // Filling the Field Mapping
            foreach (FieldMappingRow row in FieldMappingRows)
            {
                if (row.WIField != null)
                {
                    IWorkItemField wiField = m_wizardInfo.WorkItemGenerator.TfsNameToFieldMapping[row.WIField.TfsName];
                    wiField.SourceName = row.DataSourceField;
                    sourceNameToFieldMapping[row.DataSourceField] = wiField;
                    if (wiField.IsStepsField && m_wizardInfo.DataSourceType == DataSourceType.MHT)
                    {
                        MHTStorageInfo info = m_wizardInfo.DataSourceParser.StorageInfo as MHTStorageInfo;
                        info.StepsField = row.DataSourceField;
                    }
                }
            }

            foreach (IWorkItemField field in m_wizardInfo.WorkItemGenerator.TfsNameToFieldMapping.Values)
            {
                if (m_wizardInfo.DataSourceType == DataSourceType.MHT && field.IsTitleField)
                {
                    MHTStorageInfo info = m_wizardInfo.DataSourceParser.StorageInfo as MHTStorageInfo;
                    info.IsFirstLineTitle = IsFirstLineTitle;
                    info.IsFileNameTitle = IsFileNameTitle;
                    if (IsFirstLineTitle || IsFileNameTitle)
                    {
                        field.SourceName = MHTParser.TestTitleDefaultTag;
                        sourceNameToFieldMapping.Add(MHTParser.TestTitleDefaultTag, field);
                    }
                    else
                    {
                        info.TitleField = field.SourceName;
                    }
                }
            }

            // finding out whether all mandatory fields are mapped or not
            foreach (IWorkItemField field in m_wizardInfo.WorkItemGenerator.TfsNameToFieldMapping.Values)
            {
                if (field.IsMandatory && !sourceNameToFieldMapping.ContainsKey(field.SourceName))
                {
                    Warning = Resources.MandatoryFieldsNotMappedErrorTitle;

                    return false;
                }
            }

            m_wizardInfo.Migrator.SourceNameToFieldMapping = sourceNameToFieldMapping;
            m_wizardInfo.DataSourceParser.FieldNameToFields = sourceNameToFieldMapping;
            m_wizardInfo.WorkItemGenerator.SourceNameToFieldMapping = sourceNameToFieldMapping;

            if (m_wizardInfo.Reporter != null)
            {
                m_wizardInfo.Reporter.Dispose();
            }

            string reportDirectory = Path.Combine(Path.GetDirectoryName(m_wizardInfo.DataSourceParser.StorageInfo.Source),
                                                  "Report" + DateTime.Now.ToString("g", System.Globalization.CultureInfo.CurrentCulture).Replace(":", "_").Replace(" ", "_").Replace("/", "_"));
            switch (m_wizardInfo.DataSourceType)
            {
                case DataSourceType.Excel:
                    m_wizardInfo.Reporter = new ExcelReporter(m_wizardInfo);
                    string fileNameWithoutExtension = "Report";
                    string fileExtension = Path.GetExtension(m_wizardInfo.DataSourceParser.StorageInfo.Source);
                    m_wizardInfo.Reporter.ReportFile = Path.Combine(reportDirectory, fileNameWithoutExtension + fileExtension);
                    break;

                case DataSourceType.MHT:
                    m_wizardInfo.Reporter = new XMLReporter(m_wizardInfo);
                    string fileName = "Report.xml"; ;
                    m_wizardInfo.Reporter.ReportFile = Path.Combine(reportDirectory, fileName);
                    break;

                default:
                    throw new InvalidEnumArgumentException("Invalid data source type");
            }

            int count = 0;
            m_wizardInfo.InitializeProgressView();
            if (m_wizardInfo.ProgressPart != null)
            {
                m_wizardInfo.ProgressPart.Header = "Parsing...";
                m_wizardInfo.ProgressPart.Header = "Initializing Parsing...";
            }

            IList<ISourceWorkItem> sourceWorkItems = new List<ISourceWorkItem>();
            IList<ISourceWorkItem> rawSourceWorkItems = new List<ISourceWorkItem>();

            // Parse MHT DataSource Files
            if (m_wizardInfo.DataSourceType == DataSourceType.MHT)
            {
                MHTStorageInfo sampleInfo = m_wizardInfo.DataSourceParser.StorageInfo as MHTStorageInfo;

                IList<IDataStorageInfo> storageInfos = m_wizardInfo.DataStorageInfos;

                try
                {
                    foreach (IDataStorageInfo storageInfo in storageInfos)
                    {
                        if (m_wizardInfo.ProgressPart == null)
                        {
                            break;
                        }
                        MHTStorageInfo info = storageInfo as MHTStorageInfo;
                        info.IsFirstLineTitle = sampleInfo.IsFirstLineTitle;
                        info.IsFileNameTitle = sampleInfo.IsFileNameTitle;
                        info.TitleField = sampleInfo.TitleField;
                        info.StepsField = sampleInfo.StepsField;
                        IDataSourceParser parser = new MHTParser(info);

                        parser.ParseDataSourceFieldNames();
                        info.FieldNames = sampleInfo.FieldNames;
                        parser.FieldNameToFields = sourceNameToFieldMapping;

                        while (parser.GetNextWorkItem() != null)
                        {
                            count++;
                            if (m_wizardInfo.ProgressPart != null)
                            {
                                m_wizardInfo.ProgressPart.Text = "Parsing " + count + " of " + m_wizardInfo.DataStorageInfos.Count + ":\n" + info.Source;
                                m_wizardInfo.ProgressPart.ProgressValue = (count * 100) / m_wizardInfo.DataStorageInfos.Count;
                            }
                        }
                        for (int i = 0; i < parser.ParsedSourceWorkItems.Count; i++)
                        {
                            sourceWorkItems.Add(parser.ParsedSourceWorkItems[i]);
                            rawSourceWorkItems.Add(parser.RawSourceWorkItems[i]);
                        }
                        parser.Dispose();
                    }
                }
                catch (WorkItemMigratorException te)
                {
                    Warning = te.Args.Title;
                    m_wizardInfo.ProgressPart = null;
                    return false;
                }
            }
            else if (m_wizardInfo.DataSourceType == DataSourceType.Excel)
            {
                var excelInfo = m_wizardInfo.DataSourceParser.StorageInfo as ExcelStorageInfo;
                m_wizardInfo.DataSourceParser.Dispose();
                m_wizardInfo.DataSourceParser = new ExcelParser(excelInfo);
                m_wizardInfo.DataSourceParser.ParseDataSourceFieldNames();
                m_wizardInfo.DataSourceParser.FieldNameToFields = sourceNameToFieldMapping;
                var parser = m_wizardInfo.DataSourceParser;
                while (parser.GetNextWorkItem() != null)
                {
                    if (m_wizardInfo.ProgressPart == null)
                    {
                        break;
                    }
                    count++;
                    m_wizardInfo.ProgressPart.ProgressValue = excelInfo.ProgressPercentage;
                    m_wizardInfo.ProgressPart.Text = "Parsing work item# " + count;
                }
                for (int i = 0; i < parser.ParsedSourceWorkItems.Count; i++)
                {
                    sourceWorkItems.Add(parser.ParsedSourceWorkItems[i]);
                    rawSourceWorkItems.Add(parser.RawSourceWorkItems[i]);
                }
            }
            if (m_wizardInfo.ProgressPart != null)
            {
                m_wizardInfo.ProgressPart.Text = "Completing...";
                m_wizardInfo.Migrator.RawSourceWorkItems = rawSourceWorkItems;
                m_wizardInfo.Migrator.SourceWorkItems = sourceWorkItems;
            }
            else
            {
                Warning = "Parsing is cancelled";
                return false;
            }
            m_wizardInfo.ProgressPart = null;

            m_prerequisite.Save();
            return true;
        }

        #endregion

        #region protected/private methods

        /// <summary>
        /// Initializes the Field Mapping Rows
        /// </summary>
        /// <param name="workItemFieldByFieldName"></param>
        private void InitializeFieldMappingRows(IDictionary<string, IWorkItemField> workItemFieldByTFSFieldName)
        {
            IList<SourceField> fieldNames = m_wizardInfo.DataSourceParser.StorageInfo.FieldNames;
            bool isMultipleHistoryMappingMode = false;
            if (m_wizardInfo.DataSourceType == DataSourceType.MHT)
            {
                isMultipleHistoryMappingMode = true;
            }

            if (fieldNames != null)
            {
                // For each header create a field mapping row
                foreach (SourceField field in fieldNames)
                {
                    FieldMappingRow row = new FieldMappingRow(field.FieldName, null, workItemFieldByTFSFieldName, isMultipleHistoryMappingMode);
                    row.PropertyChanged += new PropertyChangedEventHandler(UpdateAvailableTFSFields);
                    App.CallMethodInUISynchronizationContext(AddFieldMappingRow, row);
                }
            }
        }

        private void AddFieldMappingRow(object row)
        {
            FieldMappingRows.Add(row as FieldMappingRow);
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
                m_prerequisite.IsSettingsFilePathChanged() ||
                m_prerequisite.AreFieldsChanged())
            {
                return true;
            }
            return false;
        }

        protected override bool IsUpdationRequired()
        {
            return m_wizardInfo.Migrator.SourceNameToFieldMapping == null ||
                m_wizardInfo.Migrator.SourceNameToFieldMapping.Count == 0 ||
                m_prerequisite.IsFieldMappingModified(FieldMappingRows, IsFirstLineTitle, IsFileNameTitle);
        }

        /// <summary>
        /// The Wizard part state will always be true
        /// </summary>
        /// <returns></returns>
        public override bool ValidatePartState()
        {
            Warning = null;
            if (m_mandatoryFields == null)
            {
                Warning = Resources.MandatoryFieldsNotMappedErrorTitle;
                return false;
            }
            int count = 0;
            foreach (var row in FieldMappingRows)
            {
                if (row.IsValidFieldMapping && m_mandatoryFields.Contains(row.WIField.TfsName))
                {
                    count++;
                }
            }
            IWorkItemField titleField = null;
            foreach (var field in WizardInfo.WorkItemGenerator.TfsNameToFieldMapping.Values)
            {
                if (field.IsTitleField)
                {
                    titleField = field;
                    break;
                }
            }
            if (titleField != null &&
                titleField.IsMandatory &&
                (IsFileNameTitle || IsFirstLineTitle))
            {
                count++;
            }

            if (m_mandatoryFields.Count != count)
            {
                Warning = Resources.MandatoryFieldsNotMappedErrorTitle;
                return false;
            }
            return true;
        }

        /// <summary>
        /// Updates the allowed values of each Field Mapping Row if Workitem Field Selection of any mapping row is changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateAvailableTFSFields(object sender, PropertyChangedEventArgs e)
        {
            if (String.CompareOrdinal(e.PropertyName, "TFSField") == 0)
            {
                FieldMappingRow effectedRow = sender as FieldMappingRow;
                foreach (FieldMappingRow row in FieldMappingRows)
                {
                    row.UpdateAvailableFields(effectedRow);
                }
                effectedRow.PreviousTFSField = effectedRow.TFSField;
                CanNext = ValidatePartState();
            }
        }

        /// <summary>
        /// Is initialization of this wizard part possible?
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
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
            // else if there are no headers
            else if (info.DataSourceParser.StorageInfo.FieldNames == null || info.DataSourceParser.StorageInfo.FieldNames.Count == 0)
            {
                m_canShow = false;
                title = Resources.ColumnMappingPart_CantShowTitle;
                likelyCause = Resources.DataSourceFieldNamesNotFoundErrorLikelyCause;
                potentialSolution = Resources.DataSourceFieldNamesNotFoundErrorPotentialSolution;
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

        private void UpdateTitleFieldInFieldMappingRows()
        {
            if (m_isFileNameTitle || m_isFirstLineTitle)
            {
                foreach (var row in FieldMappingRows)
                {
                    if (row.WIField != null && row.WIField.IsTitleField)
                    {
                        row.TFSField = null;
                        break;
                    }
                }
            }

            foreach (var row in FieldMappingRows)
            {
                if (m_isFileNameTitle || m_isFirstLineTitle)
                {
                    row.RemoveTitleField();
                }
                else
                {
                    row.AddTitleField();
                }
            }
        }

        private void ClearFieldMappingRows(object value)
        {
            FieldMappingRows.Clear();
        }

        private void ResetTestSuiteFields()
        {
            ClearTestSuiteAvailableFields();
            foreach (var field in m_wizardInfo.DataSourceParser.StorageInfo.FieldNames)
            {
                bool isFieldAvailable = true;
                foreach (var row in FieldMappingRows)
                {
                    if (row.WIField != null &&
                        String.CompareOrdinal(field.FieldName, row.DataSourceField) == 0)
                    {
                        isFieldAvailable = false;
                        break;
                    }
                }
                if (isFieldAvailable)
                {
                    AddTestSuiteAvailableField(field.FieldName);
                }
            }
            TestSuiteField = Resources.SelectPlaceholder;
            if (TestSuiteAvailableFields.Contains(m_wizardInfo.RelationshipsInfo.TestSuiteField))
            {
                TestSuiteField = m_wizardInfo.RelationshipsInfo.TestSuiteField;
            }
        }

        private void ClearTestSuiteAvailableFieldsInUIContext(object obj)
        {
            TestSuiteAvailableFields.Clear();
            TestSuiteAvailableFields.Add(Resources.SelectPlaceholder);
        }

        private void ClearTestSuiteAvailableFields()
        {
            App.CallMethodInUISynchronizationContext(ClearTestSuiteAvailableFieldsInUIContext, null);
        }

        private void AddTestSuiteAvailableField(string fieldName)
        {
            App.CallMethodInUISynchronizationContext(AddTestSuiteAvailableFieldInUIContext, fieldName);
        }

        private void AddTestSuiteAvailableFieldInUIContext(object obj)
        {
            string fieldName = obj as string;
            if (!string.IsNullOrEmpty(fieldName))
            {
                TestSuiteAvailableFields.Add(fieldName);
            }
        }

        private void RemoveTestSuiteAvailableField(string fieldName)
        {
            App.CallMethodInUISynchronizationContext(RemoveTestSuiteAvailableFieldInUIContext, fieldName);
        }

        private void RemoveTestSuiteAvailableFieldInUIContext(object obj)
        {
            string fieldName = obj as string;
            if (!string.IsNullOrEmpty(fieldName) && TestSuiteAvailableFields.Contains(fieldName))
            {
                TestSuiteAvailableFields.Remove(fieldName);
            }
        }

        #endregion
    }

    /// <summary>
    /// Class for Field Mapping Row's WorkItemField
    /// </summary>
    internal class DisplayWorkItemField : NotifyPropertyChange
    {
        #region Fields

        // Corresponding Workitem Field
        private IWorkItemField m_field;

        // Displayed name of this Field
        private string m_displayedName;

        #endregion

        #region Constsnts

        // Star Charcter to show Mandatory Field
        private const string StarCharacter = "*";

        // Plus Character to show Auto generated fields
        private const string PlusCharacter = "+";

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="wiField"></param>
        public DisplayWorkItemField(IWorkItemField wiField)
        {
            WIField = wiField;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Corresponding WorkitemField of thus DisplayedWorkitemField
        /// </summary>
        public IWorkItemField WIField
        {
            get
            {
                return m_field;
            }
            set
            {
                m_field = value;
                SetDisplayedName();
            }
        }

        /// <summary>
        /// Displayed Name
        /// </summary>
        public string DisplayedName
        {
            get
            {
                return m_displayedName;
            }
            set
            {
                m_displayedName = value;
                NotifyPropertyChanged("DisplayedName");
            }
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Set Dispplayed Name based on WorkitemField's value
        /// </summary>
        private void SetDisplayedName()
        {
            // if Workitem Field is null then set Display name as <Ignore>
            if (m_field == null)
            {
                DisplayedName = Resources.IgnoreLabel;
                return;
            }
            string postChar = string.Empty;

            // If this is a mandatory field then add suffix *
            if (m_field.IsMandatory)
            {
                postChar += StarCharacter;
            }
            // if this is auto generated field then add suffix +
            if (m_field.IsAutoGenerated)
            {
                postChar += PlusCharacter;
            }
            DisplayedName = m_field.TfsName + postChar;
        }

        #endregion

    }

    /// <summary>
    /// Field Mapping Row Class
    /// </summary>
    internal class FieldMappingRow : NotifyPropertyChange
    {
        #region Nested Definition

        private struct InsertField
        {
            public int index;
            public string tfsField;
        }

        #endregion

        #region Fields

        // Hash table of WorkItemFields by Field names
        private Dictionary<string, DisplayWorkItemField> m_workItemFieldByFieldName = new Dictionary<string, DisplayWorkItemField>();

        // Display WorkItemField
        private DisplayWorkItemField m_field;

        private bool m_isMultipleHistoryMappingMode;

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="dataSourceField"></param>
        /// <param name="wiField"></param>
        /// <param name="workItemFieldByFieldName"></param>
        public FieldMappingRow(string dataSourceField, IWorkItemField wiField, IDictionary<string, IWorkItemField> workItemFieldByFieldName, bool isMultipleHistoryMappingMode)
        {
            DataSourceField = dataSourceField;
            PreviousTFSField = Resources.IgnoreLabel;
            AvailableTFSFields = new ObservableCollection<string>();
            ResetAvailableFields(workItemFieldByFieldName);
            m_field = new DisplayWorkItemField(wiField);
            m_isMultipleHistoryMappingMode = isMultipleHistoryMappingMode;

            if (wiField == null && m_isMultipleHistoryMappingMode)
            {
                foreach (var field in m_workItemFieldByFieldName.Values)
                {
                    if (field.WIField.IsHtmlField)
                    {
                        TFSField = field.DisplayedName;
                    }
                }
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Data Source Field Name
        /// </summary>
        public string DataSourceField
        {
            get;
            set;
        }

        /// <summary>
        /// Display Name of WorkitemField
        /// </summary>
        public string TFSField
        {
            get
            {
                return m_field.DisplayedName;
            }
            set
            {
                SetTFSField(value);
                NotifyPropertyChanged("TFSField");
            }
        }

        /// <summary>
        /// Previous Display Name of Work item Field
        /// </summary>
        public string PreviousTFSField
        {
            get;
            set;
        }

        /// <summary>
        /// WorkItem Field
        /// </summary>
        public IWorkItemField WIField
        {
            get
            {
                return m_field.WIField;
            }
        }

        /// <summary>
        /// Is this a valid field mapping
        /// </summary>
        public bool IsValidFieldMapping
        {
            get
            {
                return m_field.WIField != null;
            }
        }

        /// <summary>
        /// List of name of available TFS Fields to be mapped
        /// </summary>
        public ObservableCollection<string> AvailableTFSFields
        {
            get;
            private set;
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Updates the available tfs fields based on the selection change in effectedRow
        /// </summary>
        /// <param name="effectedRow"></param>
        public void UpdateAvailableFields(FieldMappingRow effectedRow)
        {
            // if there is really a change in TFS Field selection
            if (String.CompareOrdinal(effectedRow.PreviousTFSField, effectedRow.TFSField) != 0)
            {
                // If the current mappping row is not the effected row
                if (this != effectedRow)
                {
                    string ignore = Resources.IgnoreLabel;

                    // If effected row become a valid Field Mapping Row then remove the corresponding 
                    // tfs field from the list of available fields
                    if (String.CompareOrdinal(effectedRow.PreviousTFSField, ignore) == 0)
                    {
                        RemoveAvailableTFSField(effectedRow.TFSField);
                    }
                    // else if the effected row was earlier mapped but now mapped to noting
                    // then Add the previous mapped tfs field to the list of allowed values
                    else if (String.CompareOrdinal(effectedRow.TFSField, ignore) == 0)
                    {
                        AddAvailableTFSField(effectedRow.PreviousTFSField);
                    }
                    // else it is only a change of mapped WorkItem field in the effected row. just updates the
                    // corresponding fields in the current list of allowed tfs fields
                    else
                    {
                        AddAvailableTFSField(effectedRow.PreviousTFSField);
                        RemoveAvailableTFSField(effectedRow.TFSField);
                    }
                }
            }
        }

        /// <summary>
        /// Resets the list of available fields
        /// </summary>
        /// <param name="workItemFieldByFieldName"></param>
        public void ResetAvailableFields(IDictionary<string, IWorkItemField> workItemFieldByFieldName)
        {
            using (new AutoWaitCursor())
            {
                m_workItemFieldByFieldName.Clear();
                App.CallMethodInUISynchronizationContext(ClearAvailableFields, null);
                AddAvailableTFSField(Resources.IgnoreLabel);
                foreach (IWorkItemField field in workItemFieldByFieldName.Values)
                {
                    DisplayWorkItemField displayedField = new DisplayWorkItemField(field);
                    m_workItemFieldByFieldName.Add(displayedField.DisplayedName, displayedField);
                    AddAvailableTFSField(displayedField.DisplayedName);
                }
            }
        }

        public void RemoveTitleField()
        {
            string titleField = null;
            foreach (string field in AvailableTFSFields)
            {
                if (m_workItemFieldByFieldName.ContainsKey(field) &&
                    m_workItemFieldByFieldName[field].WIField.IsTitleField)
                {
                    titleField = m_workItemFieldByFieldName[field].DisplayedName;
                    break;
                }
            }
            if (!string.IsNullOrEmpty(titleField))
            {
                RemoveAvailableTFSField(titleField);
            }
        }

        public void AddTitleField()
        {
            string titleField = null;
            foreach (var kvp in m_workItemFieldByFieldName)
            {
                if (kvp.Value.WIField.IsTitleField)
                {
                    titleField = kvp.Key;
                    break;
                }
            }
            if (!string.IsNullOrEmpty(titleField))
            {
                AddAvailableTFSField(titleField);
            }
        }

        /// <summary>
        /// Sets WorkitemField
        /// </summary>
        /// <param name="tfsField"></param>
        public void SetWorkItemField(IWorkItemField tfsField)
        {
            m_field.WIField = tfsField;
            NotifyPropertyChanged("TFSField");
            NotifyPropertyChanged("IsValidFieldMapping");
        }

        #endregion

        #region private methods

        /// <summary>
        /// Removes given fieldname from the list of avaible fields
        /// </summary>
        /// <param name="tfsField"></param>
        private void RemoveAvailableTFSField(string tfsField)
        {
            if (!string.IsNullOrEmpty(tfsField) &&
                String.CompareOrdinal(TFSField, tfsField) != 0)
            {
                if (!(m_isMultipleHistoryMappingMode && m_workItemFieldByFieldName[tfsField].WIField.IsHtmlField))
                {
                    App.CallMethodInUISynchronizationContext(RemoveAvailableField, tfsField);
                }
            }
        }

        /// <summary>
        /// Adds tfsField in list of available fields such that the list remains sorted
        /// </summary>
        /// <param name="tfsField"></param>
        private void AddAvailableTFSField(string tfsField)
        {
            if (AvailableTFSFields.Contains(tfsField))
            {
                return;
            }

            if (!string.IsNullOrEmpty(tfsField))
            {
                int index;
                if (String.CompareOrdinal(tfsField, Resources.IgnoreLabel) == 0)
                {
                    index = 0;
                }
                else
                {
                    int i = 0;
                    foreach (string field in AvailableTFSFields)
                    {
                        i++;
                        if (String.CompareOrdinal(field, Resources.IgnoreLabel) != 0 &&
                            string.CompareOrdinal(tfsField, field) < 0)
                        {
                            i--;
                            break;
                        }
                    }
                    index = i;
                }
                InsertField insertField = new InsertField();
                insertField.index = index;
                insertField.tfsField = tfsField;
                App.CallMethodInUISynchronizationContext(AddAvailableField, insertField);
            }
        }

        private void ClearAvailableFields(object value)
        {
            AvailableTFSFields.Clear();
        }

        private void RemoveAvailableField(object value)
        {
            AvailableTFSFields.Remove(value as string);
        }

        private void AddAvailableField(object value)
        {
            InsertField field = (InsertField)value;
            AvailableTFSFields.Insert(field.index, field.tfsField);
        }

        /// <summary>
        /// Sets corresponding Display TFS Field
        /// </summary>
        /// <param name="fieldName"></param>
        private void SetTFSField(string fieldName)
        {
            if (string.IsNullOrEmpty(fieldName) ||
                String.CompareOrdinal(fieldName, Resources.IgnoreLabel) == 0)
            {
                m_field.WIField = null;
            }
            else
            {
                m_field.WIField = m_workItemFieldByFieldName[fieldName].WIField;
            }
            NotifyPropertyChanged("TFSField");
            NotifyPropertyChanged("IsValidFieldMapping");
        }

        #endregion
    }
}
