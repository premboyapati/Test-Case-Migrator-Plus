//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.IO;
using System.Threading;

    /// <summary>
    /// Wizard Part responsible for getting Data Source Information.
    /// </summary>
    internal class SelectDataSourcePart : BaseWizardPart, IDisposable
    {
        #region Fields

        private string m_excelFilePath;
        private string m_selectedExcelSheet;
        private string m_excelHeaderRow;
        private DataSourceType m_dataSourceType;
        private string m_selectedDataSource;
        private ObservableCollection<string> m_fields;

        private bool m_isMHTFolder;
        private string m_mhtFolderPath;
        private string m_listOfMHTsFilePath;
        private Dictionary<string, DataSourceType> m_dataSourceNameToTypeMapping;
        private int m_mhtCount;
        private bool m_isLoadingFiles;
        protected BackgroundWorker m_worker;


        #endregion

        #region Constants

        // Default Excel Row which has Field Names
        private const string DefaultExcelHeaderRow = "1";

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor to initialize the wizard Part properties
        /// </summary>
        public SelectDataSourcePart()
        {
            Header = Resources.SelectDataSource_Header;
            Description = Resources.SelectDataSource_Description;
            CanBack = true;
            WizardPage = WizardPage.SelectDataSource;
            ExcelSheets = new ObservableCollection<string>();
            DataSources = new List<string>();
            m_fields = new ObservableCollection<string>();
            IsMHTFolder = true;
            CanShow = true;
            ExcelHeaderRow = DefaultExcelHeaderRow;

            m_dataSourceNameToTypeMapping = new Dictionary<string, DataSourceType>();
            m_dataSourceNameToTypeMapping.Add("Excel Workbook", DataSourceType.Excel);
            m_dataSourceNameToTypeMapping.Add("VS 2005/2008 Manual Test Format(MHT/Word)", DataSourceType.MHT);
            foreach (var kvp in m_dataSourceNameToTypeMapping)
            {
                DataSources.Add(kvp.Key);
                if (kvp.Value == DataSourceType.Excel)
                {
                    SelectedDataSource = kvp.Key;
                }
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Bool Check needed to enable/disable Preview button for previewing Data Source
        /// </summary>
        public bool IsPreviewEnabled
        {
            get
            {
                switch (DataSourceType)
                {
                    case DataSourceType.Excel:
                        return ExcelSheets != null && ExcelSheets.Count > 0;

                    case DataSourceType.MHT:

                        return m_wizardInfo != null &&
                                           m_wizardInfo.DataStorageInfos != null &&
                                           m_wizardInfo.DataStorageInfos.Count > 0;

                    default:
                        return false;
                }
            }
        }

        /// <summary>
        /// Current Data Source Type in selection
        /// </summary>
        public DataSourceType DataSourceType
        {
            get
            {
                return m_dataSourceType;
            }
            set
            {
                if (m_dataSourceType != value)
                {
                    Reset();
                    m_dataSourceType = value;
                    NotifyPropertyChanged("DataSourceType");
                }
            }
        }

        /// <summary>
        /// The string representation of Data Source Types(MHt /Excel)
        /// </summary>
        public List<string> DataSources
        {
            get;
            private set;
        }

        /// <summary>
        /// String form of Selected Data Source Type
        /// </summary>
        public string SelectedDataSource
        {
            get
            {
                return m_selectedDataSource;
            }
            set
            {
                if (m_dataSourceNameToTypeMapping.ContainsKey(value))
                {
                    m_selectedDataSource = value;
                    DataSourceType = m_dataSourceNameToTypeMapping[value];
                    NotifyPropertyChanged("SelectedDataSource");
                }
            }
        }

        /// <summary>
        /// IS MHT folder to take as input?
        /// </summary>
        public bool IsMHTFolder
        {
            get
            {
                return m_isMHTFolder;
            }
            set
            {
                m_isMHTFolder = value;
                if (IsMHTFolder)
                {
                    MHTFolderPath = MHTFolderPath;
                }
                else
                {
                    ListOfMHTsFilePath = ListOfMHTsFilePath;
                }
                NotifyPropertyChanged("IsMHTFolder");
            }
        }

        /// <summary>
        /// Is loading mht files from the folder?
        /// </summary>
        public bool IsLoadingFiles
        {
            get
            {
                return m_isLoadingFiles;
            }
            set
            {
                m_isLoadingFiles = value;
                NotifyPropertyChanged("IsLoadingFiles");
            }
        }

        /// <summary>
        /// MHT Source Folder Path
        /// </summary>
        public string MHTFolderPath
        {
            get
            {
                return m_mhtFolderPath;
            }
            set
            {
                if (m_worker != null)
                {
                    m_worker.CancelAsync();
                }
                m_mhtFolderPath = value;
                NotifyPropertyChanged("MHTFolderPath");
                MHTCount = 0;
                CanNext = false;
                try
                {
                    using (new AutoWaitCursor())
                    {
                        if (m_wizardInfo != null &&
                            !string.IsNullOrEmpty(MHTFolderPath))
                        {
                            m_wizardInfo.DataStorageInfos = new List<IDataStorageInfo>();
                            try
                            {
                                m_worker = new BackgroundWorker();
                                m_worker.WorkerSupportsCancellation = true;

                                // Setting the event handlers of background worker
                                m_worker.DoWork += new DoWorkEventHandler(GetFilesInBackgroundThread);
                                m_worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Worker_GetFilesInBackgroundThreadCompleted);
                                m_worker.RunWorkerAsync(null);
                            }
                            catch (UnauthorizedAccessException UAEx)
                            {
                                Warning = "Unauthorize error:" + UAEx.Message;
                            }
                        }
                    }
                }
                catch (IOException ioEx)
                {
                    MHTCount = 0;
                    Warning = ioEx.Message;
                }
                catch (WorkItemMigratorException te)
                {
                    MHTCount = 0;
                    Warning = te.Args.Title;
                }
            }
        }

        /// <summary>
        /// COunt of MHT files found
        /// </summary>
        public int MHTCount
        {
            get
            {
                return m_mhtCount;
            }
            set
            {
                m_mhtCount = value;
                NotifyPropertyChanged("MHTCount");
            }
        }

        /// <summary>
        /// Path of Text File containing mht file paths
        /// </summary>
        public string ListOfMHTsFilePath
        {
            get
            {
                return m_listOfMHTsFilePath;
            }
            set
            {
                if (m_worker != null)
                {
                    m_worker.CancelAsync();
                }
                m_listOfMHTsFilePath = value;
                NotifyPropertyChanged("ListOfMHTsFilePath");
                MHTCount = 0;
                CanNext = false;
                try
                {
                    using (new AutoWaitCursor())
                    {
                        if (m_wizardInfo != null && !string.IsNullOrEmpty(ListOfMHTsFilePath) && File.Exists(ListOfMHTsFilePath))
                        {
                            m_wizardInfo.DataStorageInfos = new List<IDataStorageInfo>();
                            using (StreamReader tr = new StreamReader(ListOfMHTsFilePath))
                            {
                                while (!tr.EndOfStream)
                                {
                                    try
                                    {
                                        string filePath = tr.ReadLine();
                                        if (File.Exists(filePath) && IsMHTFile(filePath))
                                        {
                                            MHTStorageInfo info = new MHTStorageInfo(filePath);
                                            m_wizardInfo.DataStorageInfos.Add(info);
                                            MHTCount++;
                                        }
                                    }
                                    catch (FileFormatException)
                                    { }
                                    catch (UnauthorizedAccessException)
                                    { }
                                }
                            }
                        }
                    }
                    CanNext = ValidatePartState();
                }
                catch (IOException)
                {
                    MHTCount = 0;
                }
                catch (WorkItemMigratorException te)
                {
                    MHTCount = 0;
                    Warning = te.Args.Title;
                }
            }
        }

        /// <summary>
        /// File Path for Excel Data Source
        /// </summary>
        public string ExcelFilePath
        {
            get
            {
                return m_excelFilePath;
            }
            set
            {
                m_excelFilePath = value.Trim();

                try
                {
                    if (!string.IsNullOrEmpty(m_excelFilePath))
                    {
                        LoadExcelSource();
                    }
                    else
                    {
                        App.CallMethodInUISynchronizationContext(ClearExcelSheets, null);
                        App.CallMethodInUISynchronizationContext(ClearFields, null);
                        SelectedExcelSheet = string.Empty;
                    }
                }
                catch (WorkItemMigratorException te)
                {
                    Warning = te.Args.Title;
                }
                NotifyPropertyChanged("ExcelFilePath");
                NotifyPropertyChanged("IsPreviewEnabled");
                NotifyPropertyChanged("ExcelSheetsCount");
            }
        }

        /// <summary>
        /// List of Field names found in Data Source
        /// </summary>
        public ObservableCollection<string> Fields
        {
            get
            {
                return m_fields;
            }
        }

        /// <summary>
        /// Count of Excel Sheets found in the data source
        /// </summary>
        public int ExcelSheetsCount
        {
            get
            {
                return ExcelSheets.Count;
            }
        }

        /// <summary>
        /// List of Excel Sheets to be shown on a selection of an excel file path
        /// </summary>
        public ObservableCollection<string> ExcelSheets
        {
            get;
            private set;
        }

        /// <summary>
        /// Currently selected excel file path
        /// </summary>
        public string SelectedExcelSheet
        {
            get
            {
                return m_selectedExcelSheet;
            }
            set
            {
                m_selectedExcelSheet = value;
                UpdateFields();
                NotifyPropertyChanged("SelectedExcelSheet");
            }
        }

        /// <summary>
        /// The Row in Excel WorkSheet that contains the Excel Field Names
        /// </summary>
        public string ExcelHeaderRow
        {
            get
            {
                return m_excelHeaderRow;
            }
            set
            {
                m_excelHeaderRow = value;
                UpdateFields();
                NotifyPropertyChanged("ExcelHeaderRow");
            }
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Resets this Wizard Part(Select Data Source)
        /// </summary>
        public override void Reset()
        {
            if (m_worker != null)
            {
                m_worker.CancelAsync();
            }

            if (m_wizardInfo != null)
            {
                if (m_wizardInfo.DataSourceParser != null)
                {
                    m_wizardInfo.DataSourceParser.Dispose();
                    m_wizardInfo.DataSourceParser = null;
                }
                m_wizardInfo.DataStorageInfos = null;
            }

            App.CallMethodInUISynchronizationContext(ClearExcelSheets, null);
            App.CallMethodInUISynchronizationContext(ClearFields, null);
            ExcelHeaderRow = DefaultExcelHeaderRow;
            ExcelFilePath = string.Empty;
            MHTFolderPath = null;
            ListOfMHTsFilePath = null;
        }

        /// <summary>
        /// Updates the Wizard with the Data Source Information
        /// </summary>
        /// <returns>Returns whether updation was successful or not</returns>
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

            if (WizardInfo.DataSourceType == DataSourceType.MHT)
            {
                WizardInfo.IsLinking = false;
            }

            try
            {
                if (DataSourceType == DataSourceType.MHT)
                {
                    LoadMHTSource();
                }

                // If part state is not valid return false
                if (m_wizardInfo.DataSourceParser == null)
                {
                    Warning = Resources.DataSourceInvalidFile;

                    return false;
                }


                if (DataSourceType == DataSourceType.Excel)
                {
                    m_wizardInfo.DataSourceParser.StorageInfo.FieldNames.Clear();
                    foreach (string field in m_fields)
                    {
                        m_wizardInfo.DataSourceParser.StorageInfo.FieldNames.Add(new SourceField(field, false));
                    }

                    m_wizardInfo.DataStorageInfos = new List<IDataStorageInfo> { m_wizardInfo.DataSourceParser.StorageInfo };
                }
                else if (DataSourceType == DataSourceType.MHT)
                {
                    m_wizardInfo.MHTSource = IsMHTFolder ? MHTFolderPath : ListOfMHTsFilePath;
                }

                // Update the Wizard Info with the current DataSourceType
                m_wizardInfo.DataSourceType = DataSourceType;

                if (m_wizardInfo.Migrator.SourceNameToFieldMapping != null)
                {
                    m_wizardInfo.Migrator.SourceNameToFieldMapping.Clear();
                }
                WizardInfo.Migrator.SourceNameToFieldMapping = null;
                m_prerequisite.Save();
                return true;
            }
            catch (WorkItemMigratorException te)
            {
                Warning = te.Args.Title;
                return false;
            }
        }

        /// <summary>
        /// Launches the Data Source Input File in the associated Application
        /// </summary>
        public void PreviewDataSourceFile()
        {
            // Check whether preview is enabled or not
            if (IsPreviewEnabled)
            {
                try
                {
                    switch (DataSourceType)
                    {
                        case DataSourceType.Excel:
                            ExcelParser.OpenWorkSheet(ExcelFilePath, SelectedExcelSheet);
                            break;

                        case DataSourceType.MHT:
                            break;

                        default:
                            throw new InvalidEnumArgumentException("Invalid Enum Value");
                    }
                }
                catch (WorkItemMigratorException te)
                {
                    MessageHelper.ShowMessageWindow(te.Args);
                }
            }
        }

        public MessageEventArgs ValidateExcelFilePath()
        {
            if (!IsPreviewEnabled)
            {
                var args = new MessageEventArgs();
                args.Title = Resources.DataSourceInvalidFile;
                args.LikelyCause = Resources.DataSourceInvalidFileLikelyCause;
                args.PotentialSolution = Resources.DataSourceInvalidFilePotentialSolution;
                return args;
            }
            return null;
        }

        public MessageEventArgs ValidateExcelHeaderRow()
        {
            int headerRow = -1;
            // If HeaderInfo is not valid raise the error and return false
            if (!int.TryParse(ExcelHeaderRow, out headerRow) || headerRow <= 0)
            {
                var args = new MessageEventArgs();
                args.Title = Resources.DataSourceInvalidHeaderInfo;
                args.LikelyCause = Resources.DataSourceInvalidHeaderInfoLikelyCause;
                args.PotentialSolution = Resources.DataSourceInvalidHeaderInfoPotentialSolution;
                return args;
            }
            return null;
        }

        #endregion

        #region protected/private methods

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

        /// <summary>
        /// Update the list of Field names
        /// </summary>
        private void UpdateFields()
        {
            App.CallMethodInUISynchronizationContext(ClearFields, null);
            if (m_wizardInfo == null || m_wizardInfo.DataSourceParser == null || string.IsNullOrEmpty(SelectedExcelSheet))
            {
                return;
            }

            int headerRow = -1;
            int.TryParse(ExcelHeaderRow, out headerRow);
            if (headerRow < 1)
            {
                Fields.Clear();
                return;
            }

            // If Data Source Wizard Part is in valid state and Storage Info is not null
            using (new AutoWaitCursor())
            {
                switch (DataSourceType)
                {
                    case DataSourceType.Excel:
                        ExcelStorageInfo xlInfo = m_wizardInfo.DataSourceParser.StorageInfo as ExcelStorageInfo;
                        xlInfo.WorkSheetName = SelectedExcelSheet;
                        xlInfo.RowContainingFieldNames = ExcelHeaderRow;
                        break;

                    case DataSourceType.MHT:
                        break;

                    default:
                        throw new InvalidEnumArgumentException("Invalid Enum Value");
                }
                try
                {
                    m_wizardInfo.DataSourceParser.ParseDataSourceFieldNames();

                    if (m_wizardInfo.DataSourceParser.StorageInfo.FieldNames != null)
                    {
                        m_fields.Clear();
                        foreach (SourceField field in m_wizardInfo.DataSourceParser.StorageInfo.FieldNames)
                        {
                            m_fields.Add(field.FieldName);
                        }
                    }
                }
                catch (WorkItemMigratorException)
                { }
            }
        }

        /// <summary>
        /// Initializes the Data Source
        /// </summary>
        private void InitializeDataSource()
        {
            if (m_wizardInfo.DataSourceParser != null)
            {
                m_wizardInfo.DataSourceParser.Dispose();
            }

            if (IsPreviewEnabled)
            {
                switch (DataSourceType)
                {
                    case DataSourceType.Excel:

                        ExcelStorageInfo xlInfo = new ExcelStorageInfo(ExcelFilePath);
                        m_wizardInfo.DataSourceParser = new ExcelParser(xlInfo);
                        break;

                    case DataSourceType.MHT:
                        string mhtFile = m_wizardInfo.DataStorageInfos[0].Source;
                        foreach (IDataStorageInfo info in m_wizardInfo.DataStorageInfos)
                        {
                            DateTime mhtFileDateTime = File.GetLastWriteTime(mhtFile);
                            DateTime currentFileDateTime = File.GetLastWriteTime(info.Source);
                            if (mhtFileDateTime < currentFileDateTime)
                            {
                                mhtFile = info.Source;
                            }
                        }

                        MHTStorageInfo mhtInfo = new MHTStorageInfo(mhtFile);
                        m_wizardInfo.DataSourceParser = new MHTParser(mhtInfo);
                        break;

                    default:
                        throw new InvalidEnumArgumentException("Invalid Enum Value");
                }
            }
        }

        /// <summary>
        /// Load Excel Data Source
        /// </summary>
        private void LoadExcelSource()
        {
            try
            {
                using (new AutoWaitCursor())
                {
                    LoadExcelWorkSheets();

                    InitializeDataSource();

                    SetDefaultExcelSheet();
                }
            }
            catch (WorkItemMigratorException)
            {
                App.CallMethodInUISynchronizationContext(ClearExcelSheets, null);
                App.CallMethodInUISynchronizationContext(ClearFields, null);
                throw;
            }
            catch (ArgumentException)
            {
                App.CallMethodInUISynchronizationContext(ClearExcelSheets, null);
                App.CallMethodInUISynchronizationContext(ClearFields, null);
            }
        }

        private void LoadMHTSource()
        {
            using (new AutoWaitCursor())
            {
                InitializeDataSource();
            }
        }

        private void SetDefaultExcelSheet()
        {
            if (IsPreviewEnabled)
            {
                //Set Selected Sheet to be the first Worksheet in the Excel Sheets
                SelectedExcelSheet = ExcelSheets[0];
            }
        }

        /// <summary>
        /// Validates Wizard Part State
        /// </summary>
        /// <returns></returns>
        public override bool ValidatePartState()
        {
            Warning = null;
            // If preview is not enable means that Data Source Input File is not valid
            // So return false
            if (DataSourceType == DataSourceType.Excel)
            {
                if (!IsPreviewEnabled)
                {
                    Warning = Resources.DataSourceInvalidFile;
                    return false;
                }
                else
                {
                    int headerRow = -1;
                    // If HeaderInfo is not valid raise the error and return false
                    if (!int.TryParse(ExcelHeaderRow, out headerRow) || headerRow <= 0)
                    {
                        Warning = Resources.DataSourceInvalidHeaderInfo;
                        return false;
                    }

                    if (Fields.Count == 0)
                    {
                        Warning = "Field Names are not found";
                        return false;
                    }

                }
            }
            else if (DataSourceType == DataSourceType.MHT)
            {
                if (IsMHTFolder)
                {
                    if (string.IsNullOrEmpty(MHTFolderPath))
                    {
                        Warning = "Invalid folder path";
                        return false;
                    }

                    else if (!Directory.Exists(MHTFolderPath))
                    {
                        Warning = "Invalid folder path";
                        return false;
                    }
                }

                if (!IsMHTFolder)
                {
                    if (string.IsNullOrEmpty(ListOfMHTsFilePath))
                    {
                        Warning = "Invalid text file path containing mht/word files";
                        return false;
                    }
                    else if (!File.Exists(ListOfMHTsFilePath) ||
                        String.CompareOrdinal(Path.GetExtension(ListOfMHTsFilePath), ".txt") != 0)
                    {
                        Warning = "Invalid text file path containing mht/word files";
                        return false;
                    }
                }

                if (MHTCount == 0)
                {
                    Warning = "No MHT/Word files are found";
                    return false;
                }
                if (m_isLoadingFiles)
                {
                    Warning = "All documents are not loaded";
                    return false;
                }


            }
            return true;
        }

        /// <summary>
        /// Load Excel WorkSheets
        /// </summary>
        private void LoadExcelWorkSheets()
        {
            // Clear the Excel Sheets 
            App.CallMethodInUISynchronizationContext(ClearExcelSheets, null);

            // Clear Selected Sheet to empty
            SelectedExcelSheet = string.Empty;

            string filePath = ExcelFilePath;

            // Get the List of Name of Worksheets
            List<string> worksheets = ExcelParser.GetWorksheetNames(filePath);

            //If List of worksheet names is not working then add them to ExcelWorksheets
            if (worksheets.Count > 0)
            {
                foreach (string worksheet in worksheets)
                {
                    ExcelSheets.Add(worksheet);
                }
            }
        }

        /// <summary>
        /// We can always initialize the Data Source
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
        protected override bool CanInitializeWizardPage(WizardInfo info)
        {
            m_canShow = true;
            return m_canShow;
        }

        protected override bool IsUpdationRequired()
        {
            if (m_prerequisite.IsDataSourceTypeModified(DataSourceType))
            {
                return true;
            }
            else if (DataSourceType == DataSourceType.Excel &&
                    m_prerequisite.IsExcelSourceModified(ExcelFilePath, SelectedExcelSheet, ExcelHeaderRow))
            {
                return true;
            }
            else if (DataSourceType == DataSourceType.MHT &&
                m_prerequisite.IsMHTSourceModified(IsMHTFolder, MHTFolderPath, ListOfMHTsFilePath))
            {
                return true;
            }
            return false;
        }

        private void GetFilesInBackgroundThread(object sender, DoWorkEventArgs e)
        {
            IsLoadingFiles = true;
            using (new AutoWaitCursor())
            {
                if (Directory.Exists(MHTFolderPath))
                {
                    GetMHTFiles(MHTFolderPath);
                }
            }
            IsLoadingFiles = false;
        }

        private void Worker_GetFilesInBackgroundThreadCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // If any exception is thrown during the work then sets the state as failed and set the corresponding message.
            if (e.Error != null)
            {
                throw e.Error;
            }
            if (m_worker != null)
            {
                m_worker.Dispose();
                m_worker = null;
            }
            CanNext = ValidatePartState();
        }

        private void GetMHTFiles(string rootDirectory)
        {
            try
            {
                // First Iterating Files
                foreach (string mhtFile in Directory.GetFiles(rootDirectory))
                {
                    try
                    {
                        if (!IsMHTFolder)
                        {
                            MHTCount = 0;
                            return;
                        }
                        if (IsMHTFile(mhtFile) && !Path.GetFileNameWithoutExtension(mhtFile).StartsWith("~", StringComparison.Ordinal))
                        {
                            if (m_worker == null ||
                                (m_worker != null && m_worker.CancellationPending))
                            {
                                return;
                            }
                            MHTStorageInfo info = new MHTStorageInfo(mhtFile);
                            m_wizardInfo.DataStorageInfos.Add(info);
                            MHTCount++;
                        }
                    }
                    catch (FileFormatException)
                    { }
                    catch (UnauthorizedAccessException)
                    { }
                    catch (ArgumentException)
                    { }
                }
                // Iterating Sub directories
                foreach (string directory in Directory.GetDirectories(rootDirectory))
                {
                    if (!IsMHTFolder)
                    {
                        MHTCount = 0;
                        return;
                    }
                    if (m_worker == null ||
                        (m_worker != null && m_worker.CancellationPending))
                    {
                        return;
                    }
                    GetMHTFiles(directory);
                }
            }
            catch (UnauthorizedAccessException)
            {
                MHTCount = 0;
            }
            catch (ArgumentException)
            {
                MHTCount = 0;
            }
        }

        private void ClearFields(object obj)
        {
            Fields.Clear();
        }

        private void ClearExcelSheets(object obj)
        {
            ExcelSheets.Clear();
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose()
        {
            if (m_worker != null)
            {
                m_worker.Dispose();
            }
        }

        #endregion
    }
}
