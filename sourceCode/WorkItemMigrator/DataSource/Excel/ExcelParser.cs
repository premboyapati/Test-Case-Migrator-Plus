//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.IO;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Text;
    using Microsoft.Office.Interop.Excel;

    /// <summary>
    /// Parses the Excel Data Source and generates field names and source workitems
    /// </summary>
    internal class ExcelParser : IDataSourceParser
    {
        #region Fields

        private ExcelStorageInfo m_storageInfo;
        // Excel Application used to handle workbooks for read/write
        private static Application s_application;

        // Excel Application used to open and view the workbooks
        private static Application s_userApplication;

        // The Excel Workbook that contains the Workitems to be imported into the TFS Server
        private Workbook m_workBook;

        // The current row of operation in the Excel Data Source
        private int m_currentRow;

        // The last row to be processed in Excel Data Source
        private int m_lastRow;

        // This is a column representing the mandatory field of the workitem which is to be imported.
        // It is used to find out the starting of new workitem
        private IList<int> m_separatingColumns = new List<int>();

        // This is the field name to column mapping of Excel fields that are going to be migrated
        private IDictionary<string, int> m_headersToColumn = new Dictionary<string, int>();

        // This is the field name to column mapping of Excel fields that are going to be migrated
        private IDictionary<string, int> m_selectedHeadersToColumn = new Dictionary<string, int>();

        private IDictionary<string, IList<string>> m_fieldNameToUniqueValuesMapping = new Dictionary<string, IList<string>>();

        private IDictionary<string, IWorkItemField> m_fieldNameToFields;

        private int m_headerRow;

        #endregion

        #region Constant

        public const string TestStepsFieldKeyName = "Test Steps";

        // The character used by TCM for identifying parameter
        private const string TCMParameterizationCharacter = "@";

        // White space character
        private const string WhiteSpaceCharacter = " ";

        #endregion

        #region Constructor

        /// <summary>
        /// Initialzes Excel Parser from Excel Storage Info
        /// </summary>
        /// <param name="info"></param>
        public ExcelParser(ExcelStorageInfo info)
        {
            if (info == null)
            {
                throw new ArgumentNullException("info");
            }
            try
            {

                // Gets the Workbook located at passed file path
                m_workBook = OpenWorkbookInReadMode(info.Source);

                StorageInfo = info;
                StorageInfo.PropertyChanged += new PropertyChangedEventHandler(StorageInfo_PropertyChanged);
                RawSourceWorkItems = new List<ISourceWorkItem>();
                ParsedSourceWorkItems = new List<ISourceWorkItem>();
                FieldNameToFields = new Dictionary<string, IWorkItemField>();
            }
            catch (COMException ex)
            {
                throw new WorkItemMigratorException("Unable to load excel workbook",
                                                    ex.Message,
                                                    ex.InnerException != null ? ex.InnerException.Message : string.Empty);
            }
            catch (InvalidCastException icEx)
            {
                throw new WorkItemMigratorException("Unable to load excel workbook",
                                                    icEx.Message,
                                                    icEx.InnerException != null ? icEx.InnerException.Message : string.Empty);
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Contains all information required by Excel Parser
        /// </summary>
        public IDataStorageInfo StorageInfo
        {
            get
            {
                return m_storageInfo;
            }
            private set
            {
                m_storageInfo = value as ExcelStorageInfo;
            }
        }

        /// <summary>
        /// List of workitems without any settings
        /// </summary>
        public IList<ISourceWorkItem> RawSourceWorkItems
        {
            get;
            private set;
        }

        /// <summary>
        /// List of Workitems with Parametrization and Multiline settings applied.
        /// </summary>
        public IList<ISourceWorkItem> ParsedSourceWorkItems
        {
            get;
            private set;
        }

        /// <summary>
        /// Field Name To Field Mapping
        /// </summary>
        public IDictionary<string, IWorkItemField> FieldNameToFields
        {
            get
            {
                return m_fieldNameToFields;
            }
            set
            {
                m_fieldNameToFields = value;
                SetSetparatingColumns();
                ResetPointerToNextWorkItem();
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Parses the Data Source For Field Names
        /// </summary>
        public void ParseDataSourceFieldNames()
        {
            try
            {
                // If HeadersToColumn Is Not Empty the we redundantly calling this method
                if (m_headersToColumn.Count > 0)
                {
                    return;
                }

                m_storageInfo.FieldNames = new List<SourceField>();

                // Gets the worksheet containg the workitems
                Worksheet workSheet = (Worksheet)m_workBook.Worksheets[m_storageInfo.WorkSheetName];

                // Gets the last Row containg data in the Excel Data Source Sheet
                Missing missing = Missing.Value;
                Range range = workSheet.Cells.Find("*", missing, missing, missing, XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, false, missing, missing);

                // This sheet does not contain any data so returns null
                if (range == null)
                {
                    return;
                }

                m_lastRow = range.Row;

                // Gets the header row number from the Data Store
                if (!int.TryParse(m_storageInfo.RowContainingFieldNames, out m_headerRow) || m_headerRow < 0)
                {
                    throw new WorkItemMigratorException("Incorrect Header Row", null, null);
                }

                m_currentRow = m_headerRow;
                var headers = m_storageInfo.FieldNames;

                // Getting the Excel Range that contains the Row having Excel Field names
                Range row = workSheet.Rows[m_headerRow] as Range;

                // getting Values present in that row as an array of values
                Array array = row.Value2 as Array;

                int currentColumn = 1;

                // Parsing through each value in that array and filling headers and header to column mapping
                foreach (object o in array)
                {
                    // Only fills the data structures if the value is not null
                    if (o != null)
                    {
                        // getting Field Name
                        string header = o.ToString().Trim();

                        // update the datastructure is the value is not null or empty
                        if (!string.IsNullOrEmpty(header))
                        {
                            // Add Field Name in Headers
                            SourceField field = new SourceField(o.ToString().Trim(), false);
                            headers.Add(field);

                            // Update the Header to column Mapping
                            m_headersToColumn[header] = currentColumn;
                        }
                    }

                    // Done with current column. increment it to parse next cell in the row.
                    currentColumn++;
                }
            }
            catch (COMException)
            {
                // This may be due to the corrupted excel file format
                throw new WorkItemMigratorException("Field names are not found",
                                                    Resources.DataSourceFieldNamesNotFoundErrorLikelyCause,
                                                    Resources.DataSourceFieldNamesNotFoundErrorPotentialSolution);
            }

        }

        /// <summary>
        /// Parses and returns the next workitem present in the Excel Source
        /// </summary>
        /// <returns></returns>
        public ISourceWorkItem GetNextWorkItem()
        {
            if (m_currentRow > m_lastRow)
            {
                return null;
            }

            int workItemStartingRow = -1;

            string error = null;

            // It represents the state that it has started the parsing of the excel workitem or not
            bool isReadingNextWorkItemStarted = false;

            // The Internal Representation of TFS workItem.
            ISourceWorkItem xlWorkItem = new SourceWorkItem();

            // The List of Test Steps
            List<SourceTestStep> steps = new List<SourceTestStep>();

            // While The end of work item is not encountered parse the current row and update the Excel Work Item
            while (!IsWorkItemCompleted(isReadingNextWorkItemStarted))
            {
                if (workItemStartingRow == -1)
                {
                    workItemStartingRow = m_currentRow;
                }

                m_storageInfo.ProgressPercentage = (m_currentRow * 100) / m_lastRow;
                // Parsing of work item is now started. So switch on the bool variable
                isReadingNextWorkItemStarted = true;

                string testStepTitle = string.Empty;
                string testStepExpectedResult = string.Empty;

                // Itrerating throgh each fields and reading value present in corresponding cells and then updating
                // m_dataStore.DataValuesByFieldName and Excel WorkItem
                foreach (KeyValuePair<string, int> kvp in m_selectedHeadersToColumn)
                {
                    // Get Value at(currentrow,columnForCurrentHeaderInProcess)
                    string value = GetValueAt(m_currentRow, kvp.Value);
                    if (!string.IsNullOrEmpty(value))
                    {
                        // If the Current field is mapped to TFS Test Step Title field then update the step's title
                        if (FieldNameToFields[kvp.Key] is TestStepTitleField)
                        {
                            testStepTitle = value;
                        }
                        // else if the Current field is mapped to TFS Test Step Expected result field then update the step's expected result
                        else if (FieldNameToFields[kvp.Key] is TestStepExpectedResultField)
                        {
                            testStepExpectedResult = value;
                        }
                        // else if it is mapped to a date time field then parse it
                        else if (FieldNameToFields[kvp.Key].Type == typeof(DateTime))
                        {
                            double excelDate;
                            if (double.TryParse(value, out excelDate))
                            {
                                DateTime dateOfReference = new DateTime(1900, 1, 1);
                                excelDate = excelDate - 2;
                                try
                                {
                                    xlWorkItem.FieldValuePairs[kvp.Key] = dateOfReference.AddDays(excelDate);
                                }
                                catch (ArgumentException)
                                {
                                    // If argument provided is wrong then don't set anyting
                                    xlWorkItem.FieldValuePairs[kvp.Key] = null;
                                }
                            }
                            else
                            {
                                xlWorkItem.FieldValuePairs[kvp.Key] = null;
                            }
                        }
                        // else just update the exel workitem
                        else
                        {
                            if (!xlWorkItem.FieldValuePairs.ContainsKey(kvp.Key))
                            {
                                xlWorkItem.FieldValuePairs.Add(kvp.Key, value);
                            }
                            else
                            {
                                xlWorkItem.FieldValuePairs[kvp.Key] += "\n" + value;
                            }
                        }
                    }
                    // else if this field is not already filled and it is not test step(title or expected result) field then set it to null.
                    else if (!xlWorkItem.FieldValuePairs.ContainsKey(kvp.Key) &&
                             !(FieldNameToFields[kvp.Key] is TestStepTitleField) &&
                             !(FieldNameToFields[kvp.Key] is TestStepExpectedResultField))
                    {
                        xlWorkItem.FieldValuePairs.Add(kvp.Key, null);
                    }
                }

                // If we found step in the current row then update the list of excel steps
                if (!string.IsNullOrEmpty(testStepTitle) || !string.IsNullOrEmpty(testStepExpectedResult))
                {
                    SourceTestStep step = new SourceTestStep();
                    step.title = testStepTitle.Trim();
                    step.expectedResult = testStepExpectedResult.Trim();
                    steps.Add(step);
                }

                // Done with parsing of current row. Move to next row
                m_currentRow++;
            }

            if (!string.IsNullOrEmpty(StorageInfo.SourceIdFieldName))
            {
                string value = GetValueAt(workItemStartingRow, m_headersToColumn[StorageInfo.SourceIdFieldName]);
                if (string.IsNullOrEmpty(value))
                {
                    error += "Source Id Is Not found\n";
                }
                else
                {
                    xlWorkItem.SourceId = value;
                }
            }
            if (!string.IsNullOrEmpty(StorageInfo.TestSuiteFieldName))
            {
                string value = GetValueAt(workItemStartingRow, m_headersToColumn[StorageInfo.TestSuiteFieldName]);
                if (!string.IsNullOrEmpty(value))
                {
                    foreach (string testSuite in value.Split(';', '\n'))
                    {
                        string s = testSuite.Trim();
                        if (!string.IsNullOrEmpty(s))
                        {
                            xlWorkItem.TestSuites.Add(s);
                        }
                    }
                }
            }

            foreach (ILinkRule linkInfo in StorageInfo.LinkRules)
            {
                string value = GetValueAt(workItemStartingRow, m_headersToColumn[linkInfo.SourceFieldNameOfEndWorkItemCategory]);
                if (!string.IsNullOrEmpty(value))
                {
                    foreach (string linkedWorkItemId in value.Split(';'))
                    {
                        string trimmedId = linkedWorkItemId.Trim();
                        if (!string.IsNullOrEmpty(trimmedId))
                        {
                            Link link = new Link();
                            link.StartWorkItemCategory = linkInfo.StartWorkItemCategory;
                            link.StartWorkItemSourceId = xlWorkItem.SourceId;
                            link.LinkTypeName = linkInfo.LinkTypeReferenceName;
                            link.EndWorkItemCategory = linkInfo.EndWorkItemCategory;
                            link.EndWorkItemSourceId = trimmedId;
                            xlWorkItem.Links.Add(link);
                        }
                    }
                }
            }

            if (steps != null && steps.Count > 0)
            {
                // Done with parsing of XL Work Item. Add the List of steps in the XL Work Item.
                xlWorkItem.FieldValuePairs.Add(TestStepsFieldKeyName, steps);
            }

            xlWorkItem.SourcePath = StorageInfo.Source;

            if (!string.IsNullOrEmpty(error))
            {
                xlWorkItem = new FailedSourceWorkItem(xlWorkItem, error);
            }

            RawSourceWorkItems.Add(xlWorkItem);

            xlWorkItem = ApplySettings(xlWorkItem);

            ParsedSourceWorkItems.Add(xlWorkItem);

            // return the Excel WorkItem
            return xlWorkItem;

        }

        #endregion

        #region Private Methods

        private void StorageInfo_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (String.CompareOrdinal(e.PropertyName, "WorkSheetName") == 0 ||
                String.CompareOrdinal(e.PropertyName, "HeaderRow") == 0)
            {
                m_lastRow = -1;
                m_headerRow = -1;
                m_separatingColumns.Clear();
                m_headersToColumn.Clear();
                m_selectedHeadersToColumn.Clear();
                m_fieldNameToUniqueValuesMapping.Clear();
            }
        }

        private void SetSetparatingColumns()
        {
            m_selectedHeadersToColumn.Clear();
            m_separatingColumns.Clear();

            foreach (var kvp in FieldNameToFields)
            {
                if (m_headersToColumn.ContainsKey(kvp.Key))
                {
                    m_selectedHeadersToColumn.Add(kvp.Key, m_headersToColumn[kvp.Key]);
                    if (kvp.Value.IsMandatory)
                    {
                        m_separatingColumns.Add(m_headersToColumn[kvp.Key]);
                    }
                }
            }
        }

        private ISourceWorkItem ApplySettings(ISourceWorkItem xlWorkItem)
        {
            if (!StorageInfo.IsMultilineSense &&
                string.IsNullOrEmpty(StorageInfo.StartParameterizationDelimeter)
                && string.IsNullOrEmpty(StorageInfo.EndParameterizationDelimeter))
            {
                return xlWorkItem;
            }
            ISourceWorkItem parsedSourceWorkItem = new SourceWorkItem();
            foreach (var kvp in xlWorkItem.FieldValuePairs)
            {
                if (String.CompareOrdinal(kvp.Key, TestStepsFieldKeyName) != 0)
                {
                    parsedSourceWorkItem.FieldValuePairs.Add(kvp.Key, kvp.Value);
                }
                else
                {
                    var newSteps = new List<SourceTestStep>();
                    List<SourceTestStep> steps = kvp.Value as List<SourceTestStep>;
                    foreach (var step in steps)
                    {
                        if (StorageInfo.IsMultilineSense)
                        {
                            List<string> testStepTitles = new List<string>();
                            List<string> testStepExpectedResults = new List<string>();
                            foreach (string s in step.title.Split('\r', '\n'))
                            {
                                testStepTitles.Add(s);
                            }
                            foreach (string s in step.expectedResult.Split('\r', '\n'))
                            {
                                testStepExpectedResults.Add(s);
                            }
                            // If we found step in the current row then update the list of excel steps
                            int i = 0;
                            while (i < testStepTitles.Count || i < testStepExpectedResults.Count)
                            {
                                string title = string.Empty;
                                string expectedResult = string.Empty;
                                if (i < testStepTitles.Count)
                                {
                                    title = CreateParameterizedText(testStepTitles[i]);
                                }
                                if (i < testStepExpectedResults.Count)
                                {
                                    expectedResult = CreateParameterizedText(testStepExpectedResults[i]);
                                }
                                if (!string.IsNullOrWhiteSpace(title) || !string.IsNullOrWhiteSpace(expectedResult))
                                {
                                    SourceTestStep newStep = new SourceTestStep();
                                    newStep.title = title.Trim();
                                    newStep.expectedResult = expectedResult.Trim();
                                    newSteps.Add(newStep);
                                }
                                i++;
                            }
                        }
                        else
                        {
                            SourceTestStep newStep = new SourceTestStep();
                            newStep.title = CreateParameterizedText(step.title).Trim();
                            newStep.expectedResult = CreateParameterizedText(step.expectedResult).Trim();
                            newSteps.Add(newStep);
                        }
                    }
                    parsedSourceWorkItem.FieldValuePairs.Add(TestStepsFieldKeyName, newSteps);
                }
            }
            parsedSourceWorkItem.SourcePath = xlWorkItem.SourcePath;
            return parsedSourceWorkItem;
        }

        /// <summary>
        /// Utility method to know whether parsing of current Excel workitem is completed or not.
        /// </summary>
        /// <param name="isReadingNextWorkItemStarted"></param>
        /// <returns></returns>
        private bool IsWorkItemCompleted(bool isReadingNextWorkItemStarted)
        {
            // If parsed the whole data source then return true
            if (m_currentRow > m_lastRow)
            {
                return true;
            }

            // If reading new Workitem is not started then return false
            if (isReadingNextWorkItemStarted == false)
            {
                ResetPointerToNextWorkItem();
                return false;
            }

            // If encounetres the header Row then assume it to be end of the workitem and return true
            if (IsHeaderRow(m_currentRow))
            {
                m_currentRow++;
                return true;
            }

            // Swallow all empty lines before next data line
            while (m_currentRow <= m_lastRow && IsEmptyLine())
            {
                m_currentRow++;
            }

            // It will be new Work Item if the value at the cell corresponding to the mandatory field is not empty
            foreach (int column in m_separatingColumns)
            {
                if (!string.IsNullOrEmpty(GetValueAt(m_currentRow, column)))
                {
                    return true;
                }
            }
            return false;
        }


        /// <summary>
        /// Checks whether current is replica of header row or not
        /// </summary>
        /// <returns></returns>
        private bool IsHeaderRow(int row)
        {
            bool isHeaderRow = true;
            foreach (KeyValuePair<string, int> kvp in m_selectedHeadersToColumn)
            {
                if (String.CompareOrdinal(kvp.Key, GetValueAt(row, kvp.Value)) != 0)
                {
                    isHeaderRow = false;
                    break;
                }
            }
            return isHeaderRow;
        }

        /// <summary>
        /// Checks whether current row is empty or not
        /// </summary>
        /// <returns></returns>
        private bool IsEmptyLine()
        {
            bool isEmptyRow = true;
            foreach (KeyValuePair<string, int> kvp in m_selectedHeadersToColumn)
            {
                if (!string.IsNullOrEmpty(kvp.Key))
                {
                    isEmptyRow = false;
                    break;
                }
            }
            return isEmptyRow;
        }

        /// <summary>
        /// Return the value at (rowNo,columnNo) in current excelsheet
        /// </summary>
        /// <param name="rowNo"></param>
        /// <param name="columnNo"></param>
        /// <returns></returns>
        private string GetValueAt(int rowNo, int columnNo)
        {
            Worksheet workSheet = (Worksheet)m_workBook.Worksheets[m_storageInfo.WorkSheetName];
            Range range = (Range)(workSheet.Cells.get_Item(rowNo, columnNo));

            if (range.Value2 != null &&
                !string.IsNullOrEmpty(range.Value2.ToString().Trim()))
            {
                return range.Value2.ToString().Trim();
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Resets the current row to the starting of next Work Item
        /// </summary>
        private void ResetPointerToNextWorkItem()
        {
            while (m_currentRow <= m_lastRow)
            {
                if (!IsHeaderRow(m_currentRow))
                {
                    foreach (int column in m_separatingColumns)
                    {
                        if (!string.IsNullOrEmpty(GetValueAt(m_currentRow, column)))
                        {
                            return;
                        }
                    }
                }
                m_currentRow++;
            }
        }

        /// <summary>
        /// Creates TCM Parameterized string from Data Source test step string
        /// </summary>
        /// <param name="rawText"></param>
        /// <returns></returns>
        private string CreateParameterizedText(string rawText)
        {
            // If Data Source steps are having only start Delimeter
            if (!string.IsNullOrEmpty(StorageInfo.StartParameterizationDelimeter) && string.IsNullOrEmpty(StorageInfo.EndParameterizationDelimeter))
            {
                return RemoveStartDelimeter(rawText);
            }
            // else if Data Source test Steps are having only end delimeter
            else if (string.IsNullOrEmpty(StorageInfo.StartParameterizationDelimeter) && !string.IsNullOrEmpty(StorageInfo.EndParameterizationDelimeter))
            {
                return RemoveEndDelimeter(rawText);
            }
            //else if Data Source test steps are having both start and end delimeters
            else if (!string.IsNullOrEmpty(StorageInfo.StartParameterizationDelimeter) && !string.IsNullOrEmpty(StorageInfo.EndParameterizationDelimeter))
            {
                return RemoveStartAndEndDelimeter(rawText);
            }
            else
            {
                return rawText;
            }
        }

        /// <summary>
        /// Parses the test step texts and remove Data Source's StartDelimeter with TCM parameterization character
        /// </summary>
        /// <param name="rawText"></param>
        /// <returns></returns>
        private string RemoveStartDelimeter(string rawText)
        {
            // Getting Array of words from text
            string[] words = rawText.Split(' ');
            StringBuilder resultString = new StringBuilder();

            // iterating through each word
            foreach (string word in words)
            {
                string resultWord = word;

                // if word is having Data Source Start Delimeter then replace it with TCM's start Delimeter
                if (word.IndexOf(StorageInfo.StartParameterizationDelimeter, StringComparison.Ordinal) == 0)
                {
                    resultWord = TCMParameterizationCharacter + resultWord.Substring(StorageInfo.StartParameterizationDelimeter.Length);
                }
                // else if word is having TCM's start delimeter then remove it
                else if (resultWord.IndexOf(TCMParameterizationCharacter, StringComparison.Ordinal) == 0)
                {
                    resultWord = resultWord.Substring(1);
                }

                // Add the updated wrd in result string
                resultString.Append(resultWord);
                resultString.Append(WhiteSpaceCharacter);
            }
            return resultString.ToString(0, resultString.Length - 1);
        }

        /// <summary>
        /// Parses the test step texts and remove Data Source's EndDelimeter with TCM parameterization character
        /// </summary>
        /// <param name="rawText"></param>
        /// <returns></returns>
        private string RemoveEndDelimeter(string rawText)
        {
            // Getting Array of words from text
            string[] words = rawText.Split(' ');
            StringBuilder resultString = new StringBuilder();

            // iterating through each word
            foreach (string word in words)
            {
                string resultWord = word;

                // If word is having TCM's start delimeter then remove it
                if (resultWord.IndexOf(TCMParameterizationCharacter, StringComparison.Ordinal) == 0)
                {
                    resultWord = resultWord.Substring(1);
                }

                // if word is having Data Source End Delimeter then remove it and add TCM's start Delimeter
                if (resultWord.LastIndexOf(StorageInfo.EndParameterizationDelimeter, StringComparison.Ordinal) == (resultWord.Length - StorageInfo.EndParameterizationDelimeter.Length))
                {
                    resultWord = TCMParameterizationCharacter + resultWord.Substring(0, (resultWord.Length - StorageInfo.EndParameterizationDelimeter.Length));
                }

                // Add the updated wrd in result string
                resultString.Append(resultWord);
                resultString.Append(WhiteSpaceCharacter);
            }
            return resultString.ToString(0, resultString.Length - 1);
        }

        /// <summary>
        /// Parses the test step texts and remove Data Source's Satrt and EndDelimeter with TCM parameterization character
        /// </summary>
        /// <param name="rawText"></param>
        /// <returns></returns>
        private string RemoveStartAndEndDelimeter(string rawText)
        {
            // Getting Array of words from text
            string[] words = rawText.Split(' ');
            StringBuilder resultString = new StringBuilder();

            // iterating through each word
            foreach (string word in words)
            {
                string resultWord = word;

                // if word is having Data Source Start and End Delimeter then remove it and add TCM's start Delimeter
                if (word.IndexOf(StorageInfo.StartParameterizationDelimeter, StringComparison.Ordinal) == 0 &&
                    word.LastIndexOf(StorageInfo.EndParameterizationDelimeter, StringComparison.Ordinal) == (word.Length - StorageInfo.EndParameterizationDelimeter.Length))
                {
                    resultWord = TCMParameterizationCharacter + word.Substring(StorageInfo.StartParameterizationDelimeter.Length,
                                                      (word.Length - StorageInfo.EndParameterizationDelimeter.Length - StorageInfo.StartParameterizationDelimeter.Length));
                }
                // Else if word is having TCM's start delimeter then remove it
                else if (word.IndexOf(TCMParameterizationCharacter, StringComparison.Ordinal) == 0)
                {
                    resultWord = resultWord.Substring(1);
                }

                // Append updated word in result string
                resultString.Append(resultWord);
                resultString.Append(WhiteSpaceCharacter);
            }
            return resultString.ToString(0, resultString.Length - 1);
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose()
        {
            if (m_workBook != null)
            {
                CloseWorkbook(ref m_workBook, true);
                m_workBook = null;
                GC.WaitForPendingFinalizers();
            }
        }

        #endregion

        #region static_members

        private static Application UserApplication
        {
            get
            {
                if (s_userApplication == null)
                {
                    s_userApplication = new Application();
                    s_userApplication.Visible = true;
                    s_userApplication.UserControl = true;
                }
                return s_userApplication;
            }
        }

        private static Application Application
        {
            get
            {
                if (s_application == null)
                {
                    try
                    {
                        s_application = new Application();
                        s_application.Visible = false;
                        s_application.UserControl = false;
                    }
                    catch (COMException)
                    {
                        // Exception during initialization hints towards corrupt/not installation of Excel Application
                        throw new WorkItemMigratorException(Resources.WizardInitializationError,
                                                            Resources.WizardInitializationErrorLikelyCause,
                                                            Resources.WizardInitializationErrorPotentialSolution);
                    }
                }
                return s_application;
            }
        }

        /// <summary>
        /// Returns the List of names of all worksheets present in the workbook located at absolute 'filePath'
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static List<string> GetWorksheetNames(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException("filePath");
            }

            List<string> worksheets = new List<string>();
            if (File.Exists(filePath))
            {
                try
                {
                    Workbook workbook = OpenWorkbookInReadMode(filePath);
                    foreach (Worksheet worksheet in workbook.Worksheets)
                    {
                        worksheets.Add(worksheet.Name);
                    }
                    CloseWorkbook(ref workbook, true);
                }
                catch (COMException comEx)
                {
                    throw new WorkItemMigratorException(Resources.ExcelLoadError,
                                                        comEx.Message,
                                                        comEx.InnerException != null ? comEx.InnerException.Message : string.Empty);
                }
                catch (InvalidCastException icEx)
                {
                    throw new WorkItemMigratorException(Resources.ExcelLoadError,
                                                        icEx.Message,
                                                        icEx.InnerException != null ? icEx.InnerException.Message : string.Empty);
                }
            }
            return worksheets;
        }

        /// <summary>
        /// Opens the Workbook located at 'filepath' with Worksheet having name 'worksheetname' as active sheet.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="workSheetName"></param>
        public static void OpenWorkSheet(string filePath, string workSheetName)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException("filePath");
            }

            try
            {
                Application app = UserApplication;
                Workbook book = app.Workbooks.Open(filePath);
                app.Visible = true;
                if (!string.IsNullOrEmpty(workSheetName))
                {
                    Worksheet sheet = book.Worksheets[workSheetName] as Worksheet;
                    sheet.Select(Type.Missing);
                }
            }
            catch (COMException ex)
            {
                throw new WorkItemMigratorException("Unable to load excel workbook",
                                                    ex.Message,
                                                    ex.InnerException != null ? ex.InnerException.Message : string.Empty);
            }
            catch (InvalidCastException icEx)
            {
                throw new WorkItemMigratorException("Unable to load excel workbook",
                                                    icEx.Message,
                                                    icEx.InnerException != null ? icEx.InnerException.Message : string.Empty);
            }
        }

        /// <summary>
        /// Open Workbook for reading
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private static Workbook OpenWorkbookInReadMode(string filePath)
        {
            try
            {
                // This workbook can be opened by user externally and this may interfere with the current
                // processing of workbook. So first copying the file to temp file and then opening it for parsing
                string newFilePath = Path.GetTempFileName() + Path.GetExtension(filePath);
                File.Copy(filePath, newFilePath);
                return Application.Workbooks.Open(newFilePath);
            }
            catch (IOException e)
            {
                throw new WorkItemMigratorException(Resources.ExcelLoadError,
                                                    e.Message,
                                                    e.InnerException != null ? e.InnerException.Message : null);
            }
            catch (AccessViolationException accEx)
            {
                throw new WorkItemMigratorException(Resources.ExcelLoadError,
                                                    accEx.Message,
                                                    accEx.InnerException != null ? accEx.InnerException.Message : null);
            }
            catch (COMException comEx)
            {
                throw new WorkItemMigratorException(Resources.ExcelLoadError,
                                                    comEx.Message,
                                                    comEx.InnerException != null ? comEx.InnerException.Message : null);
            }
        }

        /// <summary>
        /// Closes the workbook for reading.
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="delete">Whether to delete excel file after close.</param>
        private static void CloseWorkbook(ref Workbook workBook, bool delete)
        {
            try
            {
                string filePath = workBook.FullName;
                workBook.Close(true);
                workBook = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // if delete file is set then delete it.
                if (delete)
                {
                    File.Delete(filePath);
                }
            }
            catch (UnauthorizedAccessException)
            { }
            catch (COMException)
            { }
            catch (IOException)
            { }
        }

        public static void Quit()
        {
            try
            {
                if (s_application != null)
                {
                    try
                    {
                        s_application.Workbooks.Close();
                    }
                    catch (COMException)
                    { }
                    s_application.Quit();
                    s_application = null;
                }

                if (s_userApplication != null && s_userApplication.Workbooks.Count == 0)
                {
                    try
                    {
                        s_userApplication.Visible = false;
                        s_userApplication.UserControl = false;

                        s_userApplication.Workbooks.Close();
                    }
                    catch (InvalidComObjectException)
                    { }
                    catch (COMException)
                    { }
                    catch (InvalidCastException)
                    { }
                    s_userApplication.Quit();
                    s_userApplication = null;
                }
            }
            catch (InvalidComObjectException)
            { }
            catch (COMException)
            { }
            catch (InvalidCastException)
            { }
            finally
            {
                s_application = null;
                s_userApplication = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        #endregion
    }
}
