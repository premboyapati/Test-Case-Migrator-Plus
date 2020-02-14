//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Text;
    using Microsoft.Office.Interop.Excel;

    internal class ExcelReporter : NotifyPropertyChange, IReporter
    {
        #region Fields

        IDictionary<string, IWorkItemField> m_sourceNameToFieldMapping;

        private string m_reportFile;

        private Application m_application;

        // Workbook into which the success/failure of migrated workitems to be written
        private Workbook m_reportBook;

        // The current row of the processing in the worksheet having successfully migrated workitems
        private int m_currentSuccessRow;

        // The current row of the processing in the worksheet having successfully migrated workitems with warnings
        private int m_currentWarningRow;

        // The current row of the processing in the worksheet having workitems which are failed to migrate
        private int m_currentErrorRow;

        // The Field Name to Column mapping of Succcessful Migrated Workitems
        private Dictionary<string, int> m_successHeaderToColumn;

        // The Field Name to Column mapping of Succcessful Migrated Workitems with Warnings
        private Dictionary<string, int> m_warningHeaderToColumn;

        // The Field Name to Column mapping of Failed Workitems
        private Dictionary<string, int> m_errorHeaderToColumn;

        // Field name for Test Step in Excel
        private string m_stepTitleHeader;

        // Field name for Test Step Expected Result in Excel
        private string m_stepExpectedResultHeader;

        private WizardInfo m_wizardInfo;


        #endregion

        #region Constants

        // Field Name for ID of Successfully Migrated Workitems
        private const string IDField = "TFS Workitem ID";

        // Field Name for Failed Workitems showing the error occured while migration
        private const string ErrorField = "Error";

        // Field Name for Warning Workitems showing the waring occured while migration
        private const string WarningField = "Warning";

        // The default width of the Columns of Excel Report 
        private const int DefaultColumnWidth = 30;

        // The Worksheet Name of the sheet containg successful workitems
        public const string SuccessSheetName = "Passed";

        // The Worksheet Name of the sheet containg successful workitems with warning
        public const string WarningSheetName = "Warning";

        // The Worksheet name for the sheet having failed workitems
        public const string ErrorSheetName = "Failed";

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor to initialize Success and Failed Sheet. Initializes corresponding header to column Mapping. 
        /// Gets test step/expected field name.
        /// </summary>
        /// <param name="headerToColumn">The Orignal Header to column mapping</param>
        /// <param name="columnMapping">Excel Field Name to TFS Workitem field name mapping</param>
        public ExcelReporter(WizardInfo wizardInfo)
        {
            m_wizardInfo = wizardInfo;
            m_sourceNameToFieldMapping = m_wizardInfo.Migrator.SourceNameToFieldMapping;
        }

        #endregion

        #region Properties

        public string ReportFile
        {
            get
            {
                return m_reportFile;
            }
            set
            {
                m_reportFile = value;
                NotifyPropertyChanged("ReportFile");
            }
        }

        #endregion

        #region Public Methods

        public void AddEntry(ISourceWorkItem sourceWorkItem)
        {
            if (m_application == null)
            {
                Initialize();
            }

            var warningSourceWorkItem = sourceWorkItem as WarningSourceWorkItem;
            if (warningSourceWorkItem != null)
            {
                ReportWarning(warningSourceWorkItem);
            }
            else
            {
                var passedSourceWorkItem = sourceWorkItem as PassedSourceWorkItem;
                if (passedSourceWorkItem != null)
                {
                    ReportSuccess(passedSourceWorkItem);
                }
                else
                {
                    var failedSourceWorkItem = sourceWorkItem as FailedSourceWorkItem;
                    if (failedSourceWorkItem != null)
                    {
                        ReportError(failedSourceWorkItem);
                    }
                    else
                    {
                        var skippedSourceWorkItem = sourceWorkItem as SkippedSourceWorkItem;
                        if (skippedSourceWorkItem != null)
                        {
                            string category = m_wizardInfo.WorkItemGenerator.WorkItemTypeToCategoryMapping[m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName];
                            ReportWarning(new WarningSourceWorkItem(skippedSourceWorkItem,
                                                                    m_wizardInfo.LinksManager.WorkItemCategoryToIdMappings[category][sourceWorkItem.SourceId].TfsId,
                                                                    "This work item is already migrated"));
                        }
                        else
                        {
                            throw new WorkItemMigratorException("Incorrect Source workItem Type", null, null);
                        }
                    }
                }
            }
        }

        public void Publish()
        {
            try
            {
                // if report file already exists then delete it
                if (File.Exists(ReportFile))
                {
                    File.Delete(ReportFile);
                }

                string directory = Path.GetDirectoryName(ReportFile);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                string settingsPath = Path.Combine(directory, m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName + "-settings.xml");
                m_wizardInfo.SaveSettings(settingsPath);

                try
                {
                    m_reportBook.SaveAs(ReportFile, XlFileFormat.xlExcel8);
                }
                catch (COMException)
                {
                    m_reportBook.SaveAs(ReportFile);
                }


                // Close the workbook after save
                m_reportBook.Close();
            }
            // throw WorkitemMigratorexception of publish report failure in case of any exception
            catch (COMException comEx)
            {
                throw new WorkItemMigratorException(Resources.PublisReportError,
                                                    comEx.Message,
                                                    comEx.InnerException != null ? comEx.InnerException.Message : string.Empty);
            }
            catch (InvalidCastException icEx)
            {
                throw new WorkItemMigratorException(Resources.PublisReportError,
                                                    icEx.Message,
                                                    icEx.InnerException != null ? icEx.InnerException.Message : string.Empty);
            }

        }

        #endregion

        #region Private methods

        private void Initialize()
        {
            m_application = new Application();

            // Creating Workbook that will contain info related to migrated workitems
            m_application.SheetsInNewWorkbook = 1;
            m_reportBook = m_application.Workbooks.Add(Type.Missing);

            // Initializes Worksheet containing failed workitems
            InitErrorSheet(m_sourceNameToFieldMapping);

            InitWarningSheet(m_sourceNameToFieldMapping);

            // Initializes Worksheet containing successfully migrated workitems
            InitSuccessSheet(m_sourceNameToFieldMapping);

            // Retrieving Excel Field names corresponding to Test Step/Expected Result fields of a workitem(testcase)
            foreach (var kvp in m_sourceNameToFieldMapping)
            {
                if (kvp.Value is TestStepExpectedResultField)
                {
                    m_stepExpectedResultHeader = kvp.Key;
                }
                else if (kvp.Value is TestStepTitleField)
                {
                    m_stepTitleHeader = kvp.Key;
                }
            }
        }

        /// <summary>
        /// Initializes the Warning sheet of the Report with field names.
        /// Also initializes the internal data structures needed for manipulating warning sheet.
        /// </summary>
        /// <param name="headerToColumn"></param>
        private void InitWarningSheet(IDictionary<string, IWorkItemField> sourceNameToFieldMapping)
        {
            // Creating Warning Worksheet
            Worksheet warningSheet = m_reportBook.Worksheets.Add() as Worksheet;
            warningSheet.Name = WarningSheetName;

            // Initializes the Header to column mapping for warning sheet
            m_warningHeaderToColumn = new Dictionary<string, int>();

            int column = 1;

            // Setting current row of operation in warning sheet to one
            m_currentWarningRow = 1;

            WriteFieldName(m_warningHeaderToColumn, warningSheet, m_currentWarningRow, ref column, IDField);

            column = WriteFieldNames(sourceNameToFieldMapping, m_warningHeaderToColumn, warningSheet, m_currentWarningRow, column);

            WriteFieldName(m_warningHeaderToColumn, warningSheet, m_currentWarningRow, ref column, WarningField);

            // Setting the Widths of Sheets's Columns
            SetDefaultWidths(warningSheet, column);

            // Skipping one line so that Failed workitems can be inserted after one blank line of the field name
            m_currentWarningRow += 2;

        }

        /// <summary>
        /// Initializes the worksheet that will contain the Successfully migrated workitems.
        /// Also initializes m_successHeaderToColumn(Header to Column mapping for successful migrated 
        /// workitems) and m_currentSuccessRow(the current row of processing in this sheet)
        /// </summary>
        /// <param name="headerToColumn">The orignal header to column mapping</param>
        private void InitSuccessSheet(IDictionary<string, IWorkItemField> sourceNameToFieldMapping)
        {
            // Creating Success Worksheet
            Worksheet successSheet = m_reportBook.Worksheets.Add() as Worksheet;
            successSheet.Name = SuccessSheetName;

            // Initializes the Header to column mapping for success sheet
            m_successHeaderToColumn = new Dictionary<string, int>();

            // Setting current row of operation in success sheet to one
            m_currentSuccessRow = 1;

            int column = 1;

            WriteFieldName(m_successHeaderToColumn, successSheet, m_currentSuccessRow, ref column, IDField);

            column = WriteFieldNames(sourceNameToFieldMapping, m_successHeaderToColumn, successSheet, m_currentSuccessRow, column);

            // Setting the Widths of Sheets's Columns
            SetDefaultWidths(successSheet, column);

            // Skipping one line so that Successfull workitems can be inserted after one blank line of the field name
            m_currentSuccessRow += 2;
        }

        /// <summary>
        /// Initializes the worksheet that will contain the workitems which are failed to migrate.
        /// Also initializes m_errorHeaderToColumn(Header to Column mapping for failed workitems) 
        /// and m_currentErrorRow(the current row of processing in this sheet)
        /// </summary>
        /// <param name="headerToColumn"></param>
        private void InitErrorSheet(IDictionary<string, IWorkItemField> sourceNameToFieldMapping)
        {
            // Creating Error Worksheet
            Worksheet errorSheet = m_reportBook.Worksheets.Add() as Worksheet;
            errorSheet.Name = ErrorSheetName; ;

            // Initializes the Header to column mapping for error sheet
            m_errorHeaderToColumn = new Dictionary<string, int>();

            // Setting current row of operation in error sheet to one
            m_currentErrorRow = 1;

            // it will be used to decide column number for each field name. initaizling it with one.
            int column = 1;

            column = WriteFieldNames(sourceNameToFieldMapping, m_errorHeaderToColumn, errorSheet, m_currentErrorRow, column);

            WriteFieldName(m_errorHeaderToColumn, errorSheet, m_currentErrorRow, ref column, ErrorField);

            // Setting the Widths of Sheets's Columns
            SetDefaultWidths(errorSheet, column);

            // Skipping one line so that Failed workitems can be inserted after one blank line of the field name
            m_currentErrorRow += 2;
        }

        private int WriteFieldNames(IDictionary<string, IWorkItemField> sourceNameToFieldMapping,
                            Dictionary<string, int> headersToColumn,
                            Worksheet workSheet,
                            int row,
                            int column)
        {
            if (!string.IsNullOrEmpty(m_wizardInfo.DataSourceParser.StorageInfo.SourceIdFieldName))
            {
                WriteFieldName(headersToColumn, workSheet, row, ref column, m_wizardInfo.DataSourceParser.StorageInfo.SourceIdFieldName);
            }

            // Iterating throgh all field names and filling header to column mapping for work sheet
            // It also writes these field names in the sheet
            foreach (var kvp in sourceNameToFieldMapping)
            {
                WriteFieldName(headersToColumn, workSheet, row, ref column, kvp.Key);
            }

            if (!string.IsNullOrEmpty(m_wizardInfo.DataSourceParser.StorageInfo.TestSuiteFieldName))
            {
                WriteFieldName(headersToColumn, workSheet, row, ref column, m_wizardInfo.DataSourceParser.StorageInfo.TestSuiteFieldName);
            }

            if (m_wizardInfo.DataSourceParser.StorageInfo.LinkRules != null)
            {
                foreach (ILinkRule linkRule in m_wizardInfo.DataSourceParser.StorageInfo.LinkRules)
                {
                    WriteFieldName(headersToColumn, workSheet, row, ref column, linkRule.SourceFieldNameOfEndWorkItemCategory);
                }
            }

            return column;
        }


        private void WriteFieldName(Dictionary<string, int> headersToColumn,
                                    Worksheet workSheet,
                                    int row,
                                    ref int column,
                                    string fieldName)
        {
            if (!headersToColumn.ContainsKey(fieldName))
            {
                // Writes field name at (row,column)
                WriteValueAt(workSheet, row, column, fieldName);

                // Updates the header to column mapping
                headersToColumn.Add(fieldName, column);

                // Column number for next field will be +2 of the column number of current field
                column += 2;
            }
        }

        /// <summary>
        /// It writes the passed workitems in success worksheet
        /// </summary>
        /// <param name="passedWorkItem"></param>
        private void ReportSuccess(PassedSourceWorkItem passedWorkItem)
        {
            // Gets the success sheet
            Worksheet successSheet = m_reportBook.Worksheets[SuccessSheetName] as Worksheet;

            // Write ID Value
            WriteValueAt(successSheet, m_currentSuccessRow, m_successHeaderToColumn[IDField], passedWorkItem.TFSId.ToString(CultureInfo.CurrentCulture));

            // Write rest of Workitem
            m_currentSuccessRow = WriteXLWorkItem(successSheet, m_currentSuccessRow, m_successHeaderToColumn, passedWorkItem);
        }

        /// <summary>
        /// It writes the failed workitems in error worksheet
        /// </summary>
        /// <param name="failedDSWorkItem"></param>
        private void ReportError(FailedSourceWorkItem failedDSWorkItem)
        {
            // gets the error sheet
            Worksheet errorSheet = m_reportBook.Worksheets[ErrorSheetName] as Worksheet;

            // Write error in error column
            WriteValueAt(errorSheet, m_currentErrorRow, m_errorHeaderToColumn[ErrorField], failedDSWorkItem.Error);

            // Write rest of workitem
            m_currentErrorRow = WriteXLWorkItem(errorSheet, m_currentErrorRow, m_errorHeaderToColumn, failedDSWorkItem);
        }

        private void ReportWarning(WarningSourceWorkItem warnedWI)
        {
            // gets the warning sheet
            Worksheet warningSheet = m_reportBook.Worksheets[WarningSheetName] as Worksheet;

            // Write ID Value
            WriteValueAt(warningSheet, m_currentWarningRow, m_warningHeaderToColumn[IDField], warnedWI.TFSId.ToString(CultureInfo.CurrentCulture));

            // Write warning in warning column
            WriteValueAt(warningSheet, m_currentWarningRow, m_warningHeaderToColumn[WarningField], warnedWI.Warning);

            // Write rest of workitem
            m_currentWarningRow = WriteXLWorkItem(warningSheet, m_currentWarningRow, m_warningHeaderToColumn, warnedWI);
        }


        /// <summary>
        /// Writes XL Workitem to the provided worksheet at given row 
        /// </summary>
        /// <param name="sheet">Excel Worksheet at which workitem to be written</param>
        /// <param name="currentRow"> The Starting row of workitem</param>
        /// <param name="headerToColumn">Header to column mapping required for getting the column nuber of fields</param>
        /// <param name="xlWorkItem">Excel Workitem</param>
        /// 
        /// <returns>The Row number for next workitem</returns>
        private int WriteXLWorkItem(Worksheet sheet, int currentRow, Dictionary<string, int> headerToColumn, ISourceWorkItem xlWorkItem)
        {
            // Writing source id
            if (!string.IsNullOrEmpty(m_wizardInfo.DataSourceParser.StorageInfo.SourceIdFieldName))
            {
                WriteValueAt(sheet,
                             currentRow,
                             headerToColumn[m_wizardInfo.DataSourceParser.StorageInfo.SourceIdFieldName],
                             xlWorkItem.SourceId);
            }

            // Writing test suites
            if (!string.IsNullOrEmpty(m_wizardInfo.DataSourceParser.StorageInfo.TestSuiteFieldName) &&
                xlWorkItem.TestSuites != null)
            {
                StringBuilder testSuites = new StringBuilder();
                foreach (string testSuite in xlWorkItem.TestSuites)
                {
                    testSuites.Append(testSuite);
                    testSuites.Append(";");
                }
                WriteValueAt(sheet,
                             currentRow,
                             headerToColumn[m_wizardInfo.DataSourceParser.StorageInfo.TestSuiteFieldName],
                             testSuites.ToString());
            }

            // Writing link rules
            if (m_wizardInfo.DataSourceParser.StorageInfo.LinkRules != null)
            {
                foreach (ILinkRule linkRule in m_wizardInfo.DataSourceParser.StorageInfo.LinkRules)
                {
                    if (xlWorkItem.Links != null)
                    {
                        StringBuilder linkedWorkItemIds = new StringBuilder();
                        foreach (ILink link in xlWorkItem.Links)
                        {
                            if (string.CompareOrdinal(link.EndWorkItemCategory, linkRule.EndWorkItemCategory) == 0 &&
                                string.CompareOrdinal(link.LinkTypeName, linkRule.LinkTypeReferenceName) == 0)
                            {
                                linkedWorkItemIds.Append(link.EndWorkItemSourceId);
                                linkedWorkItemIds.Append(";");
                            }
                        }
                        WriteValueAt(sheet,
                                     currentRow,
                                     headerToColumn[linkRule.SourceFieldNameOfEndWorkItemCategory],
                                     linkedWorkItemIds.ToString());
                    }
                }
            }

            // The List of internal representation of Test Step
            List<SourceTestStep> steps = null;

            // Iteration through each field of Workitem and writing it to the excel sheet
            foreach (KeyValuePair<string, object> kvp in xlWorkItem.FieldValuePairs)
            {

                List<SourceTestStep> locSteps = kvp.Value as List<SourceTestStep>;

                // If value is a string then write it to the corresponding cell else it is a List of steps.
                if (locSteps != null)
                {
                    // Steps spreads into multiple rows so just save it right now.
                    steps = locSteps;
                }
                else
                {
                    WriteValueAt(sheet, currentRow, headerToColumn[kvp.Key], kvp.Value);
                }
            }

            // If We found steps then write steps in excel sheet
            if (steps != null)
            {
                // Iterating through each step and writing step title and expected result at consecutive lines
                foreach (SourceTestStep step in steps)
                {
                    if (!string.IsNullOrEmpty(m_stepTitleHeader) && headerToColumn.ContainsKey(m_stepTitleHeader))
                    {
                        WriteValueAt(sheet, currentRow, headerToColumn[m_stepTitleHeader], step.title);
                    }
                    if (!string.IsNullOrEmpty(m_stepExpectedResultHeader) && headerToColumn.ContainsKey(m_stepExpectedResultHeader))
                    {
                        WriteValueAt(sheet, currentRow, headerToColumn[m_stepExpectedResultHeader], step.expectedResult);
                    }
                    currentRow++;
                }
            }

            // returning row number for next workitem
            return currentRow + 2;
        }

        /// <summary>
        /// Writes value in sheet at (row,column) if value is not null or empty
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="value"></param>
        public static void WriteValueAt(Worksheet sheet, int row, int column, object value)
        {
            Range range = (Range)(sheet.Cells.get_Item(row, column));

            string stringValue = value as string;
            if (!string.IsNullOrEmpty(stringValue))
            {
                range.Value2 = stringValue.Trim();
                return;
            }
            else if (value != null)
            {
                range.Value2 = value.ToString();
            }
        }

        /// <summary>
        /// Sets the width of all data fields of a sheet of report to default size
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="endcolumn"></param>
        private void SetDefaultWidths(Worksheet sheet, int endcolumn)
        {
            for (int column = 1; column <= endcolumn; column += 2)
            {
                (sheet.Columns.get_Item(column) as Range).ColumnWidth = DefaultColumnWidth;
            }
        }

        #endregion

        #region IDisposibale implementation

        public void Dispose()
        {
            try
            {
                if (m_application != null)
                {
                    m_application.Workbooks.Close();
                    m_application.Quit();
                    m_application = null;
                }
            }
            catch (InvalidComObjectException)
            { }
            catch (COMException)
            { }
            finally
            {
                m_application = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        #endregion
    }
}