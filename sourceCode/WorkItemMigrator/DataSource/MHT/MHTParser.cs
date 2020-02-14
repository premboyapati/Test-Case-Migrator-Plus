//  Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Office.Interop.Word;

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    /// <summary>
    /// Class responsible for Parsing MHT and word Doc Files
    /// </summary>
    internal class MHTParser : IDataSourceParser
    {
        #region Nested definitions

        struct ImageDetails
        {
            public Range imageRange;
            public string filePath;
        }

        #endregion

        #region Fields

        // VSTO object for MHT Document
        private Document m_document;

        // Field Name to Work Item Fields Mapping
        private IDictionary<string, IWorkItemField> m_fieldNameToFields;

        // Work Item Attachments Counter
        private int m_workItemAttachmentCount = 0;

        private MHTStorageInfo m_storageInfo;

        private bool m_isProcessed = false;

        private bool m_isCopied = false;

        private static Application s_application;

        #endregion

        #region Constants

        public const string TestTitleDefaultTag = "_Test_Title_";

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="info"></param>
        public MHTParser(MHTStorageInfo info)
        {
            try
            {
                InitializeStorage(info);

                RawSourceWorkItems = new List<ISourceWorkItem>();
                ParsedSourceWorkItems = new List<ISourceWorkItem>();
                FieldNameToFields = new Dictionary<string, IWorkItemField>();
            }
            catch (COMException)
            {
                throw new WorkItemMigratorException("Word is not installed", null, null);
            }
        }

        #endregion

        #region Properties

        // Word Applictaion object
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
                    }
                    catch (COMException)
                    {
                        throw new WorkItemMigratorException("Initialization error",
                                                            "Error in initializing Word",
                                                            "Please verify that Microsoft Word is installed on your system");
                    }
                }
                return s_application;
            }
        }

        /// <summary>
        /// Needed for clipboard related operations
        /// </summary>
        public static SynchronizationContext STAThreadContext
        {
            get;
            set;
        }

        /// <summary>
        /// Field Name to Fiels Mapping
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
            }
        }

        public IList<ISourceWorkItem> RawSourceWorkItems
        {
            get;
            private set;
        }

        public IList<ISourceWorkItem> ParsedSourceWorkItems
        {
            get;
            private set;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Initializes the StorageInfo.FieldNames
        /// </summary>
        public void ParseDataSourceFieldNames()
        {
            m_storageInfo.PossibleFieldNames = new List<string>();
            StorageInfo.FieldNames = new List<SourceField>();
            // The Complete range of the MHT file
            Range documentRange = m_document.Range(Type.Missing, Type.Missing);

            foreach (Paragraph paragraph in documentRange.Paragraphs)
            {
                if (paragraph.Range.Tables.Count > 0 || paragraph.Range.InlineShapes.Count > 0)
                {
                    continue;
                }
                string field = CleanString(paragraph.Range.Text);
                if (!string.IsNullOrEmpty(field) && (paragraph.get_Style() as Style).NameLocal.Contains("Heading"))
                {
                    SourceField sourceField = new SourceField(field, false);
                    StorageInfo.FieldNames.Add(sourceField);
                }
                string styleName = (paragraph.get_Style() as Style).NameLocal;
                if (!string.IsNullOrEmpty(field) &&
                    (styleName.Contains("Heading") ||
                    styleName.Contains("Normal") ||
                    styleName.Contains("Comment")))
                {
                    m_storageInfo.PossibleFieldNames.Add(field);
                }
            }
        }

        /// <summary>
        /// parses the complete MHT source and returns the corresponding Source Workitem
        /// </summary>
        /// <returns></returns>
        public ISourceWorkItem GetNextWorkItem()
        {
            if (m_isProcessed)
            {
                return null;
            }
            else
            {
                m_isProcessed = true;
            }
            // The Source WorkItem which has to return
            ISourceWorkItem workItem = new SourceWorkItem();

            // The Complete range of the MHT file
            Range documentRange = m_document.Range(Type.Missing, Type.Missing);

            MHTSection mhtSection = MHTSection.Skip;

            string fieldName = string.Empty;
            bool skipMHTSectionCheck = false;
            if (m_storageInfo.IsFirstLineTitle)
            {
                m_storageInfo.TitleField = TestTitleDefaultTag;
                skipMHTSectionCheck = true;
                mhtSection = MHTSection.Title;
            }

            // Initializing paragraphNumber
            int paragraphNumber = 1;

            // String Builder to store the text parsed from mht
            StringBuilder textBuilder = new StringBuilder();
            var testSteps = new List<SourceTestStep>();

            // bool to check that whether line to parse is first line in the section or not
            bool isFirstLineInSection = true;

            // Loop to traverse the complete MHT Source
            while (paragraphNumber <= documentRange.Paragraphs.Count)
            {
                // The Current pargraph to parse 
                Paragraph paragraph = documentRange.Paragraphs[paragraphNumber];

                string paragraphText = CleanString(paragraph.Range.Text);

                // If this paragraph is empty line then continue to next paragraph
                if (string.IsNullOrEmpty(CleanString(paragraph.Range.Text)))
                {
                    paragraphNumber++;
                }
                // else parse the paragraph
                else
                {
                    bool isFieldName = false;
                    foreach (SourceField field in StorageInfo.FieldNames)
                    {
                        if (String.CompareOrdinal(field.FieldName, paragraphText) == 0)
                        {
                            isFieldName = true;
                            break;
                        }
                    }
                    if (!skipMHTSectionCheck && isFieldName)
                    {
                        MHTSection previousMHTSection = mhtSection;
                        if (String.CompareOrdinal(paragraphText, m_storageInfo.TitleField) == 0)
                        {
                            mhtSection = MHTSection.Title;
                        }
                        else if (String.CompareOrdinal(paragraphText, m_storageInfo.StepsField) == 0)
                        {

                            mhtSection = MHTSection.Steps;
                        }
                        else if (m_fieldNameToFields.ContainsKey(paragraphText))
                        {
                            if (m_fieldNameToFields[paragraphText].IsHtmlField)
                            {
                                mhtSection = MHTSection.History;
                            }
                            else
                            {
                                mhtSection = MHTSection.Default;
                            }
                        }
                        else
                        {
                            mhtSection = MHTSection.Skip;
                        }

                        isFirstLineInSection = true;
                        UpdateSourceWorkItem(previousMHTSection, workItem, fieldName, textBuilder, testSteps);
                        fieldName = paragraphText;
                    }

                    switch (mhtSection)
                    {
                        case MHTSection.Title:

                            // If the current paragraph has test Title Label then skip
                            if (String.CompareOrdinal(paragraphText, m_storageInfo.TitleField) == 0)
                            {
                                paragraphNumber++;
                            }
                            // else if we have not reached the end of Test Title Section
                            else
                            {
                                textBuilder.Append(CleanString(paragraph.Range.Text));
                                if (isFirstLineInSection)
                                {
                                    isFirstLineInSection = false;
                                }
                                paragraphNumber++;
                            }
                            skipMHTSectionCheck = false;

                            break;

                        case MHTSection.History:

                            if (isFirstLineInSection)
                            {
                                textBuilder.Append("<div>");
                                isFirstLineInSection = false;
                            }

                            if (paragraph.Range.Tables.Count > 0)
                            {
                                textBuilder.Append(GetHTMLTable(paragraph.Range.Tables[1]));
                                paragraphNumber = GetParagraphNumberAfterSpecifiedRange(documentRange, paragraph.Range.Tables[1].Range.End);
                            }
                            else if (paragraph.Range.InlineShapes.Count > 0)
                            {
                                string fileName = "Image" + (++m_workItemAttachmentCount) + ".bmp";
                                string filePath = Path.Combine(Path.GetTempPath(), fileName);
                                ExtractImage(paragraph.Range.InlineShapes[1].Range, filePath);
                                textBuilder.Append("<div><img src='" + filePath + "' /></div>");
                                paragraphNumber++;
                            }
                            else if (String.CompareOrdinal((paragraph.get_Style() as Style).NameLocal, "Heading 2") == 0)
                            {
                                textBuilder.Append("<h2>" + CleanString(paragraph.Range.Text).Replace("<", "&lt;").Replace(">", "&gt;") + "</h2>");
                                paragraphNumber++;
                            }
                            else
                            {
                                textBuilder.Append("<div>" + SecurityElement.Escape(CleanString(paragraph.Range.Text)) + "</div>");
                                paragraphNumber++;
                            }

                            break;

                        case MHTSection.Steps:
                            if (paragraph.Range.Tables.Count == 0)
                            {
                                paragraphNumber++;
                            }
                            else
                            {
                                Tables stepsTables = paragraph.Range.Tables;
                                foreach (Table table in stepsTables)
                                {
                                    int rowNumber = 0;
                                    foreach (Row currentRow in table.Rows)
                                    {
                                        if (!currentRow.IsFirst)
                                        {
                                            rowNumber++;

                                            string title = string.Empty;
                                            string expectedResult = string.Empty;
                                            List<string> attachments = new List<string>();
                                            int attachmentCount = 0;
                                            string attachmentPrefix = "Step" + rowNumber;

                                            if (currentRow.Cells.Count == 1)
                                            {
                                                title = ParseTestStepCell(currentRow.Cells[1], attachments, attachmentPrefix, ref attachmentCount, 1);
                                            }
                                            else if (currentRow.Cells.Count == 3)
                                            {
                                                title = ParseTestStepCell(currentRow.Cells[2], attachments, attachmentPrefix, ref attachmentCount, 1);
                                                expectedResult = ParseTestStepCell(currentRow.Cells[3], attachments, attachmentPrefix, ref attachmentCount, 1);
                                            }
                                            else
                                            {
                                                throw new WorkItemMigratorException("Steps table is not in correct format",
                                                                                    "Steps table is not having three columns.",
                                                                                    "Please make sure that you have mapped the correct field to 'Steps'.");
                                            }
                                            if (!string.IsNullOrEmpty(title) || !string.IsNullOrEmpty(expectedResult))
                                            {
                                                testSteps.Add(new SourceTestStep(title, expectedResult, attachments));
                                            }
                                        }
                                    }
                                }
                                paragraphNumber = GetParagraphNumberAfterSpecifiedRange(documentRange, paragraph.Range.Tables[1].Range.End);
                            }
                            break;

                        case MHTSection.Default:
                            if (isFirstLineInSection)
                            {
                                isFirstLineInSection = false;
                            }
                            else
                            {
                                textBuilder.Append(CleanString(paragraph.Range.Text));
                            }
                            paragraphNumber++;
                            break;

                        case MHTSection.Skip:
                            paragraphNumber++;
                            break;

                        default:
                            break;
                    }
                }
            }
            UpdateSourceWorkItem(mhtSection, workItem, fieldName, textBuilder, testSteps);

            if (m_storageInfo.IsFileNameTitle)
            {
                string title = Path.GetFileNameWithoutExtension(StorageInfo.Source);
                workItem.FieldValuePairs.Add(MHTParser.TestTitleDefaultTag, title);
            }
            workItem.SourcePath = StorageInfo.Source;
            RawSourceWorkItems.Add(workItem);
            ParsedSourceWorkItems.Add(workItem);

            return workItem;
        }

        public IList<string> GetUniqueValuesForField(string fieldName)
        {
            var values = new List<string>();

            if (!string.IsNullOrEmpty(fieldName) &&
                m_fieldNameToFields.ContainsKey(fieldName) &&
                ParsedSourceWorkItems[0].FieldValuePairs.ContainsKey(fieldName))
            {
                values.Add(ParsedSourceWorkItems[0].FieldValuePairs[fieldName].ToString());
            }
            return values;
        }

        public IDataStorageInfo StorageInfo
        {
            get
            {
                return m_storageInfo;
            }
            private set
            {
                m_storageInfo = value as MHTStorageInfo;
            }
        }

        public static void Preview(string sourcePath, IList<string> fields)
        {
            if (File.Exists(sourcePath))
            {
                string newPath = Path.GetTempFileName() + Path.GetExtension(sourcePath);
                try
                {
                    if (File.Exists(newPath))
                    {
                        File.Delete(newPath);
                    }
                    File.Copy(sourcePath, newPath);
                    File.SetAttributes(newPath, FileAttributes.Normal);
                    Document document = Application.Documents.Open(newPath);
                    Range docRange = document.Range(Type.Missing, Type.Missing);
                    docRange.HighlightColorIndex = WdColorIndex.wdGray25;
                    foreach (Paragraph paragraph in docRange.Paragraphs)
                    {
                        if (paragraph.Range.Tables.Count > 0 || paragraph.Range.InlineShapes.Count > 0)
                        {
                            continue;
                        }
                        string text = CleanString(paragraph.Range.Text);
                        if (string.IsNullOrEmpty(text))
                        {
                            paragraph.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                        }
                        else if (fields.Contains(text))
                        {
                            paragraph.Range.HighlightColorIndex = WdColorIndex.wdYellow;
                        }
                    }

                    document.Save();
                    document.Close();
                    document = null;
                    Process.Start(newPath);
                }
                catch (COMException)
                {
                    throw new WorkItemMigratorException("Unable to preview the MHt File", null, null);
                }
            }
        }

        #endregion

        #region Private Methods

        private void UpdateSourceWorkItem(MHTSection mhtSection, ISourceWorkItem workItem, string fieldName, StringBuilder textBuilder, List<SourceTestStep> testSteps)
        {
            if (string.IsNullOrEmpty(textBuilder.ToString()) &&
                (testSteps == null || testSteps.Count == 0))
            {
                return;
            }
            switch (mhtSection)
            {
                case MHTSection.Title:

                    string title = textBuilder.ToString();
                    Regex regEx1 = new Regex(@"\d+[a-z]*[A-Z]*[ ]*–[ ]*");
                    Regex regEx2 = new Regex(@"\d+[a-z]*[A-Z]*[ ]*-[ ]*");
                    if (regEx1.IsMatch(title))
                    {
                        title = title.Substring(title.IndexOf('–') + 1).Trim();
                    }
                    else if (regEx2.IsMatch(title))
                    {
                        title = title.Substring(title.IndexOf('-') + 1).Trim();
                    }
                    workItem.FieldValuePairs.Add(m_storageInfo.TitleField, title);
                    textBuilder.Clear();
                    break;

                case MHTSection.Steps:
                    if (!workItem.FieldValuePairs.ContainsKey(m_storageInfo.StepsField))
                    {
                        workItem.FieldValuePairs.Add(m_storageInfo.StepsField, new List<SourceTestStep>());
                    }
                    var workItemSteps = workItem.FieldValuePairs[m_storageInfo.StepsField] as List<SourceTestStep>;
                    foreach (var step in testSteps)
                    {
                        workItemSteps.Add(step);
                    }
                    testSteps.Clear();
                    break;

                case MHTSection.History:

                    textBuilder.Append("</div>");
                    if (workItem.FieldValuePairs.ContainsKey(fieldName))
                    {
                        workItem.FieldValuePairs[fieldName] += textBuilder.ToString();
                    }
                    else
                    {
                        workItem.FieldValuePairs.Add(fieldName, textBuilder.ToString());
                    }
                    textBuilder.Clear();
                    break;

                case MHTSection.Default:
                    if (workItem.FieldValuePairs.ContainsKey(fieldName))
                    {
                        workItem.FieldValuePairs[fieldName] += textBuilder.ToString();
                    }
                    else
                    {
                        workItem.FieldValuePairs.Add(fieldName, textBuilder.ToString());
                    }
                    textBuilder.Clear();
                    break;

                default:
                    break;
            }
        }

        /// <summary>
        /// Initializes the Storage Information
        /// </summary>
        /// <param name="info"></param>
        private void InitializeStorage(MHTStorageInfo info)
        {
            try
            {
                string newMHTFilePath = Path.GetTempFileName() + ".mht";
                File.Copy(Path.GetFullPath(info.Source), newMHTFilePath);
                m_document = Application.Documents.Open(newMHTFilePath);
                m_isCopied = true;
            }
            catch (COMException)
            {
                throw new FileFormatException();
            }
            Application.Visible = false;
            StorageInfo = info;
        }

        private static string CleanString(string s)
        {
            if (!string.IsNullOrEmpty(s))
            {
                s = s.Replace("\r", "").Replace("\a", "").Replace("\v", "").Trim();
            }
            return s;
        }

        private string ParseTestStepCell(Cell cell, List<string> attachments, string attachmentprefix, ref int attachmentCount, int nestingLevel)
        {
            StringBuilder textBuilder = new StringBuilder();
            int paragraphNumber = 1;
            while (paragraphNumber <= cell.Range.Paragraphs.Count)
            {
                Paragraph paragraph = cell.Range.Paragraphs[paragraphNumber];

                if (paragraph.Range.Tables.NestingLevel > nestingLevel)
                {
                    int level = 0;
                    Table nestedTable = paragraph.Range.Tables[1];
                    while (level < nestingLevel)
                    {
                        nestedTable = nestedTable.Range.Tables[1];
                        level++;
                    }

                    foreach (Row row in nestedTable.Rows)
                    {
                        StringBuilder rowString = new StringBuilder();
                        foreach (Cell innerCell in row.Cells)
                        {
                            rowString.Append(ParseTestStepCell(innerCell, attachments, attachmentprefix, ref attachmentCount, nestingLevel + 1));
                            rowString.Append("\t");
                        }
                        if (!string.IsNullOrEmpty(rowString.ToString().Trim()))
                        {
                            if (!textBuilder.ToString().EndsWith("\n", StringComparison.Ordinal))
                            {
                                textBuilder.Append("\n");
                            }
                            textBuilder.Append(rowString.ToString(0, rowString.Length - 1));
                        }
                    }
                    paragraphNumber = GetParagraphNumberAfterSpecifiedRange(cell.Range, nestedTable.Range.End);
                }
                else if (paragraph.Range.InlineShapes.Count > 0)
                {
                    if (!textBuilder.ToString().EndsWith("\n", StringComparison.Ordinal))
                    {
                        textBuilder.Append("\n");
                    }

                    object rangeStart = paragraph.Range.Start;
                    object rangeEnd = paragraph.Range.InlineShapes[1].Range.Start;

                    string text = CleanString(m_document.Range(ref rangeStart, ref rangeEnd).Text);
                    if (!string.IsNullOrEmpty(text))
                    {
                        textBuilder.Append(text);
                    }
                    int shapeNumber = 0;

                    while (shapeNumber < paragraph.Range.InlineShapes.Count)
                    {
                        shapeNumber++;
                        ++attachmentCount;
                        string fileName = attachmentprefix + "Attachment" + attachmentCount + ".bmp";
                        string filePath = Path.Combine(Path.GetTempPath(), fileName);

                        ExtractImage(paragraph.Range.InlineShapes[shapeNumber].Range, filePath);

                        textBuilder.Append("[");
                        textBuilder.Append(fileName);
                        textBuilder.Append("]");

                        attachments.Add(filePath);

                        rangeStart = paragraph.Range.InlineShapes[shapeNumber].Range.End;
                        if (shapeNumber < paragraph.Range.InlineShapes.Count)
                        {
                            rangeEnd = paragraph.Range.InlineShapes[shapeNumber + 1].Range.Start;
                        }
                        else
                        {
                            rangeEnd = paragraph.Range.End;
                        }
                        text = CleanString(m_document.Range(ref rangeStart, ref rangeEnd).Text);
                        if (!string.IsNullOrEmpty(text))
                        {
                            textBuilder.Append(text);
                        }
                    }
                    paragraphNumber++;
                }
                else
                {
                    string s = CleanString(paragraph.Range.Text);
                    if (!string.IsNullOrEmpty(s))
                    {
                        textBuilder.Append(s);
                        if (String.CompareOrdinal((paragraph.get_Style() as Style).NameLocal, "List Paragraph") == 0)
                        {
                            textBuilder.Append("\n");
                        }
                    }
                    if (!textBuilder.ToString().EndsWith("\n", StringComparison.Ordinal))
                    {
                        textBuilder.Append("\n");
                    }
                    paragraphNumber++;
                }
            }

            string t = textBuilder.ToString();
            if (t.EndsWith("\n", StringComparison.Ordinal))
            {
                t = t.Substring(0, t.Length - 1);
            }
            if (t.StartsWith("\n", StringComparison.Ordinal))
            {
                t = t.Substring(1);
            }
            return t;
        }

        private void ExtractImage(Range imageRange, string filePath)
        {
            ImageDetails imageInfo = new ImageDetails();
            imageInfo.imageRange = imageRange;
            imageInfo.filePath = filePath;

            if (STAThreadContext != null)
            {
                STAThreadContext.Send(ExtractImageInSTAThreadContext, imageInfo);
            }
            else
            {
                ExtractImageInSTAThreadContext(imageInfo);
            }
        }

        private void ExtractImageInSTAThreadContext(object imageDetails)
        {
            ImageDetails imageInfo = (ImageDetails)imageDetails;
            Range imageRange = imageInfo.imageRange;
            string filePath = imageInfo.filePath;
            try
            {
                imageRange.Copy();

                System.Windows.IDataObject data = System.Windows.Clipboard.GetDataObject();
                using (Bitmap bitmap = (Bitmap)data.GetData(typeof(System.Drawing.Bitmap)))
                {
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                    bitmap.Save(filePath);
                }
                System.Windows.Clipboard.Clear();
            }
            catch (OutOfMemoryException)
            {
                // try again
                System.Windows.Clipboard.Clear();
                ExtractImage(imageRange, filePath);
            }
        }

        private int GetParagraphNumberAfterSpecifiedRange(Range wordRange, int range)
        {
            int paragraphNumber = 0;
            foreach (Paragraph p in wordRange.Paragraphs)
            {
                paragraphNumber++;
                if (p.Range.Start >= range)
                {
                    break;
                }
            }
            return paragraphNumber;
        }

        private string GetHTMLTable(Table table)
        {
            StringBuilder htmlBuilder = new StringBuilder();
            htmlBuilder.Append("<table border='1'>");
            foreach (Row row in table.Rows)
            {
                htmlBuilder.Append("<tr>");
                foreach (Cell cell in row.Cells)
                {
                    htmlBuilder.Append("<td>");

                    foreach (Paragraph paragraph in cell.Range.Paragraphs)
                    {
                        if (paragraph.Range.InlineShapes.Count > 0)
                        {
                            string fileName = "Image" + (++m_workItemAttachmentCount) + ".bmp";
                            string filePath = Path.Combine(Path.GetTempPath(), fileName);
                            ExtractImage(paragraph.Range.InlineShapes[1].Range, filePath);
                            htmlBuilder.Append("<div><img src='" + filePath + "' /></div>");
                        }
                        else
                        {
                            string cleanString = SecurityElement.Escape(CleanString(paragraph.Range.Text));
                            if (!string.IsNullOrEmpty(cleanString))
                            {
                                htmlBuilder.Append("<div>" + cleanString + "</div>");
                            }
                        }
                    }
                    htmlBuilder.Append("</td>");
                }
                htmlBuilder.Append("</tr>");
            }
            htmlBuilder.Append("</table>");

            return htmlBuilder.ToString();
        }

        #endregion

        #region IDisposable Implementation


        public static void Quit()
        {
            try
            {
                if (s_application != null)
                {
                    try
                    {
                        s_application.Documents.Close();
                        s_application.Visible = false;
                    }
                    catch (COMException)
                    { }

                    s_application.Quit(Type.Missing, Type.Missing, Type.Missing);
                }
            }
            catch (COMException)
            { }
            finally
            {
                s_application = null;
                GC.WaitForPendingFinalizers();
            }

        }
        public void Dispose()
        {
            string path = null;
            try
            {
                if (m_document != null)
                {
                    if (File.Exists(m_document.FullName))
                    {
                        path = m_document.FullName;
                    }
                    m_document.Close(false);
                }
            }
            catch (COMException)
            { }
            catch (InvalidCastException)
            { }
            finally
            {

                m_document = null;

                GC.WaitForPendingFinalizers();
                try
                {
                    if (m_isCopied && File.Exists(path))
                    {
                        File.Delete(path);
                    }
                }
                catch (UnauthorizedAccessException)
                { }
                catch (IOException)
                { }
            }
        }

        #endregion
    }
}
