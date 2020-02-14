using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace Microsoft.VisualStudio.TestTools.TestCaseMigrator
{
    internal class MHTDataSource : IDataSourceParser
    {
        #region Fields

        // Word Applictaion object
        private Application m_application;

        // VSTO object for MHT Document
        private Document m_document;

        // Field Name to Work Item Fields Mapping
        private IDictionary<string, IWorkItemField> m_fieldNameToFields;

        // Work Item Attachments Counter
        private int m_workItemAttachmentCount = 0;
        
        #endregion

        #region Constructor_Initialization
        
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="info"></param>
        public MHTDataSource(MHTStorageInfo info)
        {
            m_application = new Application();
            m_application.Visible = false;

            InitializeStorage(info);
        }

        /// <summary>
        /// Initializes the Storage Information
        /// </summary>
        /// <param name="info"></param>
        private void InitializeStorage(MHTStorageInfo info)
        {
            try
            {
                string fullPath = Path.GetFullPath(info.Source);
                m_document = m_application.Documents.Open(fullPath);
            }
            catch (COMException)
            {
                throw new FileFormatException();
            }
            StorageInfo = info;
        }

        public MHTDataSource(Application application, MHTStorageInfo info)
        {
            m_application = application;
            InitializeStorage(info);
        }

        #endregion

        public void ParseDataSourceFieldNames()
        {
            StorageInfo.FieldNames = new List<string> { "Test Title", "Test Details", "Test Steps" };
        }


        public IDictionary<string, IWorkItemField> FieldNameToFields
        {
            set
            {
                m_fieldNameToFields = value;
            }
        }

        private string CleanString(string s, bool addSpaceAtStart)
        {
            s = s.Replace("\r", "").Replace("\a", "").Trim();
            if (addSpaceAtStart && s != string.Empty)
            {
                s = " " + s;
            }
            return s;
        }

        public SourceWorkItem GetNextWorkItem()
        {
            SourceWorkItem workItem = new SourceWorkItem();
            Range documentRange = m_document.Range(Type.Missing, Type.Missing);
            MHTSections mhtSection = MHTSections.Title;
            int paragraphNumber = 1;
            while (paragraphNumber <= documentRange.Paragraphs.Count)
            {
                Paragraph paragraph = documentRange.Paragraphs[paragraphNumber];
                if (string.IsNullOrEmpty(CleanString(paragraph.Range.Text, false)))
                {
                    paragraphNumber++;
                    continue;
                }

                if (!string.IsNullOrEmpty(paragraph.Range.Text))
                {
                    switch (mhtSection)
                    {
                        case MHTSections.Title:

                            if (paragraph.Range.Text.Contains("Test Title"))
                            {
                                paragraphNumber++;
                                paragraphNumber = GetTitle(workItem, documentRange, paragraphNumber);
                                mhtSection = MHTSections.Details;
                            }
                            else
                            {
                                paragraphNumber = GetTitle(workItem, documentRange, paragraphNumber);
                                mhtSection = MHTSections.Details;
                            }
                            break;

                        case MHTSections.Details:
                            if (paragraph.Range.Text.Contains("Test Details"))
                            {
                                string details = GetDetails(documentRange, ref paragraphNumber);
                                workItem.workItem.Add("Test Details", details);
                                mhtSection = MHTSections.Steps;
                            }
                            break;

                        case MHTSections.Steps:
                            if (paragraph.Range.Text.Contains("Test Steps"))
                            {
                                paragraphNumber++;
                                paragraph = documentRange.Paragraphs[paragraphNumber];
                                
                                if (paragraph.Range.Tables.Count == 0)
                                {
                                    paragraphNumber++;
                                    paragraph = documentRange.Paragraphs[paragraphNumber];
                                }

                                Tables stepsTables = paragraph.Range.Tables;
                                var testSteps = new List<SourceTestStep>();
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
                                                title = ParseCell(currentRow.Cells[1], attachments, attachmentPrefix, ref attachmentCount);
                                            }
                                            else
                                            {
                                                title = ParseCell(currentRow.Cells[2], attachments, attachmentPrefix, ref attachmentCount);
                                                expectedResult = ParseCell(currentRow.Cells[3], attachments, attachmentPrefix, ref attachmentCount);
                                            }
                                            if (!string.IsNullOrEmpty(title) || !string.IsNullOrEmpty(expectedResult))
                                            {
                                                testSteps.Add(new SourceTestStep(title, expectedResult, attachments));
                                            }
                                        }
                                    }
                                }
                                workItem.workItem.Add("Test Steps", testSteps);
                                paragraphNumber = GetParagraphNumberAfterSpecifiedRange(documentRange, paragraph.Range.Tables[1].Range.End);
                                mhtSection = MHTSections.RevisionHistory;
                            }
                            break;

                        case MHTSections.RevisionHistory:

                            string revisionHistory = GetDetails(documentRange, ref paragraphNumber);
                            workItem.workItem["Test Details"] += revisionHistory;
                            break;

                        default:
                            break;
                    }
                }
            }
            return workItem;
        }

        private string ParseCell(Cell cell, List<string> attachments, string attachmentprefix, ref int attachmentCount)
        {
            string text = string.Empty;
            int paragraphNumber = 1;
            while(paragraphNumber <= cell.Range.Paragraphs.Count)
            {
                Paragraph p = cell.Range.Paragraphs[paragraphNumber];

                if (p.Range.Tables.NestingLevel > 1)
                {
                    Table nestedTable = p.Range.Tables[1].Range.Tables[1];
                    foreach (Row row in nestedTable.Rows)
                    {
                        string rowString = string.Empty;
                        for (int cellNo = 1; cellNo <= row.Cells.Count; cellNo++)
                        {
                            rowString += CleanString(row.Cells[cellNo].Range.Text, false) + "\t";
                        }
                        if (rowString.Trim() != string.Empty)
                        {
                            if (!text.EndsWith("\n"))
                            {
                                text += "\n";
                            }
                            text += rowString.Substring(0, rowString.Length - 1);
                        }
                    }
                    paragraphNumber = GetParagraphNumberAfterSpecifiedRange(cell.Range, nestedTable.Range.End);
                }
                else if (p.Range.InlineShapes.Count > 0)
                {
                    ++attachmentCount;
                    string fileName = attachmentprefix + "Attachment" + attachmentCount + ".bmp";
                    string filePath = Path.Combine(Path.GetTempPath(), fileName);

                    ExtractImage(p.Range.InlineShapes[1].Range, filePath);

                    text += (text.EndsWith("\n") ? "[" : "\n[") + fileName + "]\n";
                    attachments.Add(filePath);
                    paragraphNumber++;
                }
                else
                {
                    string s = CleanString(p.Range.Text, false);
                    if (!string.IsNullOrEmpty(s))
                    {
                        text += s;
                        if ((p.get_Style() as Style).NameLocal == "List Paragraph")
                        {
                            text += "\n";
                        }
                    }
                    if(!text.EndsWith("\n"))
                    {
                        text += "\n";
                    }
                    paragraphNumber++;
                }
            }
            if (text.EndsWith("\n"))
            {
                text = text.Substring(0, text.Length - 1);
            }
            if (text.StartsWith("\n"))
            {
                text = text.Substring(1);
            }
            return text;
        }

        private void ExtractImage(Range imageRange, string filePath)
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

        private string GetDetails(Range documentRange, ref int paragraphNumber)
        {
            Paragraph paragraph = documentRange.Paragraphs[paragraphNumber];

            string details = "<div>";
            while (paragraphNumber <= documentRange.Paragraphs.Count &&
                   !documentRange.Paragraphs[paragraphNumber].Range.Text.Contains("Test Steps"))
            {
                paragraph = documentRange.Paragraphs[paragraphNumber];
                if (paragraph.Range.Tables.Count > 0)
                {
                    details += GetHTMLTable(paragraph.Range.Tables[1]);
                    paragraphNumber = GetParagraphNumberAfterSpecifiedRange(documentRange, paragraph.Range.Tables[1].Range.End);
                }
                else if (paragraph.Range.InlineShapes.Count > 0)
                {
                    string fileName = "Image" + (++m_workItemAttachmentCount) + ".bmp";
                    string filePath = Path.Combine(Path.GetTempPath(), fileName);
                    ExtractImage(paragraph.Range.InlineShapes[1].Range, filePath);
                    details += "<div><img src='" + filePath + "' /></div>";
                    paragraphNumber++;
                }
                else if ((paragraph.get_Style() as Style).NameLocal == "Heading 2")
                {
                    details += "<h2>" + CleanString(paragraph.Range.Text, false).Replace("<", "&lt;").Replace(">", "&gt;") + "</h2>";
                    paragraphNumber++;
                }
                else if ((paragraph.get_Style() as Style).NameLocal == "Comment" ||
                         (paragraph.get_Style() as Style).NameLocal == "Normal" ||
                         (paragraph.get_Style() as Style).NameLocal == "List Paragraph")
                {
                    details += "<div>" + CleanString(paragraph.Range.Text, false).Replace("<", "&lt;").Replace(">", "&gt;") + "</div>";
                    paragraphNumber++;
                }
                else
                {
                    paragraphNumber++;
                    throw new Exception();
                }
            }
            details += "</div>";
            return details;
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
            string htmlTable = "<table border='1'>";
            foreach (Row row in table.Rows)
            {
                htmlTable += "<tr>";
                foreach (Cell cell in row.Cells)
                {
                    htmlTable += "<td>";
                    string[] lines = cell.Range.Text.Split('\r');
                    foreach (string line in lines)
                    {
                        htmlTable += "<div>" + CleanString(line, false).Replace("<", "&lt;").Replace(">", "&gt;") + "</div>";
                    }
                    htmlTable += "</td>";
                }
                htmlTable += "</tr>";
            }
            htmlTable += "</table>";

            return htmlTable;
        }

        private int GetTitle(SourceWorkItem workItem, Range documentRange, int paragraphNumber)
        {
            string title = CleanString(documentRange.Paragraphs[paragraphNumber].Range.Text, false);

            paragraphNumber++;

            while (!documentRange.Paragraphs[paragraphNumber].Range.Text.Contains("Test Details"))
            {
                title += CleanString(documentRange.Paragraphs[paragraphNumber].Range.Text, true);
            }

            Regex regEx1 = new Regex(@"\d+[ ]*–[ ]*");
            Regex regEx2 = new Regex(@"\d+[ ]*-[ ]*");
            if (regEx1.IsMatch(title))
            {
                title = title.Substring(title.IndexOf('–') + 1).Trim();
            }
            else if (regEx2.IsMatch(title))
            {
                title = title.Substring(title.IndexOf('-') + 1).Trim();
            }
            workItem.workItem.Add("Test Title", title);

            return paragraphNumber;
        }


        public IList<string> GetUniqueValuesForField(string fieldName)
        {
            return new List<string>();
        }

        public IDataStorageInfo StorageInfo
        {
            get;
            private set;
        }

        public void Dispose()
        {
            m_application.Documents.Close();
            m_document = null;

            m_application.Visible = false;

            m_application.Quit(Type.Missing, Type.Missing, Type.Missing);
            m_application = null;

            GC.WaitForPendingFinalizers();
        }
    }

    internal enum MHTSections
    {
        Title,
        Details,
        Steps,
        RevisionHistory
    }
}
