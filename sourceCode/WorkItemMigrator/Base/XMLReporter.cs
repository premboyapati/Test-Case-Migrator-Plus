//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;
    using System.Xml;

    /// <summary>
    /// Generates XML report with XSl Stylesheet
    /// </summary>
    internal class XMLReporter : NotifyPropertyChange, IReporter
    {
        #region Constants

        public const string RootNodeName = "TestCaseMigratorPlus";
        public const string SummaryNodeName = "Summary";
        public const string PassedWorkItemsNodeName = "PassedWorkItems";
        public const string WarningWorkItemsNodeName = "WarningWorkItems";
        public const string FailedWorkItemsNodeName = "FailedWorkItems";
        public const string FileAttributeName = "File";
        public const string CountAttributeName = "Count";
        public const string WorkitemsNodeName = "WorkItems";
        public const string PassedWorkItemNodeName = "PassedWorkItem";
        public const string WarningWorkItemNodeName = "WarningWorkItem";
        public const string FailedWorkItemNodeName = "FailedWorkItem";
        public const string WarningAttributeName = "Warning";
        public const string ErrorAttributeName = "Error";
        public const string IdAttributeName = "TFSId";
        public const string SourceAttributeName = "Source";
        public const string CommandUsedNodeName = "CommandLine";

        public const string ListOfPassedWorkItemsFileName = "Passed.txt";
        public const string ListOfWarningWorkItemsFileName = "Warning.txt";
        public const string ListOfFailedWorkItemsFileName = "Failed.txt";

        public const string XSLFileName = "Transform.xsl";

        #endregion

        #region Fields

        private XmlDocument m_document;
        private XmlElement m_summaryNode;
        private XmlElement m_passedWorkItemsNode;
        private XmlElement m_warningWorkItemsNode;
        private XmlElement m_failedWorkItemsNode;
        private int m_passedWorkItemsCount;
        private int m_warningWorkItemsCount;
        private int m_failedWorkItemsCount;
        private string m_reportFile;
        private IList<string> m_passedWorkItemSourceFiles;
        private IList<string> m_warningWorkItemSourceFiles;
        private IList<string> m_failedWorkItemSourceFiles;
        private WizardInfo m_wizardInfo;

        #endregion

        #region Constructor

        public XMLReporter(WizardInfo wizardInfo)
        {
            m_wizardInfo = wizardInfo;

            m_document = new XmlDocument();
            var rootNode = m_document.CreateElement(RootNodeName);
            XmlDeclaration dec = m_document.CreateXmlDeclaration("1.0", null, null);
            m_document.AppendChild(dec);
            String pi = "type='text/xsl' href='" + XSLFileName + "'";
            var ins = m_document.CreateProcessingInstruction("xml-stylesheet", pi);
            m_document.AppendChild(ins);

            m_summaryNode = m_document.CreateElement(SummaryNodeName);
            rootNode.AppendChild(m_summaryNode);

            var workItemsNode = m_document.CreateElement(WorkitemsNodeName);

            m_passedWorkItemsNode = m_document.CreateElement(PassedWorkItemsNodeName);
            workItemsNode.AppendChild(m_passedWorkItemsNode);

            m_warningWorkItemsNode = m_document.CreateElement(WarningWorkItemsNodeName);
            workItemsNode.AppendChild(m_warningWorkItemsNode);

            m_failedWorkItemsNode = m_document.CreateElement(FailedWorkItemsNodeName);
            workItemsNode.AppendChild(m_failedWorkItemsNode);

            rootNode.AppendChild(workItemsNode);

            m_document.AppendChild(rootNode);

            m_passedWorkItemSourceFiles = new List<string>();
            m_warningWorkItemSourceFiles = new List<string>();
            m_failedWorkItemSourceFiles = new List<string>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Complete file Path of the Report File
        /// </summary>
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

        /// <summary>
        /// Add SourceWorkItem Entry in the report
        /// </summary>
        /// <param name="sourceWorkItem"></param>
        public void AddEntry(ISourceWorkItem sourceWorkItem)
        {
            var warningSourceWorkItem = sourceWorkItem as WarningSourceWorkItem;
            if (warningSourceWorkItem != null)
            {
                m_warningWorkItemsCount++;
                var warningWorkItemNode = m_document.CreateElement(WarningWorkItemNodeName);
                Utilities.AppendAttribute(warningWorkItemNode, SourceAttributeName, sourceWorkItem.SourcePath);
                Utilities.AppendAttribute(warningWorkItemNode, IdAttributeName, warningSourceWorkItem.TFSId.ToString(System.Globalization.CultureInfo.CurrentCulture));
                Utilities.AppendAttribute(warningWorkItemNode, WarningAttributeName, warningSourceWorkItem.Warning);
                m_warningWorkItemsNode.AppendChild(warningWorkItemNode);
                m_warningWorkItemSourceFiles.Add(sourceWorkItem.SourcePath);
            }
            else
            {
                var passedSourceWorkItem = sourceWorkItem as PassedSourceWorkItem;
                if (passedSourceWorkItem != null)
                {
                    m_passedWorkItemsCount++;
                    var passedWorkItemNode = m_document.CreateElement(PassedWorkItemNodeName);
                    Utilities.AppendAttribute(passedWorkItemNode, SourceAttributeName, sourceWorkItem.SourcePath);
                    Utilities.AppendAttribute(passedWorkItemNode, IdAttributeName, passedSourceWorkItem.TFSId.ToString(System.Globalization.CultureInfo.CurrentCulture));
                    m_passedWorkItemsNode.AppendChild(passedWorkItemNode);
                    m_passedWorkItemSourceFiles.Add(sourceWorkItem.SourcePath);
                }
                else
                {
                    var failedSourceWorkItem = sourceWorkItem as FailedSourceWorkItem;
                    if (failedSourceWorkItem != null)
                    {
                        m_failedWorkItemsCount++;
                        var failedWorkItemNode = m_document.CreateElement(FailedWorkItemNodeName);
                        Utilities.AppendAttribute(failedWorkItemNode, SourceAttributeName, sourceWorkItem.SourcePath);
                        Utilities.AppendAttribute(failedWorkItemNode, ErrorAttributeName, failedSourceWorkItem.Error);
                        m_failedWorkItemsNode.AppendChild(failedWorkItemNode);
                        m_failedWorkItemSourceFiles.Add(sourceWorkItem.SourcePath);
                    }
                    else
                    {
                        throw new WorkItemMigratorException("Incorrect Source workItem Type", null, null);
                    }
                }
            }
        }

        /// <summary>
        /// Publishes the report with the style sheet
        /// </summary>
        public void Publish()
        {
            string directory = Path.GetDirectoryName(ReportFile);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            m_wizardInfo.SaveSettings(Path.Combine(directory, m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName + "-settings.xml"));

            var listOfPassedWorkItemsFilePath = Path.Combine(directory, ListOfPassedWorkItemsFileName);
            var listOfWarningWorkItemsFilePath = Path.Combine(directory, ListOfWarningWorkItemsFileName);
            var listOfFailedWorkItemsFilePath = Path.Combine(directory, ListOfFailedWorkItemsFileName);

            SaveList(listOfPassedWorkItemsFilePath, m_passedWorkItemSourceFiles);
            SaveList(listOfWarningWorkItemsFilePath, m_warningWorkItemSourceFiles);
            SaveList(listOfFailedWorkItemsFilePath, m_failedWorkItemSourceFiles);

            var passedWorkItemsNode = m_document.CreateElement(PassedWorkItemsNodeName);
            Utilities.AppendAttribute(passedWorkItemsNode, FileAttributeName, listOfPassedWorkItemsFilePath);
            Utilities.AppendAttribute(passedWorkItemsNode, CountAttributeName, m_passedWorkItemsCount.ToString(System.Globalization.CultureInfo.CurrentCulture));
            m_summaryNode.AppendChild(passedWorkItemsNode);

            var warningWorkItemsNode = m_document.CreateElement(WarningWorkItemsNodeName);
            Utilities.AppendAttribute(warningWorkItemsNode, FileAttributeName, listOfWarningWorkItemsFilePath);
            Utilities.AppendAttribute(warningWorkItemsNode, CountAttributeName, m_warningWorkItemsCount.ToString(System.Globalization.CultureInfo.CurrentCulture));
            m_summaryNode.AppendChild(warningWorkItemsNode);

            var failedWorkItemsNode = m_document.CreateElement(FailedWorkItemsNodeName);
            Utilities.AppendAttribute(failedWorkItemsNode, FileAttributeName, listOfFailedWorkItemsFilePath);
            Utilities.AppendAttribute(failedWorkItemsNode, CountAttributeName, m_failedWorkItemsCount.ToString(System.Globalization.CultureInfo.CurrentCulture));
            m_summaryNode.AppendChild(failedWorkItemsNode);

            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                string fullProcessPath = assembly.Location;
                string processDir = Path.GetDirectoryName(fullProcessPath);
                string XSLFilePath = Path.Combine(processDir, XSLFileName);

                if (File.Exists(XSLFilePath))
                {
                    string destination = Path.Combine(directory, XSLFileName);
                    if (File.Exists(destination))
                    {
                        File.Delete(destination);
                    }
                    File.Copy(XSLFilePath, destination);
                }
            }
            catch (IOException)
            { }
            catch (UnauthorizedAccessException)
            { }

            if (!string.IsNullOrEmpty(m_wizardInfo.CommandUsed))
            {
                var commandUsedNode = m_document.CreateElement(CommandUsedNodeName);
                commandUsedNode.InnerXml = m_wizardInfo.CommandUsed;
                m_document.LastChild.AppendChild(commandUsedNode);
            }
            m_document.Save(ReportFile);
        }

        #endregion

        #region private methods

        private void SaveList(string filePath, IList<string> list)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                using (StreamWriter tw = new StreamWriter(filePath))
                {
                    foreach (string s in list)
                    {
                        tw.WriteLine(s);
                    }
                }
            }
            catch (IOException)
            {
                // TODO: Show error
            }
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose()
        { }

        #endregion

    }
}
