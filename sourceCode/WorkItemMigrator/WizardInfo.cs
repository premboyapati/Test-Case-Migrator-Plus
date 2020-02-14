//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.IO;
    using System.Xml;

    /// <summary>
    /// This contains the Complete Information needed acrosss the Wizard Session.
    /// It also supports the Loading/Saving of Settings File.
    /// </summary>
    internal class WizardInfo : NotifyPropertyChange, IDisposable
    {
        #region Fields

        private DataSourceType m_dataSourceType;
        private string m_mhtSource;
        private string m_outputSettingsFilePath;
        private string m_inputSettingsFilePath;
        private ProgressWindow m_progressPart;
        private bool m_canConfirm;
        private IWorkItemGenerator m_workItemGenerator;
        private bool m_isLinking;

        #endregion

        #region Constants
        /// <summary>
        /// These constants are name of XML Nodes/Attributes of Settings File. These needed for Saving/Loading Settings File.
        /// </summary>
        private const string c_rootNodeTag = "TestCaseMigratorPlus";
        private const string c_fieldsRootNodeTag = "Fields";
        private const string c_sampleFileAttributeTag = "SampleFIle";
        private const string c_fieldNodeTag = "Field";
        private const string c_stepsFieldNodeTag = "StepsField";
        private const string c_canDeleteFieldNameAttributeTag = "CanDelete";
        private const string c_nameAttributeTag = "Name";
        private const string c_fieldMappingRootNodeTag = "FieldMapping";
        private const string c_fieldMappingNodeTag = "FieldMappingRow";
        private const string c_dataMappingRootNodeTag = "DataMapping";
        private const string c_dataMappingNodeTag = "DataMappingRow";
        private const string c_dataSourceFieldNameAttributeTag = "DataSourceField";
        private const string c_wiFieldNameAttributeTag = "WorkItemField";
        private const string c_dataSourceValueAttributeTag = "PreviousValue";
        private const string c_newValueAttributeTag = "NewValue";
        private const string c_isFileNameTitleAttributeTag = "IsFileNameTitle";
        private const string c_isFirstLineTitleAttributeTag = "IsFirstLineTitle";
        private const string c_multiLineSenseNode = "MultiLineSense";
        private const string c_parameterizationNode = "Parameterization";
        private const string c_createAreaIterationPathAttribute = "CreateAreaIterationPath";
        private const string c_startParameterizationAttribute = "Start";
        private const string c_endParameterizationAttribute = "End";
        private const string c_relationshipsNodeName = "Relationships";
        private const string c_isLinkingAttributeName = "IsLinking";
        private const string c_descriptionAttributeName = "Description";
        private const string c_sourceIdFieldNodeName = "SourceIdField";
        private const string c_testSuiteFieldNodeName = "TestSuiteField";
        private const string c_linkTypesRootNodeName = "LinkTypes";
        private const string c_linkTypeNodeName = "LinkType";
        private const string c_linkTypeNameAttributeName = "LinkTypeName";
        private const string c_linkSourceFieldAttributeName = "SourceField";
        private const string c_linkedWorkItemAttributeName = "LinkedWorkItemType";
        private const string LTChar = "<";
        private const string GTChar = ">";
        private const string LTXMLChar = "&lt;";
        private const string GTXMLChar = "&gt;";
        private const string XMLDeclarationVersion = "1.0";

        #endregion

        #region Properties

        public bool CanConfirm
        {
            get
            {
                return m_canConfirm;
            }
            set
            {
                m_canConfirm = value;
                NotifyPropertyChanged("CanConfirm");
            }
        }

        public bool IsLinking
        {
            get
            {
                if (DataSourceType == DataSourceType.MHT)
                {
                    m_isLinking = false;
                }
                return m_isLinking;
            }
            set
            {
                m_isLinking = value;
            }
        }

        /// <summary>
        /// The Data Source Type which is going to be parsed.
        /// </summary>
        public DataSourceType DataSourceType
        {
            get
            {
                return m_dataSourceType;
            }
            set
            {
                m_dataSourceType = value;
                NotifyPropertyChanged("DataSourceType");
            }
        }

        public string MHTSource
        {
            get
            {
                return m_mhtSource;
            }
            set
            {
                m_mhtSource = value;
                NotifyPropertyChanged("MHTSource");
            }
        }

        /// <summary>
        /// The Sample Data Source used for getting the field names and field Mapping
        /// </summary>
        public IDataSourceParser DataSourceParser
        {
            get;
            set;
        }

        public IList<IDataStorageInfo> DataStorageInfos
        {
            get;
            set;
        }

        public IWorkItemGenerator WorkItemGenerator
        {
            get
            {
                return m_workItemGenerator;
            }
            set
            {
                m_workItemGenerator = value;
                NotifyPropertyChanged("WorkItemGenerator");
            }
        }

        public IMigrator Migrator
        {
            get;
            private set;
        }

        public IReporter Reporter
        {
            get;
            set;
        }

        public IList<ISourceWorkItem> ResultWorkItems
        {
            get;
            set;
        }

        /// <summary>
        /// File Path of Setting File Path from which settings are going to be imported
        /// </summary>
        public string InputSettingsFilePath
        {
            get
            {
                return m_inputSettingsFilePath;
            }
            set
            {
                m_inputSettingsFilePath = value;
                NotifyPropertyChanged("InputSettingsFilePath");
            }
        }

        /// <summary>
        /// File Path of Settings File Path where current Settings are going to be saved.
        /// </summary>
        public string OutputSettingsFilePath
        {
            get
            {
                return m_outputSettingsFilePath;
            }
            set
            {
                m_outputSettingsFilePath = value;
                NotifyPropertyChanged("OutputSettingsFilePath");
            }
        }

        /// <summary>
        /// Corresponding Command Line 
        /// </summary>
        public string CommandUsed
        {
            get;
            set;
        }

        /// <summary>
        /// Overlay part to show progress for long waiting operations
        /// </summary>
        public ProgressWindow ProgressPart
        {
            get
            {
                return m_progressPart;
            }
            set
            {
                if (value == null && m_progressPart != null)
                {
                    m_progressPart.PropertyChanged -= ProgressPart_PropertyChanged;
                    App.CallMethodInUISynchronizationContext(CloseProgressPart, null);
                }
                m_progressPart = value;
                NotifyPropertyChanged("ProgressPart");
            }
        }

        /// <summary>
        /// Name of Sample File used to get field names
        /// </summary>
        public String SampleFileForFields
        {
            get;
            set;
        }

        public IRelationshipsInfo RelationshipsInfo
        {
            get;
            private set;
        }

        public ILinksManager LinksManager
        {
            get;
            set;
        }

        #endregion

        #region Constructor

        public WizardInfo()
        {
            Migrator = new Migrator();
            IsLinking = true;
            RelationshipsInfo = new RelationshipsInfo();
            CanConfirm = false;
        }

        #endregion

        #region public methods

        /// <summary>
        /// LoadSettings from Input Settings File Path
        /// </summary>
        public void LoadSettings(string inputSettingsFile)
        {
            InitializeProgressView();
            ProgressPart.Header = "Loading settings...";

            // Load XML Document from InputSettingsFilePath
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(inputSettingsFile);

            // get the root Node of the xml Document
            XmlElement root = (XmlElement)xmlDocument.LastChild;

            // Traverse each child node of the root node
            foreach (XmlElement node in root.ChildNodes)
            {
                if (ProgressPart == null)
                {
                    throw new WorkItemMigratorException("Loading Settings is cancelled",
                                                        "You have stopped loading of the settings",
                                                        "To load settings, choose 'Next'");
                }

                switch (node.Name)
                {
                    case c_fieldsRootNodeTag:
                        LoadFieldNames(node);
                        break;

                    // If it is a fieid mapping node then Load Field Mappings from the Settings File Path
                    case c_fieldMappingRootNodeTag:
                        LoadFieldMappings(node);
                        break;

                    // if it is a data mapping node then Load Data Mappings from the Settings File Path
                    case c_dataMappingRootNodeTag:
                        LoadDataMappings(node);
                        break;

                    case c_parameterizationNode:
                        LoadParametrizationDelimeters(node);
                        break;

                    case c_multiLineSenseNode:
                        LoadMultiLineSense(node);
                        break;

                    case c_relationshipsNodeName:
                        LoadRelationships(node);
                        break;

                    // If it is any other node then it is invalid Settings File Path. Throw an error
                    default: throw new XmlException();
                }
            }
            ProgressPart = null;
        }

        public void InitializeProgressView()
        {
            App.CallMethodInUISynchronizationContext(InitializeProgressViewInUISynchronizationContext, null);
            while (ProgressPart == null) ;
        }

        public void SaveSettings(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                filePath = Path.Combine(Path.GetDirectoryName(Reporter.ReportFile),
                                        WorkItemGenerator.SelectedWorkItemTypeName + "-settings.xml");
            }

            // Create XML Document
            XmlDocument xmlDocument = new XmlDocument();
            XmlDeclaration dec = xmlDocument.CreateXmlDeclaration(XMLDeclarationVersion, null, null);
            xmlDocument.AppendChild(dec);

            //Creating Root Node
            XmlElement root = xmlDocument.CreateElement(c_rootNodeTag);

            // Appending the root node in the XML Documnet
            xmlDocument.AppendChild(root);


            if (DataSourceType == DataSourceType.MHT)
            {
                var fieldsRootNode = xmlDocument.CreateElement(c_fieldsRootNodeTag);
                foreach (SourceField field in DataSourceParser.StorageInfo.FieldNames)
                {
                    var fieldNode = xmlDocument.CreateElement(c_fieldNodeTag);
                    Utilities.AppendAttribute(fieldNode, c_nameAttributeTag, field.FieldName);
                    Utilities.AppendAttribute(fieldNode, c_canDeleteFieldNameAttributeTag, field.CanDelete.ToString());
                    fieldsRootNode.AppendChild(fieldNode);
                }
                MHTStorageInfo info = DataSourceParser.StorageInfo as MHTStorageInfo;
                var stepsFieldNode = xmlDocument.CreateElement(c_stepsFieldNodeTag);
                Utilities.AppendAttribute(stepsFieldNode, c_nameAttributeTag, info.StepsField);
                fieldsRootNode.AppendChild(stepsFieldNode);

                Utilities.AppendAttribute(fieldsRootNode, c_sampleFileAttributeTag, SampleFileForFields);
                root.AppendChild(fieldsRootNode);
            }

            // Creating Field Mapping node
            XmlElement fieldMappingRootNode = xmlDocument.CreateElement(c_fieldMappingRootNodeTag);

            if (DataSourceType == DataSourceType.MHT)
            {
                MHTStorageInfo info = DataSourceParser.StorageInfo as MHTStorageInfo;
                Utilities.AppendAttribute(fieldMappingRootNode, c_isFileNameTitleAttributeTag, info.IsFileNameTitle.ToString());
                Utilities.AppendAttribute(fieldMappingRootNode, c_isFirstLineTitleAttributeTag, info.IsFirstLineTitle.ToString());
            }

            // Iterating through each fieid mapping and creating corresponding node
            foreach (var kvp in Migrator.SourceNameToFieldMapping)
            {
                // Creating Field Mapping Row Node
                XmlNode fieidMappingNode = xmlDocument.CreateElement(c_fieldMappingNodeTag);

                // Appending Data Source Field Attribute in the fieid mapping Row Node
                Utilities.AppendAttribute(fieidMappingNode, c_dataSourceFieldNameAttributeTag, kvp.Key);

                // Appending TFS Workiten Field Attribute in the fieid mapping Row Node
                Utilities.AppendAttribute(fieidMappingNode, c_wiFieldNameAttributeTag, kvp.Value.TfsName);

                // Appending Field Mapping Row Node in Field Mapping Node
                fieldMappingRootNode.AppendChild(fieidMappingNode);
            }
            // Appending Field Mapping node in the root node
            root.AppendChild(fieldMappingRootNode);


            // Creating RelationShips Node
            XmlElement relationshipsNode = xmlDocument.CreateElement(c_relationshipsNodeName);
            Utilities.AppendAttribute(relationshipsNode, c_isLinkingAttributeName, IsLinking.ToString());

            if (RelationshipsInfo != null && !string.IsNullOrEmpty(RelationshipsInfo.SourceIdField))
            {
                XmlElement sourceIDFieldNode = xmlDocument.CreateElement(c_sourceIdFieldNodeName);
                sourceIDFieldNode.InnerXml = RelationshipsInfo.SourceIdField;
                relationshipsNode.AppendChild(sourceIDFieldNode);

                if (!string.IsNullOrEmpty(RelationshipsInfo.TestSuiteField))
                {
                    XmlElement testSuitesFieldNode = xmlDocument.CreateElement(c_testSuiteFieldNodeName);
                    testSuitesFieldNode.InnerXml = RelationshipsInfo.TestSuiteField;
                    relationshipsNode.AppendChild(testSuitesFieldNode);
                }

                XmlElement linksRootNode = xmlDocument.CreateElement(c_linkTypesRootNodeName);
                foreach (ILinkRule linkTypeInfo in RelationshipsInfo.LinkRules)
                {
                    XmlElement linkTypeInfoNode = xmlDocument.CreateElement(c_linkTypeNodeName);

                    Utilities.AppendAttribute(linkTypeInfoNode, c_linkSourceFieldAttributeName, linkTypeInfo.SourceFieldNameOfEndWorkItemCategory);
                    Utilities.AppendAttribute(linkTypeInfoNode, c_linkTypeNameAttributeName, linkTypeInfo.LinkTypeReferenceName);
                    Utilities.AppendAttribute(linkTypeInfoNode, c_linkedWorkItemAttributeName, WorkItemGenerator.WorkItemCategoryToDefaultType[linkTypeInfo.EndWorkItemCategory]);
                    Utilities.AppendAttribute(linkTypeInfoNode, c_descriptionAttributeName, linkTypeInfo.Description);

                    linksRootNode.AppendChild(linkTypeInfoNode);
                }
                relationshipsNode.AppendChild(linksRootNode);
            }
            root.AppendChild(relationshipsNode);


            // Creating Data Mappping Node
            XmlElement dataMappingRootNode = xmlDocument.CreateElement(c_dataMappingRootNodeTag);

            Utilities.AppendAttribute(dataMappingRootNode, c_createAreaIterationPathAttribute, WorkItemGenerator.CreateAreaIterationPath.ToString());

            // Iterating throw Data Mappings and Adding Each Data mapping entry into the DataMapping Node
            foreach (var sourceFieldNameToField in Migrator.SourceNameToFieldMapping)
            {
                // Iteration through Each Value Mapping for a Data Source Field
                foreach (KeyValuePair<string, string> valueMapping in sourceFieldNameToField.Value.ValueMapping)
                {
                    // Creating data mapping row node
                    XmlNode dataMappingNode = xmlDocument.CreateElement(c_dataMappingNodeTag);

                    // Appending Data Source Field Attribute in the Data mapping Row Node
                    Utilities.AppendAttribute(dataMappingNode, c_dataSourceFieldNameAttributeTag, sourceFieldNameToField.Key);

                    // Appending Data Source Value Attribute in the Data mapping Row Node
                    Utilities.AppendAttribute(dataMappingNode, c_dataSourceValueAttributeTag, valueMapping.Key);

                    // Appending New Value Attribute in the Data mapping Row Node
                    Utilities.AppendAttribute(dataMappingNode, c_newValueAttributeTag, valueMapping.Value);

                    // Appending Data Mapping Row Node in Data Mapping Node
                    dataMappingRootNode.AppendChild(dataMappingNode);
                }
            }
            // Appending Data Mapping Node in Root Node
            root.AppendChild(dataMappingRootNode);

            // Creating Parameterization Node
            XmlNode parameterizationNode = xmlDocument.CreateElement(c_parameterizationNode);

            Utilities.AppendAttribute(parameterizationNode, c_startParameterizationAttribute, DataSourceParser.StorageInfo.StartParameterizationDelimeter);
            Utilities.AppendAttribute(parameterizationNode, c_endParameterizationAttribute, DataSourceParser.StorageInfo.EndParameterizationDelimeter);

            root.AppendChild(parameterizationNode);

            XmlNode multiLineSenseNode = xmlDocument.CreateElement(c_multiLineSenseNode);
            multiLineSenseNode.InnerText = DataSourceParser.StorageInfo.IsMultilineSense.ToString();
            root.AppendChild(multiLineSenseNode);

            if (!Directory.Exists(Path.GetDirectoryName(filePath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));
            }

            // Writing XML Document in OutputSettingsFilePath
            using (StreamWriter tr = new StreamWriter(filePath))
            {
                tr.Write(xmlDocument.OuterXml);
            }
        }

        #endregion

        #region Private methods

        ///<summary>
        /// Utility method to Load Data Mappings from Data Mapping node
        ///</summary>
        ///<param name="root"></param>
        private void LoadDataMappings(XmlElement root)
        {
            ParseDataSource();

            if (ProgressPart != null)
            {
                ProgressPart.Text = "Loading Data Mapping...";
            }

            if (root.Attributes.Count == 1)
            {
                WorkItemGenerator.CreateAreaIterationPath = bool.Parse(root.Attributes[0].Value);
            }

            // Iterate through each dat mapping row node and update data mapping
            foreach (XmlElement node in root.ChildNodes)
            {
                // If node is not dat mapping row node then throw XML error
                if (String.CompareOrdinal(node.Name, c_dataMappingNodeTag) != 0)
                {
                    throw new XmlException();
                }

                // Get fieldName, fieldValue
                string fieldName = UpdateLTGTChars(node.Attributes[c_dataSourceFieldNameAttributeTag].InnerXml);
                string fieldValue = UpdateLTGTChars(node.Attributes[c_dataSourceValueAttributeTag].InnerXml);
                string newValue = UpdateLTGTChars(node.Attributes[c_newValueAttributeTag].InnerXml);

                UpdateDataMapping(fieldName, fieldValue, newValue);
            }
        }

        private void UpdateDataMapping(string fieldName, string fieldValue, string newValue)
        {
            if (string.IsNullOrEmpty(fieldName) || string.IsNullOrEmpty(fieldValue) || string.IsNullOrEmpty(newValue))
            {
                throw new XmlException();
            }
            bool isValidFieldName = false;
            foreach (SourceField field in DataSourceParser.StorageInfo.FieldNames)
            {
                if (String.CompareOrdinal(field.FieldName, fieldName) == 0)
                {
                    isValidFieldName = true;
                    break;
                }
            }

            if (isValidFieldName &&
                Migrator.SourceNameToFieldMapping.ContainsKey(fieldName) &&
                !Migrator.SourceNameToFieldMapping[fieldName].ValueMapping.ContainsKey(fieldValue))
            {
                Migrator.SourceNameToFieldMapping[fieldName].ValueMapping.Add(fieldValue, newValue);
            }
        }

        private void LoadFieldMappings(XmlElement root)
        {
            if (ProgressPart != null)
            {
                ProgressPart.Text = "Loading Field Mapping...";
            }

            foreach (var kvp in WorkItemGenerator.TfsNameToFieldMapping)
            {
                kvp.Value.SourceName = string.Empty;
            }

            if (DataSourceType == DataSourceType.MHT && root.Attributes != null)
            {
                bool isFirstLineTitle = false;
                bool isFileNameTitle = false;
                foreach (XmlAttribute attribute in root.Attributes)
                {
                    if (String.CompareOrdinal(attribute.Name, c_isFileNameTitleAttributeTag) == 0)
                    {
                        bool.TryParse(attribute.Value, out isFileNameTitle);
                    }
                    if (String.CompareOrdinal(attribute.Name, c_isFirstLineTitleAttributeTag) == 0)
                    {
                        bool.TryParse(attribute.Value, out isFirstLineTitle);
                    }
                }
                if (isFileNameTitle ^ isFirstLineTitle)
                {
                    MHTStorageInfo info = DataSourceParser.StorageInfo as MHTStorageInfo;
                    info.TitleField = MHTParser.TestTitleDefaultTag;
                    info.IsFileNameTitle = isFileNameTitle;
                    info.IsFirstLineTitle = isFirstLineTitle;
                }
            }


            var sourceNameToFieldMapping = new Dictionary<string, IWorkItemField>();
            foreach (XmlElement node in root.ChildNodes)
            {
                if (String.CompareOrdinal(node.Name, c_fieldMappingNodeTag) != 0)
                {
                    throw new XmlException();
                }

                string dataSourceField = UpdateLTGTChars(node.Attributes[c_dataSourceFieldNameAttributeTag].InnerXml);
                string wiField = UpdateLTGTChars(node.Attributes[c_wiFieldNameAttributeTag].InnerXml);

                if (string.IsNullOrEmpty(dataSourceField) || string.IsNullOrEmpty(wiField))
                {
                    throw new XmlException();
                }

                if (WorkItemGenerator.TfsNameToFieldMapping.ContainsKey(wiField))
                {
                    WorkItemGenerator.TfsNameToFieldMapping[wiField].SourceName = dataSourceField;
                    sourceNameToFieldMapping.Add(dataSourceField, WorkItemGenerator.TfsNameToFieldMapping[wiField]);
                }
            }
            DataSourceParser.FieldNameToFields = sourceNameToFieldMapping;
            WorkItemGenerator.SourceNameToFieldMapping = sourceNameToFieldMapping;
            Migrator.SourceNameToFieldMapping = sourceNameToFieldMapping;
        }

        private void ParseDataSource()
        {
            IDictionary<string, IWorkItemField> sourceNameToFieldMapping = Migrator.SourceNameToFieldMapping;
            int count = 0;
            if (ProgressPart != null)
            {
                ProgressPart.Text = "Parsing Source...";
            }
            Console.WriteLine("\nParsing Source...\n");
            IList<ISourceWorkItem> sourceWorkItems = new List<ISourceWorkItem>();
            IList<ISourceWorkItem> rawSourceWorkItems = new List<ISourceWorkItem>();
            if (DataSourceType == DataSourceType.MHT)
            {
                MHTStorageInfo sampleInfo = DataSourceParser.StorageInfo as MHTStorageInfo;

                IList<IDataStorageInfo> storageInfos = DataStorageInfos;

                foreach (IDataStorageInfo storageInfo in storageInfos)
                {
                    if (ProgressPart == null)
                    {
                        break;
                    }
                    MHTStorageInfo info = storageInfo as MHTStorageInfo;
                    info.IsFirstLineTitle = sampleInfo.IsFirstLineTitle;
                    info.IsFileNameTitle = sampleInfo.IsFileNameTitle;
                    info.TitleField = sampleInfo.TitleField;
                    info.StepsField = sampleInfo.StepsField;
                    IDataSourceParser parser = new MHTParser(info);

                    info.FieldNames = sampleInfo.FieldNames;
                    parser.FieldNameToFields = sourceNameToFieldMapping;

                    while (parser.GetNextWorkItem() != null)
                    {
                        count++;
                        if (ProgressPart != null)
                        {
                            ProgressPart.Text = "Parsing " + count + " of " + DataStorageInfos.Count + ":\n" + info.Source;
                            Console.Write("\r" + ProgressPart.Text);
                            ProgressPart.ProgressValue = (count * 100) / DataStorageInfos.Count;
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
            else if (DataSourceType == DataSourceType.Excel)
            {
                var parser = DataSourceParser;
                var excelInfo = parser.StorageInfo as ExcelStorageInfo;
                while (parser.GetNextWorkItem() != null)
                {
                    if (ProgressPart == null)
                    {
                        break;
                    }
                    count++;
                    ProgressPart.ProgressValue = excelInfo.ProgressPercentage;
                    ProgressPart.Text = "Parsing work item # " + count;
                    Console.Write("\r" + ProgressPart.Text);
                }
                for (int i = 0; i < parser.ParsedSourceWorkItems.Count; i++)
                {
                    sourceWorkItems.Add(parser.ParsedSourceWorkItems[i]);
                    rawSourceWorkItems.Add(parser.RawSourceWorkItems[i]);
                }
            }

            Migrator.RawSourceWorkItems = rawSourceWorkItems;
            Migrator.SourceWorkItems = sourceWorkItems;

            if (Reporter != null)
            {
                Reporter.Dispose();
            }
            string reportDirectory = Path.Combine(Path.GetDirectoryName(DataSourceParser.StorageInfo.Source),
                                                  "Report" + DateTime.Now.ToString("g", System.Globalization.CultureInfo.CurrentCulture).
                                                  Replace(":", "_").Replace(" ", "_").Replace("/", "_"));
            switch (DataSourceType)
            {
                case DataSourceType.Excel:
                    Reporter = new ExcelReporter(this);
                    string fileNameWithoutExtension = "Report";
                    string fileExtension = Path.GetExtension(DataSourceParser.StorageInfo.Source);
                    Reporter.ReportFile = Path.Combine(reportDirectory, fileNameWithoutExtension + fileExtension);
                    break;

                case DataSourceType.MHT:
                    Reporter = new XMLReporter(this);
                    string fileName = "Report.xml"; ;
                    Reporter.ReportFile = Path.Combine(reportDirectory, fileName);
                    break;

                default:
                    throw new InvalidEnumArgumentException("Invalid Data Source type");
            }
        }

        private void LoadMultiLineSense(XmlElement node)
        {
            DataSourceParser.StorageInfo.IsMultilineSense = bool.Parse(node.InnerXml);
        }

        private void LoadParametrizationDelimeters(XmlElement node)
        {
            DataSourceParser.StorageInfo.StartParameterizationDelimeter = UpdateLTGTChars(node.Attributes[c_startParameterizationAttribute].InnerXml);
            DataSourceParser.StorageInfo.EndParameterizationDelimeter = UpdateLTGTChars(node.Attributes[c_endParameterizationAttribute].InnerXml);
        }

        private string UpdateLTGTChars(string value)
        {
            return value.Replace(LTXMLChar, LTChar).Replace(GTXMLChar, GTChar);
        }

        private void InitializeProgressViewInUISynchronizationContext(object obj)
        {
            ProgressPart = new ProgressWindow();
            ProgressPart.PropertyChanged += new PropertyChangedEventHandler(ProgressPart_PropertyChanged);
            if (!AutoWaitCursor.IsConsoleMode)
            {
                ProgressPart.Owner = App.Application.MainWindow;
                ProgressPart.Show();
            }
        }

        private void ProgressPart_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (String.CompareOrdinal(e.PropertyName, ProgressWindow.IsClosedPropertyName) == 0)
            {
                if (m_progressPart != null && m_progressPart.IsClosed)
                {
                    m_progressPart.PropertyChanged -= ProgressPart_PropertyChanged;
                    m_progressPart = null;
                }
            }
        }

        private void LoadFieldNames(XmlElement node)
        {
            if (ProgressPart != null)
            {
                ProgressPart.Text = "Loading Field Names...";
            }
            if (DataSourceType == DataSourceType.MHT)
            {
                if (node.Attributes.Count == 1 &&
                    String.CompareOrdinal(node.Attributes[0].Name, c_sampleFileAttributeTag) == 0 &&
                    File.Exists(node.Attributes[0].Value))
                {

                    DataSourceParser = new MHTParser(new MHTStorageInfo(node.Attributes[0].Value));
                }
                DataSourceParser.ParseDataSourceFieldNames();
                DataSourceParser.StorageInfo.FieldNames.Clear();
                foreach (XmlElement fieldNode in node.ChildNodes)
                {
                    if (String.CompareOrdinal(fieldNode.Name, c_fieldNodeTag) == 0 &&
                        fieldNode.Attributes.Count == 2 &&
                        String.CompareOrdinal(fieldNode.Attributes[0].Name, c_nameAttributeTag) == 0 &&
                        String.CompareOrdinal(fieldNode.Attributes[1].Name, c_canDeleteFieldNameAttributeTag) == 0)
                    {
                        string fieldName = fieldNode.Attributes[0].Value;
                        bool canDelete = false;
                        if (bool.TryParse(fieldNode.Attributes[1].Value, out canDelete))
                        {
                            DataSourceParser.StorageInfo.FieldNames.Add(new SourceField(fieldName, canDelete));
                        }
                    }
                    else if (String.CompareOrdinal(fieldNode.Name, c_stepsFieldNodeTag) == 0 &&
                        fieldNode.Attributes.Count == 1 &&
                        String.CompareOrdinal(fieldNode.Attributes[0].Name, c_nameAttributeTag) == 0)
                    {
                        MHTStorageInfo info = DataSourceParser.StorageInfo as MHTStorageInfo;
                        info.StepsField = fieldNode.Attributes[0].Value;
                    }
                }
            }
        }

        private void LoadRelationships(XmlElement relationshipsNode)
        {
            RelationshipsInfo = new RelationshipsInfo();

            DataSourceParser.StorageInfo.SourceIdFieldName = null;
            DataSourceParser.StorageInfo.TestSuiteFieldName = null;
            DataSourceParser.StorageInfo.LinkRules.Clear();

            if (relationshipsNode.Attributes.Count == 1)
            {
                bool isLinking;
                bool.TryParse(relationshipsNode.Attributes[0].Value, out isLinking);
                IsLinking = isLinking;
            }

            // Traverse each child node of the root node
            foreach (XmlElement node in relationshipsNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case c_sourceIdFieldNodeName:
                        string sourceIdField = node.InnerXml;
                        if (IsValidFieldName(sourceIdField) &&
                            (!Migrator.SourceNameToFieldMapping.ContainsKey(sourceIdField) ||
                                Migrator.SourceNameToFieldMapping[sourceIdField].IsIdField))
                        {
                            RelationshipsInfo.SourceIdField = sourceIdField;
                            DataSourceParser.StorageInfo.SourceIdFieldName = sourceIdField;
                        }
                        break;

                    case c_testSuiteFieldNodeName:
                        string testSuiteField = node.InnerXml;
                        if (IsValidFieldName(testSuiteField) &&
                            !Migrator.SourceNameToFieldMapping.ContainsKey(testSuiteField) &&
                            String.CompareOrdinal(RelationshipsInfo.SourceIdField, testSuiteField) != 0)
                        {
                            RelationshipsInfo.TestSuiteField = testSuiteField;
                            DataSourceParser.StorageInfo.TestSuiteFieldName = testSuiteField;
                        }
                        break;

                    case c_linkTypesRootNodeName:
                        foreach (XmlElement linkNode in node.ChildNodes)
                        {
                            if (String.CompareOrdinal(linkNode.Name, c_linkTypeNodeName) == 0 &&
                                linkNode.Attributes.Count == 4 &&
                                String.CompareOrdinal(linkNode.Attributes[0].Name, c_linkSourceFieldAttributeName) == 0 &&
                                String.CompareOrdinal(linkNode.Attributes[1].Name, c_linkTypeNameAttributeName) == 0 &&
                                String.CompareOrdinal(linkNode.Attributes[2].Name, c_linkedWorkItemAttributeName) == 0 &&
                                String.CompareOrdinal(linkNode.Attributes[3].Name, c_descriptionAttributeName) == 0)
                            {
                                string sourceField = linkNode.Attributes[0].Value;
                                string linkTypeName = linkNode.Attributes[1].Value;
                                string linkedWorkItemTypeName = linkNode.Attributes[2].Value;
                                string linkDescription = linkNode.Attributes[3].Value;

                                if (IsValidFieldName(sourceField) &&
                                    !Migrator.SourceNameToFieldMapping.ContainsKey(sourceField) &&
                                    String.CompareOrdinal(sourceField, RelationshipsInfo.SourceIdField) != 0 &&
                                    String.CompareOrdinal(sourceField, RelationshipsInfo.TestSuiteField) != 0 &&
                                    WorkItemGenerator.LinkTypeNames.Contains(linkTypeName) &&
                                    WorkItemGenerator.WorkItemTypeToCategoryMapping.ContainsKey(WorkItemGenerator.SelectedWorkItemTypeName) &&
                                    WorkItemGenerator.WorkItemTypeToCategoryMapping.ContainsKey(linkedWorkItemTypeName))
                                {
                                    ILinkRule linkTypeInfo = new LinkRule(WorkItemGenerator.WorkItemTypeToCategoryMapping[WorkItemGenerator.SelectedWorkItemTypeName],
                                                                          sourceField,
                                                                          linkTypeName,
                                                                          WorkItemGenerator.WorkItemTypeToCategoryMapping[linkedWorkItemTypeName],
                                                                          linkDescription);
                                    RelationshipsInfo.LinkRules.Add(linkTypeInfo);
                                    DataSourceParser.StorageInfo.LinkRules.Add(linkTypeInfo);
                                }
                            }
                        }
                        break;

                    default:
                        break;
                }
            }
        }

        /// <summary>
        ///  Save Current Settings to the OutputSettingsFilePath
        /// </summary>


        private bool IsValidFieldName(string sourceField)
        {
            bool isValidFieldName = false;
            foreach (var field in DataSourceParser.StorageInfo.FieldNames)
            {
                if (String.CompareOrdinal(field.FieldName, sourceField) == 0)
                {
                    isValidFieldName = true;
                    break;
                }
            }
            return isValidFieldName;
        }

        private void CloseProgressPart(object obj)
        {
            if (m_progressPart != null)
            {
                m_progressPart.Close();
            }
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose()
        {
            // Disposing Data Source
            if (DataSourceParser != null)
            {
                DataSourceParser.Dispose();
            }

            if (WorkItemGenerator != null)
            {
                WorkItemGenerator.Dispose();
            }

            if (Reporter != null)
            {
                Reporter.Dispose();
            }
        }
        #endregion
    }
}