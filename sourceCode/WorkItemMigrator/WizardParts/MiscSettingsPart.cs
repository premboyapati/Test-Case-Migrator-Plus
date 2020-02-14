//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.IO;

    /// <summary>
    /// Wizard Part for taking Miscellaneous settings for test steps, output setting file and 
    /// report which is going to be published
    /// </summary>
    internal class MiscSettingsPart : BaseWizardPart
    {
        #region Fields

        // member variabled needed for data binding
        private bool m_isParameterizationChecked;
        private bool m_isMultiLineSenseEnabled;
        private string m_startParameterizationDelimeter;
        private string m_endParameterizationDelimeter;
        private string m_outputMappingsFile;
        private string m_reportFolder;

        #endregion

        #region Constants

        // Excel Extension
        private const string ExcelExtension = ".xls";

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        public MiscSettingsPart()
        {
            Header = Resources.MiscSettings_Header;
            Description = Resources.MiscSettings_ExcelFlow_Description;
            CanBack = true;
            CanNext = true;
            WizardPage = WizardPage.MiscSettings;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Is Data Source Workitem's Test Steps have Parameters
        /// </summary>
        public bool IsParameterizationChecked
        {
            get
            {
                return m_isParameterizationChecked;
            }
            set
            {
                m_isParameterizationChecked = value;
                NotifyPropertyChanged("IsParameterizationChecked");
            }
        }

        /// <summary>
        ///  Are Test Step in Data Source Workitem have multiple lines in single block of data
        ///  for Excel: block of data would be one cell
        /// </summary>
        public bool IsMultiLineSenseEnabled
        {
            get
            {
                return m_isMultiLineSenseEnabled;
            }
            set
            {
                m_isMultiLineSenseEnabled = value;
                NotifyPropertyChanged("IsMultiLineSenseEnabled");
            }
        }

        /// <summary>
        /// Start Delimeter for Parameters in the Data Source Workitem
        /// </summary>
        public string StartParameterizationDelimeter
        {
            get
            {
                return m_startParameterizationDelimeter;
            }
            set
            {
                m_startParameterizationDelimeter = value;
                NotifyPropertyChanged("StartParameterizationDelimeter");
            }
        }

        /// <summary>
        /// End Delimeter for Data Source Workitem
        /// </summary>
        public string EndParameterizationDelimeter
        {
            get
            {
                return m_endParameterizationDelimeter;
            }
            set
            {
                m_endParameterizationDelimeter = value;
                NotifyPropertyChanged("EndParameterizationDelimeter");
            }
        }

        /// <summary>
        /// Location of report which is going to publish
        /// </summary>
        public string ReportFolder
        {
            get
            {
                return m_reportFolder;
            }
            set
            {
                m_reportFolder = value;
                CanNext = ValidatePartState();
                NotifyPropertyChanged("ReportFolder");
            }
        }

        /// <summary>
        /// Location of Settings file which is going to save
        /// </summary>
        public string OutputMappingsFile
        {
            get
            {
                return m_outputMappingsFile;
            }
            set
            {
                m_outputMappingsFile = value;
                NotifyPropertyChanged("OutputMappingsFile");
                CanNext = ValidatePartState();
            }
        }

        public bool AreMiscSettingsVisible
        {
            get
            {
                return WizardInfo != null &&
                      WizardInfo.DataSourceType == DataSourceType.Excel &&
                      WizardInfo.WorkItemGenerator != null &&
                      WizardInfo.WorkItemGenerator.WorkItemCategory == WorkItemGenerator.TestCaseCategory;
            }
        }

        #endregion

        #region public methods

        /// <summary>
        /// resets the Misc Part state to initial one
        /// </summary>
        public override void Reset()
        {
            if (m_wizardInfo.DataSourceType == DataSourceType.MHT)
            {
                Description = "Specify the path for settings file and output logs.";
            }
            else if (m_wizardInfo.DataSourceType == DataSourceType.Excel)
            {
                Description = Resources.MiscSettings_ExcelFlow_Description;
            }

            // Is Multi Line sense enabled
            IsMultiLineSenseEnabled = m_wizardInfo.DataSourceParser.StorageInfo.IsMultilineSense;

            // Does Source File has parameters
            IsParameterizationChecked = (!string.IsNullOrEmpty(m_wizardInfo.DataSourceParser.StorageInfo.StartParameterizationDelimeter) ||
                                        !string.IsNullOrEmpty(m_wizardInfo.DataSourceParser.StorageInfo.EndParameterizationDelimeter));

            // Start Delimeter
            StartParameterizationDelimeter = m_wizardInfo.DataSourceParser.StorageInfo.StartParameterizationDelimeter;

            // End Delimeter
            EndParameterizationDelimeter = !string.IsNullOrEmpty(m_wizardInfo.DataSourceParser.StorageInfo.EndParameterizationDelimeter) ?
                m_wizardInfo.DataSourceParser.StorageInfo.EndParameterizationDelimeter : string.Empty;

            // Report File
            ReportFolder = Path.GetDirectoryName(m_wizardInfo.Reporter.ReportFile);

            // Save Mppings File
            if (m_wizardInfo.OutputSettingsFilePath == null)
            {
                OutputMappingsFile = Path.Combine(ReportFolder, WizardInfo.WorkItemGenerator.SelectedWorkItemTypeName + "-settings.xml");
            }
            else
            {
                OutputMappingsFile = m_wizardInfo.OutputSettingsFilePath;
            }

            NotifyPropertyChanged("AreMiscSettingsVisible");
        }

        /// <summary>
        /// Updates Wizard Info with the current state of Misc Part
        /// </summary>
        /// <returns></returns>
        public override bool UpdateWizardPart()
        {
            if (!ValidatePartState())
            {
                CanNext = false;
                return false;
            }

            // If Data Source Has parameters
            if (IsParameterizationChecked)
            {
                // then update start and end delimeters
                m_wizardInfo.DataSourceParser.StorageInfo.StartParameterizationDelimeter = StartParameterizationDelimeter;
                m_wizardInfo.DataSourceParser.StorageInfo.EndParameterizationDelimeter = EndParameterizationDelimeter;
            }
            else
            {
                // else set them empty
                m_wizardInfo.DataSourceParser.StorageInfo.StartParameterizationDelimeter = string.Empty;
                m_wizardInfo.DataSourceParser.StorageInfo.EndParameterizationDelimeter = string.Empty;
            }

            m_wizardInfo.DataSourceParser.StorageInfo.IsMultilineSense = IsMultiLineSenseEnabled;

            m_wizardInfo.OutputSettingsFilePath = OutputMappingsFile;

            if (m_wizardInfo.DataSourceType == DataSourceType.Excel)
            {
                m_wizardInfo.Reporter.ReportFile = Path.Combine(ReportFolder, "Report.xls");
            }
            else if (m_wizardInfo.DataSourceType == DataSourceType.MHT)
            {
                m_wizardInfo.Reporter.ReportFile = Path.Combine(ReportFolder, "Report.xml");
            }
            return true;
        }

        #endregion

        #region protected/private methods

        public override bool ValidatePartState()
        {
            Warning = null;
            return IsValidFolder(ReportFolder) && IsValidXMLFilePath(OutputMappingsFile);
        }

        /// <summary>
        /// We can initialize this part if data source and server connection are provided 
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
        protected override bool CanInitializeWizardPage(WizardInfo info)
        {
            string title = null;
            string likelyCause = null;
            string potentialSolution = null;

            m_canShow = true;

            if (info.DataSourceParser == null)
            {
                m_canShow = false;
                title = Resources.MiscSettingsPart_CantShowTitle;
                likelyCause = Resources.DataSourceNotEnteredErrorLikelyCause;
                potentialSolution = Resources.DataSourceNotEnteredErrorPotentialSolution;
            }
            else if (info.DataSourceParser.StorageInfo.FieldNames == null || info.DataSourceParser.StorageInfo.FieldNames.Count == 0)
            {
                m_canShow = false;
                title = Resources.MiscSettingsPart_CantShowTitle;
                likelyCause = Resources.DataSourceFieldNamesNotFoundErrorLikelyCause;
                potentialSolution = Resources.DataSourceFieldNamesNotFoundErrorPotentialSolution;
            }
            else if (info.WorkItemGenerator.TfsNameToFieldMapping == null || info.WorkItemGenerator.TfsNameToFieldMapping.Count == 0)
            {
                m_canShow = false;
                title = Resources.MiscSettingsPart_CantShowTitle;
                likelyCause = Resources.ServerNotSpecifiedErrorLikelyCause;
                potentialSolution = Resources.ServerNotSpecifiedErrorPotentialSolution;
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

            if (info.IsLinking && string.IsNullOrEmpty(info.DataSourceParser.StorageInfo.SourceIdFieldName))
            {
                title = "Source Id is not specified for Linking";
                m_canShow = false;
            }


            if (!m_canShow)
            {
                Warning = title;
            }
            return m_canShow;
        }

        private bool IsValidFolder(string folder)
        {
            try
            {
                if (Directory.Exists(folder))
                {
                    string tempFileName = Path.Combine(folder, Guid.NewGuid().ToString());
                    using (new StreamWriter(tempFileName))
                    { }
                    File.Delete(tempFileName);
                }
                else
                {
                    Directory.CreateDirectory(folder);
                    Directory.Delete(folder);
                }
                return true;
            }
            catch (IOException)
            {
                Warning = "Report Path is not accessible";
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                Warning = "Report Path is not accessible";
                return false;
            }
        }

        private bool IsValidXMLFilePath(string file)
        {
            if (string.IsNullOrEmpty(OutputMappingsFile))
            {
                return true;
            }
            try
            {
                string folder = Path.GetDirectoryName(file);
                if (Directory.Exists(folder))
                {
                    string tempFileName = Path.Combine(folder, Guid.NewGuid().ToString() + ".txt");
                    using (new StreamWriter(tempFileName))
                    { }
                    File.Delete(tempFileName);
                }
                else
                {
                    Directory.CreateDirectory(folder);
                    Directory.Delete(folder);
                }
                if (String.CompareOrdinal(Path.GetExtension(file), ".xml") != 0)
                {
                    Warning = "Invalid settings file path";
                    return false;
                }

                return true;
            }
            catch (IOException)
            {
                Warning = "Settings file path is not accessible";
                return false;
            }
            catch (ArgumentException)
            {
                Warning = "Invalid settings file path";
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                Warning = "Settings file path is not accessible";
                return false;
            }
        }

        protected override bool IsInitializationRequired(WizardInfo state)
        {
            return m_wizardInfo == null ||
                m_prerequisite.IsSettingsFilePathChanged() ||
                m_prerequisite.IsDataSourceTypeChanged() ||
                m_prerequisite.IsServerConnectionChanged();
        }

        #endregion
    }
}