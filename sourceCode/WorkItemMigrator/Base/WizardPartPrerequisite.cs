//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// This class is used by Wizard parts to see whether the changes made in other wizard parts 
    /// requires the updation on its current state or not.
    /// </summary>
    class WizardPartPrerequisite
    {
        #region Fields

        // source loaction containg mht files
        private string m_mhtSource;

        // data source type(excel/mht)
        private DataSourceType m_dataSourceType;

        // It is used to store the state of Data Source Wizard part
        private IDataStorageInfo m_dataStorageInfo;

        // Following are used to save the satte of "Destination TFS Server/project" Wizard part
        private string m_tfsServer;
        private string m_teamProject;
        private string m_selctedWIType;

        // settings file path
        private string m_settingsFile;

        // list of fields found in the data source(mht/excel). It also contains user added fields
        private List<SourceField> m_fields = new List<SourceField>();

        // Fields Mapping between source and tfs fields
        private Dictionary<string, string> m_fieldMapping = new Dictionary<string, string>();

        // Wizard Info object needed for getting the current state of Wizard
        private WizardInfo m_wizardInfo;

        #endregion

        #region Constructor

        /// <summary>
        /// Initialize the Wizard Info
        /// </summary>
        /// <param name="wizardInfo"></param>
        public WizardPartPrerequisite(WizardInfo wizardInfo)
        {
            m_wizardInfo = wizardInfo;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Saves the current state of wizard Info
        /// </summary>
        /// <param name="wizardInfo"></param>
        public void Save()
        {
            // Saving Data Source State
            SaveDataSourceInfo();

            // Saving Server State
            if (m_wizardInfo.WorkItemGenerator != null)
            {
                m_tfsServer = m_wizardInfo.WorkItemGenerator.Server;
                m_teamProject = m_wizardInfo.WorkItemGenerator.Project;
                m_selctedWIType = m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName;
            }

            // Saving settings file path
            m_settingsFile = m_wizardInfo.InputSettingsFilePath;

            // saving field names
            m_fields.Clear();
            if (m_wizardInfo.DataSourceParser != null && m_wizardInfo.DataSourceParser.StorageInfo.FieldNames != null)
            {
                foreach (SourceField field in m_wizardInfo.DataSourceParser.StorageInfo.FieldNames)
                {
                    m_fields.Add(field);
                }
            }

            // saving field mapping
            if (m_wizardInfo.Migrator != null && m_wizardInfo.Migrator.SourceNameToFieldMapping != null)
            {
                m_fieldMapping.Clear();
                foreach (var kvp in m_wizardInfo.Migrator.SourceNameToFieldMapping)
                {
                    m_fieldMapping.Add(kvp.Key, kvp.Value.TfsName);
                }
            }
        }

        /// <summary>
        /// Tells the Wizard Part that Data Source is updated from the last time or not
        /// </summary>
        /// <returns></returns>
        public bool IsDataSourceTypeChanged()
        {
            return m_wizardInfo.DataSourceType != m_dataSourceType;
        }

        /// <summary>
        /// Is Data Source Type changed by the User in Current Wizard Part
        /// </summary>
        /// <param name="dataSourceType"></param>
        /// <returns></returns>
        public bool IsDataSourceTypeModified(DataSourceType dataSourceType)
        {
            return m_dataSourceType != dataSourceType;
        }

        /// <summary>
        /// Is Data Source Changed by SelectDataSourcepart
        /// </summary>
        /// <returns></returns>
        public bool IsDataSourceChanged()
        {
            // Return true if wizard is going to be intialized first time or data source type is changed
            if (m_wizardInfo == null ||
                m_wizardInfo.DataSourceType != m_dataSourceType ||
                m_dataStorageInfo == null)
            {
                return true;
            }
            else
            {
                switch (m_wizardInfo.DataSourceType)
                {
                    case DataSourceType.Excel:
                        ExcelStorageInfo acturalInfo = m_wizardInfo.DataSourceParser.StorageInfo as ExcelStorageInfo;
                        ExcelStorageInfo expectedInfo = m_dataStorageInfo as ExcelStorageInfo;

                        // if excel source file path is changed or worksheet name is changed or row containg field names is changed then return true
                        if (String.CompareOrdinal(expectedInfo.Source, acturalInfo.Source) != 0 ||
                            String.CompareOrdinal(expectedInfo.WorkSheetName, acturalInfo.WorkSheetName) != 0 ||
                            String.CompareOrdinal(expectedInfo.RowContainingFieldNames, acturalInfo.RowContainingFieldNames) != 0)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }

                    case DataSourceType.MHT:
                        // return true If MHT Source is changed or user have added new fields 
                        if (String.CompareOrdinal(m_wizardInfo.MHTSource, m_mhtSource) != 0 ||
                            (m_wizardInfo.DataSourceParser.StorageInfo.FieldNames.Count != m_fields.Count))
                        {
                            return true;
                        }

                        // return trueif user have modified any field name
                        foreach (SourceField field in m_fields)
                        {
                            if (!m_wizardInfo.DataSourceParser.StorageInfo.FieldNames.Contains(field))
                            {
                                return true;
                            }
                        }
                        return false;

                    default:
                        return true;
                }
            }
        }

        /// <summary>
        /// Is Excel Source Information Modified in Select Data Source Part by the User
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <param name="workSheetName"></param>
        /// <param name="headerRow"></param>
        /// <returns></returns>
        public bool IsExcelSourceModified(string excelFilePath, string workSheetName, string headerRow)
        {
            ExcelStorageInfo excelInfo = m_dataStorageInfo as ExcelStorageInfo;
            if (excelInfo == null)
            {
                return true;
            }
            if (excelInfo.Source == excelFilePath &&
                excelInfo.WorkSheetName == workSheetName &&
                excelInfo.RowContainingFieldNames == headerRow)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// IS mht Source Information modified by the user in Select Data Source Part
        /// </summary>
        /// <param name="isMHTFolder"></param>
        /// <param name="mhtFolder"></param>
        /// <param name="listOfMHTsPath"></param>
        /// <returns></returns>
        public bool IsMHTSourceModified(bool isMHTFolder, string mhtFolder, string listOfMHTsPath)
        {
            string mhtSource = isMHTFolder ? mhtFolder : listOfMHTsPath;
            return String.CompareOrdinal(mhtSource, m_wizardInfo.MHTSource) != 0;
        }

        ///// <summary>
        ///// Tells the Wizard Part that Server/project settings are updated from last time or not
        ///// </summary>
        ///// <returns></returns>
        public bool IsServerConnectionChanged()
        {
            if (m_wizardInfo == null)
            {
                return true;
            }
            else
            {
                return (String.CompareOrdinal(m_tfsServer, m_wizardInfo.WorkItemGenerator.Server) != 0 ||
                        String.CompareOrdinal(m_teamProject, m_wizardInfo.WorkItemGenerator.Project) != 0 ||
                        String.CompareOrdinal(m_selctedWIType, m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName) != 0);
            }
        }

        /// <summary>
        /// Is Server Connection details modified by the user in selct destination server part
        /// </summary>
        /// <param name="serverURL"></param>
        /// <param name="project"></param>
        /// <param name="workItemtype"></param>
        /// <returns></returns>
        public bool IsServerConnectionModified(string serverURL, string project, string workItemtype)
        {
            if (m_tfsServer == serverURL &&
                m_teamProject == project &&
                m_selctedWIType == workItemtype)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Are settings file changed from the last time?
        /// </summary>
        /// <returns></returns>
        public bool IsSettingsFilePathChanged()
        {
            return String.CompareOrdinal(m_settingsFile, m_wizardInfo.InputSettingsFilePath) != 0;
        }

        /// <summary>
        /// Are fields changed from last time?
        /// </summary>
        /// <returns></returns>
        public bool AreFieldsChanged()
        {
            IList<SourceField> fields = m_wizardInfo.DataSourceParser.StorageInfo.FieldNames;
            return AreFieldsModified(fields);
        }

        /// <summary>
        /// Are fields modifed in Fields Selection Part by the user
        /// </summary>
        /// <param name="fields"></param>
        /// <returns></returns>
        public bool AreFieldsModified(IList<SourceField> fields)
        {
            if (m_fields == null || m_fields.Count != fields.Count)
            {
                return true;
            }
            for (int i = 0; i < m_fields.Count; i++)
            {
                if (String.CompareOrdinal(m_fields[i].FieldName, fields[i].FieldName) != 0 &&
                    m_fields[i].CanDelete != fields[i].CanDelete)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Is Field Mapping changed from last time?
        /// </summary>
        /// <returns></returns>
        public bool IsFieldMappingChanged()
        {
            if (m_fieldMapping == null || m_fieldMapping.Count != m_wizardInfo.Migrator.SourceNameToFieldMapping.Count)
            {
                return true;
            }
            foreach (var kvp in m_wizardInfo.Migrator.SourceNameToFieldMapping)
            {
                if (!m_fieldMapping.ContainsKey(kvp.Key) ||
                    String.CompareOrdinal(m_fieldMapping[kvp.Key], kvp.Value.TfsName) != 0)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Is User have modifed the field mappings in FieldMappingPart?
        /// </summary>
        /// <param name="columnMappingRows"></param>
        /// <param name="isFirstLineTitle"></param>
        /// <param name="isFileNameTitle"></param>
        /// <returns></returns>
        public bool IsFieldMappingModified(ICollection<FieldMappingRow> columnMappingRows, bool isFirstLineTitle, bool isFileNameTitle)
        {
            if (m_dataSourceType == DataSourceType.Excel &&
                (m_fieldMapping == null || m_fieldMapping.Count == 0))
            {
                return true;
            }

            if (m_dataSourceType == DataSourceType.MHT)
            {
                MHTStorageInfo info = m_wizardInfo.DataSourceParser.StorageInfo as MHTStorageInfo;
                if (info.IsFileNameTitle != isFileNameTitle || info.IsFirstLineTitle != isFirstLineTitle)
                {
                    return true;
                }
            }

            int fieldMappingCount = 0;
            foreach (var row in columnMappingRows)
            {
                if (row.WIField != null)
                {
                    if (!m_fieldMapping.ContainsKey(row.DataSourceField) ||
                        String.CompareOrdinal(m_fieldMapping[row.DataSourceField], row.WIField.TfsName) != 0)
                    {
                        return true;
                    }
                    fieldMappingCount++;
                }
            }

            if (m_dataSourceType == DataSourceType.MHT)
            {
                if (isFirstLineTitle || isFileNameTitle)
                {
                    fieldMappingCount++;
                }
            }

            if (fieldMappingCount == m_fieldMapping.Count)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        #endregion

        #region private methods

        /// <summary>
        /// Saves the Data Source Information
        /// </summary>
        /// <param name="wizardInfo"></param>
        private void SaveDataSourceInfo()
        {
            if (m_wizardInfo.DataSourceParser == null)
            {
                return;
            }

            m_dataSourceType = m_wizardInfo.DataSourceType;

            switch (m_wizardInfo.DataSourceType)
            {
                case DataSourceType.Excel:
                    ExcelStorageInfo sourceInfo = m_wizardInfo.DataSourceParser.StorageInfo as ExcelStorageInfo;
                    ExcelStorageInfo dataInfo = new ExcelStorageInfo(sourceInfo.Source);
                    dataInfo.WorkSheetName = sourceInfo.WorkSheetName;
                    dataInfo.RowContainingFieldNames = sourceInfo.RowContainingFieldNames;

                    m_dataStorageInfo = dataInfo;
                    break;

                case DataSourceType.MHT:
                    m_dataStorageInfo = new MHTStorageInfo(m_wizardInfo.DataSourceParser.StorageInfo.Source);
                    var fields = new List<SourceField>();
                    foreach (SourceField field in m_wizardInfo.DataSourceParser.StorageInfo.FieldNames)
                    {
                        fields.Add(field);
                    }
                    m_dataStorageInfo.FieldNames = fields;
                    m_mhtSource = m_wizardInfo.MHTSource;
                    break;

                default:
                    throw new InvalidEnumArgumentException("Invalid Enum Value");
            }
        }
        #endregion
    }
}
