//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;

    internal class LinkingPart : BaseWizardPart
    {
        #region Fields
        private string m_sourceIdField;       
        private bool m_isLinking;
        private List<string> m_fieldsPool;

        #endregion

        #region Constructor

        public LinkingPart()
        {
            Header = "Links mapping";
            Description = "Specify the links between work items that need be migrated to TFS. The link related mappings are persisted across multiple sessions. If you want to have any links (in current or future invocations of this tool), make sure to select the checkbox and specify appropriate fields.";
            CanBack = true;
            CanNext = true;
            WizardPage = WizardPage.Linking;
            LinkingRows = new ObservableCollection<LinkingRow>();
            SourceIdAvailableFields = new ObservableCollection<string>();
            m_fieldsPool = new List<string>();
            IsLinking = true;
        }

        #endregion

        #region Properties

        public override bool CanNext
        {
            get
            {
                return m_canNext;
            }
            set
            {
                if (m_canNext != value && WizardInfo != null)
                {
                    m_canNext = value;
                    WizardInfo.CanConfirm = m_canNext;
                    NotifyPropertyChanged(BaseWizardPart.CanNextPropertyName);
                }
            }
        }

        public bool IsLinking
        {
            get
            {
                return m_isLinking;
            }
            set
            {
                m_isLinking = value;
                Reset();
                CanNext = ValidatePartState();

                NotifyPropertyChanged("IsLinking");
            }
        }


        public string SourceIdField
        {
            get
            {
                return m_sourceIdField;
            }
            set
            {
                string previousSourceIdField = m_sourceIdField;
                m_sourceIdField = value;

                RemoveFieldFromPool(m_sourceIdField);
                AddFieldIntoPool(previousSourceIdField);
                NotifyPropertyChanged("SourceIdField");
                CanNext = ValidatePartState();
            }
        }

        public ObservableCollection<string> SourceIdAvailableFields
        {
            get;
            private set;
        }

        public ObservableCollection<LinkingRow> LinkingRows
        {
            get;
            private set;
        }

        #endregion

        #region Public Methods

        public override void Reset()
        {
            if (!IsLinking || m_wizardInfo == null)
            {
                return;
            }

            if (WizardInfo.DataSourceType == DataSourceType.MHT)
            {
                IsLinking = false;
            }

            m_fieldsPool.Clear();
            ClearSourceIdAvailableFields();

            foreach (var field in m_wizardInfo.DataSourceParser.StorageInfo.FieldNames)
            {
                if (!m_wizardInfo.Migrator.SourceNameToFieldMapping.ContainsKey(field.FieldName) &&
                    String.CompareOrdinal(field.FieldName, WizardInfo.RelationshipsInfo.TestSuiteField) != 0)
                {
                    m_fieldsPool.Add(field.FieldName);
                }
            }

            ClearLinkingRows();

            LinkingRow row1 = null;
            LinkingRow row2 = null;
            LinkingRow row3 = null;
            LinkingRow row4 = null;

            switch (WizardInfo.WorkItemGenerator.WorkItemCategory)
            {
                case WorkItemGenerator.TestCaseCategory:
                    row1 = new LinkingRow(this,
                                          WorkItemGenerator.TestsLinkReferenceName,
                                          WizardInfo.WorkItemGenerator.WorkItemCategoryToDefaultType[WorkItemGenerator.RequirementCategory],
                                          m_fieldsPool,
                                          null);
                    AddLinkingRow(row1);

                    row2 = new LinkingRow(this,
                                          WorkItemGenerator.TestsLinkReferenceName,
                                          WizardInfo.WorkItemGenerator.WorkItemCategoryToDefaultType[WorkItemGenerator.BugCategory],
                                          m_fieldsPool,
                                          null);
                    AddLinkingRow(row2);

                    row3 = new LinkingRow(this,
                                          WorkItemGenerator.RelatedLinkReferenceName,
                                          WizardInfo.WorkItemGenerator.SelectedWorkItemTypeName,
                                          m_fieldsPool,
                                          null);
                    AddLinkingRow(row3);
                    break;

                case WorkItemGenerator.BugCategory:
                    row1 = new LinkingRow(this,
                                          WorkItemGenerator.TestedByLinkRefernceName,
                                          WizardInfo.WorkItemGenerator.WorkItemCategoryToDefaultType[WorkItemGenerator.TestCaseCategory],
                                          m_fieldsPool,
                                          null);
                    AddLinkingRow(row1);

                    row2 = new LinkingRow(this,
                                          WorkItemGenerator.RelatedLinkReferenceName,
                                          WizardInfo.WorkItemGenerator.SelectedWorkItemTypeName,
                                          m_fieldsPool,
                                          null);
                    AddLinkingRow(row2);

                    row3 = new LinkingRow(this,
                                          WorkItemGenerator.RelatedLinkReferenceName,
                                          WizardInfo.WorkItemGenerator.WorkItemCategoryToDefaultType[WorkItemGenerator.RequirementCategory],
                                          m_fieldsPool,
                                          null);
                    AddLinkingRow(row3);
                    break;

                case WorkItemGenerator.RequirementCategory:
                    row1 = new LinkingRow(this,
                                          WorkItemGenerator.TestedByLinkRefernceName,
                                          WizardInfo.WorkItemGenerator.WorkItemCategoryToDefaultType[WorkItemGenerator.TestCaseCategory],
                                          m_fieldsPool,
                                          null);
                    AddLinkingRow(row1);

                    row2 = new LinkingRow(this,
                                          WorkItemGenerator.ParentLinkReferenceName,
                                          WizardInfo.WorkItemGenerator.SelectedWorkItemTypeName,
                                          m_fieldsPool,
                                          null);
                    AddLinkingRow(row2);

                    row3 = new LinkingRow(this,
                                          WorkItemGenerator.RelatedLinkReferenceName,
                                          WizardInfo.WorkItemGenerator.SelectedWorkItemTypeName,
                                          m_fieldsPool,
                                          null);
                    AddLinkingRow(row3);

                    row4 = new LinkingRow(this,
                                          WorkItemGenerator.RelatedLinkReferenceName,
                                          WizardInfo.WorkItemGenerator.WorkItemCategoryToDefaultType[WorkItemGenerator.BugCategory],
                                          m_fieldsPool,
                                          null);
                    AddLinkingRow(row4);
                    break;

                default:
                    break;
            }

            ResetSourceIdField();

            string startCategory = WizardInfo.WorkItemGenerator.WorkItemTypeToCategoryMapping[WizardInfo.WorkItemGenerator.SelectedWorkItemTypeName];

            foreach (var rule in WizardInfo.RelationshipsInfo.LinkRules)
            {
                if (String.CompareOrdinal(startCategory, rule.StartWorkItemCategory) == 0 &&
                        WizardInfo.WorkItemGenerator.LinkTypeNames.Contains(rule.LinkTypeReferenceName) &&
                        IsValidCategory(rule.EndWorkItemCategory) &&
                        m_fieldsPool.Contains(rule.SourceFieldNameOfEndWorkItemCategory))
                {
                    bool isRuleAlreadyAdded = false;
                    foreach (var row in LinkingRows)
                    {
                        string endCategory = WizardInfo.WorkItemGenerator.WorkItemTypeToCategoryMapping[row.EndWorkItemType];

                        if (String.CompareOrdinal(row.LinkRefernceName, rule.LinkTypeReferenceName) == 0 &&
                            String.CompareOrdinal(endCategory, rule.EndWorkItemCategory) == 0 &&
                            row.AvailableFields.Contains(rule.SourceFieldNameOfEndWorkItemCategory))
                        {
                            isRuleAlreadyAdded = true;
                            break;
                        }
                    }
                    if (!isRuleAlreadyAdded &&
                        WizardInfo.WorkItemGenerator.WorkItemCategoryToDefaultType.ContainsKey(rule.EndWorkItemCategory))
                    {
                        AddLinkingRow(new LinkingRow(this,
                                                    rule.LinkTypeReferenceName,
                                                    WizardInfo.WorkItemGenerator.WorkItemCategoryToDefaultType[rule.EndWorkItemCategory],
                                                    m_fieldsPool,
                                                    rule.Description));
                    }
                }
            }

            foreach (var row in LinkingRows)
            {
                foreach (var rule in WizardInfo.RelationshipsInfo.LinkRules)
                {
                    string endCategory = WizardInfo.WorkItemGenerator.WorkItemTypeToCategoryMapping[row.EndWorkItemType];

                    if (String.CompareOrdinal(startCategory, rule.StartWorkItemCategory) == 0 &&
                        String.CompareOrdinal(row.LinkRefernceName, rule.LinkTypeReferenceName) == 0 &&
                        String.CompareOrdinal(endCategory, rule.EndWorkItemCategory) == 0 &&
                        row.AvailableFields.Contains(rule.SourceFieldNameOfEndWorkItemCategory))
                    {
                        row.LinkedField = rule.SourceFieldNameOfEndWorkItemCategory;
                    }
                }
            }
        }

        private bool IsValidCategory(string category)
        {
            return category == WorkItemGenerator.TestCaseCategory ||
                   category == WorkItemGenerator.BugCategory ||
                   category == WorkItemGenerator.RequirementCategory;
        }

        private void ResetSourceIdField()
        {
            ClearSourceIdAvailableFields();
            foreach (var field in m_fieldsPool)
            {
                AddSourceIdAvailableField(field);
            }
            SourceIdField = Resources.SelectPlaceholder;

            foreach (var kvp in WizardInfo.Migrator.SourceNameToFieldMapping)
            {
                if (kvp.Value.IsIdField)
                {
                    AddSourceIdAvailableField(kvp.Key);
                    SourceIdField = kvp.Key;
                    break;
                }
            }

            if (SourceIdAvailableFields.Contains(m_wizardInfo.RelationshipsInfo.SourceIdField))
            {
                SourceIdField = m_wizardInfo.RelationshipsInfo.SourceIdField;
            }
        }

        public override bool UpdateWizardPart()
        {
            if (WizardInfo.DataSourceType == DataSourceType.MHT)
            {
                m_wizardInfo.IsLinking = false;
                return true;
            }

            if (!ValidatePartState())
            {
                return false;
            }

            m_wizardInfo.IsLinking = IsLinking;
            m_wizardInfo.DataSourceParser.StorageInfo.SourceIdFieldName = null;
            m_wizardInfo.DataSourceParser.StorageInfo.LinkRules.Clear();

            if (!IsLinking)
            {
                return true;
            }

            if (String.CompareOrdinal(SourceIdField, Resources.SelectPlaceholder) != 0)
            {
                m_wizardInfo.DataSourceParser.StorageInfo.SourceIdFieldName = SourceIdField;
                m_wizardInfo.RelationshipsInfo.SourceIdField = SourceIdField;
            }

            string witType = m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName;
            foreach (var row in LinkingRows)
            {
                ILinkRule linkRule = row.LinkRule;
                if (linkRule != null)
                {
                    m_wizardInfo.RelationshipsInfo.LinkRules.Add(linkRule);
                    m_wizardInfo.DataSourceParser.StorageInfo.LinkRules.Add(linkRule);
                }
            }
            return true;
        }

        #endregion

        #region Private/Protected  Methods

        private void RemoveFieldFromPool(string fieldName)
        {
            if (m_fieldsPool.Contains(fieldName) &&
                !WizardInfo.WorkItemGenerator.SourceNameToFieldMapping.ContainsKey(fieldName))
            {
                m_fieldsPool.Remove(fieldName);
                if (String.CompareOrdinal(fieldName, SourceIdField) != 0)
                {
                    RemoveSourceIdAvailableField(fieldName);
                }
                foreach (var row in LinkingRows)
                {
                    row.RemoveAvailableField(fieldName);
                }
            }
        }

        private void AddFieldIntoPool(string fieldName)
        {
            if (!string.IsNullOrEmpty(fieldName) &&
                !m_fieldsPool.Contains(fieldName) &&
                String.CompareOrdinal(fieldName, Resources.SelectPlaceholder) != 0)
            {
                m_fieldsPool.Add(fieldName);

                AddSourceIdAvailableField(fieldName);

                foreach (var row in LinkingRows)
                {
                    row.AddAvailableField(fieldName);
                }
            }
        }

        public override bool ValidatePartState()
        {
            Warning = null;
            if (WizardInfo != null && WizardInfo.DataSourceType == DataSourceType.MHT)
            {
                return true;
            }


            if (IsLinking &&
                String.CompareOrdinal(SourceIdField, Resources.SelectPlaceholder) == 0)
            {
                Warning = "Source Id field is not specified";
                return false;
            }
            return true;
        }

        protected override bool CanInitializeWizardPage(WizardInfo info)
        {
            m_canShow = true;

            if (info.Migrator.SourceNameToFieldMapping == null || info.Migrator.FieldToUniqueValues == null)
            {
                m_canShow = false;
                Warning = Resources.MandatoryFieldsNotMappedErrorTitle;
                m_canShow = false;
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
            return m_canShow;

        }

        /// <summary>
        /// Is Initialization of Data Mapping Wizard Part required?
        /// </summary>
        /// <param name="state"></param>
        /// <returns></returns>
        protected override bool IsInitializationRequired(WizardInfo state)
        {
            if (m_wizardInfo == null ||
                m_prerequisite.IsDataSourceChanged() ||
                m_prerequisite.IsFieldMappingChanged() ||
                m_prerequisite.IsSettingsFilePathChanged() ||
                m_prerequisite.IsServerConnectionChanged())
            {
                return true;
            }
            return false;
        }


        private void ClearSourceIdAvailableFields()
        {
            App.CallMethodInUISynchronizationContext(ClearSourceIdAvailableFieldsInUIContext, null);
        }

        private void ClearSourceIdAvailableFieldsInUIContext(object obj)
        {
            SourceIdAvailableFields.Clear();
            SourceIdAvailableFields.Add(Resources.SelectPlaceholder);
        }

        private void AddSourceIdAvailableField(string fieldName)
        {
            App.CallMethodInUISynchronizationContext(AddSourceIdAvailableFieldInUIContext, fieldName);
        }

        private void AddSourceIdAvailableFieldInUIContext(object obj)
        {
            string fieldName = obj as string;
            if (!string.IsNullOrEmpty(fieldName) &&
                !SourceIdAvailableFields.Contains(fieldName))
            {
                SourceIdAvailableFields.Add(fieldName);
            }
        }

        private void RemoveSourceIdAvailableField(string fieldName)
        {
            App.CallMethodInUISynchronizationContext(RemoveSourceIdAvailableFieldInUIContext, fieldName);
        }

        private void RemoveSourceIdAvailableFieldInUIContext(object obj)
        {
            string fieldName = obj as string;
            if (!string.IsNullOrEmpty(fieldName) && SourceIdAvailableFields.Contains(fieldName))
            {
                SourceIdAvailableFields.Remove(fieldName);
            }
        }

        private void ClearLinkingRows()
        {
            App.CallMethodInUISynchronizationContext(ClearLinkingRowsInUIContext, null);
        }

        private void ClearLinkingRowsInUIContext(object obj)
        {
            LinkingRows.Clear();
        }

        private void AddLinkingRow(LinkingRow row)
        {
            App.CallMethodInUISynchronizationContext(AddLinkingRowInUIContext, row);
        }

        private void AddLinkingRowInUIContext(object obj)
        {
            LinkingRow row = obj as LinkingRow;
            if (row != null)
            {
                row.PropertyChanged += LinkingRow_PropertyChanged;
                LinkingRows.Add(row);
            }
        }

        private void LinkingRow_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            LinkingRow row = sender as LinkingRow;
            AddFieldIntoPool(row.PreviousLinkedField);
            RemoveFieldFromPool(row.LinkedField);
        }

        #endregion
    }

    internal class LinkingRow : NotifyPropertyChange
    {
        #region Fields

        private string m_linkedField;
        private LinkingPart m_part;
        private string m_description;
        private string m_startWorkItemType;
        private string m_endWorkItemType;
        private string m_waterMark;

        #endregion

        #region Constants

        public const string SourceFieldPropertyName = "SourceField";

        #endregion

        #region Static Members

        private static IDictionary<string, string> s_LinkReferenceNameToLinkName;

        static LinkingRow()
        {
            s_LinkReferenceNameToLinkName = new Dictionary<string, string>();
            s_LinkReferenceNameToLinkName.Add(WorkItemGenerator.TestsLinkReferenceName, " \"Tests\" ");
            s_LinkReferenceNameToLinkName.Add(WorkItemGenerator.TestedByLinkRefernceName, " is \"Tested by\" ");
            s_LinkReferenceNameToLinkName.Add(WorkItemGenerator.RelatedLinkReferenceName, " is \"Related to\" ");
            s_LinkReferenceNameToLinkName.Add(WorkItemGenerator.ParentLinkReferenceName, " has \"Parent\" ");
        }
        #endregion

        #region Constructor

        public LinkingRow(LinkingPart part, string linkReferenceName, string endWorkItemTypeName, List<string> availableFields, string description)
        {
            m_part = part;
            m_startWorkItemType = m_part.WizardInfo.WorkItemGenerator.SelectedWorkItemTypeName;
            m_endWorkItemType = endWorkItemTypeName;
            StartWorkItemCategoryName = GetCategoryNameFromWorkItemType(m_startWorkItemType);
            LinkRefernceName = linkReferenceName;
            EndWorkItemCategoryName = GetCategoryNameFromWorkItemType(m_endWorkItemType);
            AvailableFields = new ObservableCollection<string>();
            AddAvailableField(Resources.SelectPlaceholder);
            foreach (string field in availableFields)
            {
                AddAvailableField(field);
            }
            LinkedField = Resources.SelectPlaceholder;
            Description = description;
        }

        private string GetCategoryNameFromWorkItemType(string workItemType)
        {
            switch (m_part.WizardInfo.WorkItemGenerator.WorkItemTypeToCategoryMapping[workItemType])
            {
                case WorkItemGenerator.TestCaseCategory:
                    return "Test Case";

                case WorkItemGenerator.BugCategory:
                    return "Bug";

                case WorkItemGenerator.RequirementCategory:
                    return "User Story";

                default:
                    throw new ArgumentException("invalid work item type");
            }
        }

        #endregion

        #region Properties

        public string StartWorkItemCategoryName
        {
            get;
            private set;
        }

        public string LinkName
        {
            get
            {
                if (s_LinkReferenceNameToLinkName.ContainsKey(LinkRefernceName))
                {
                    return s_LinkReferenceNameToLinkName[LinkRefernceName];
                }
                else
                {
                    return "\"" + LinkRefernceName + "\"";
                }
            }
        }

        public string Description
        {
            get
            {
                return m_description;
            }
            set
            {
                m_description = value;
                NotifyPropertyChanged("Description");
            }
        }

        public string LinkRefernceName
        {
            get;
            private set;
        }

        public string EndWorkItemCategoryName
        {
            get;
            private set;
        }

        public string EndWorkItemType
        {
            get
            {
                return m_endWorkItemType;
            }
        }

        public string LinkedField
        {
            get
            {
                return m_linkedField;
            }
            set
            {
                PreviousLinkedField = m_linkedField;
                m_linkedField = value;
                NotifyPropertyChanged("LinkedField");
            }
        }

        public string PreviousLinkedField
        {
            get;
            private set;
        }

        public ILinkRule LinkRule
        {
            get
            {
                if (!string.IsNullOrEmpty(LinkedField) &&
                    string.CompareOrdinal(LinkedField, Resources.SelectPlaceholder) != 0)
                {
                    string startCategory = m_part.WizardInfo.WorkItemGenerator.WorkItemTypeToCategoryMapping[m_startWorkItemType];
                    string endCategory = m_part.WizardInfo.WorkItemGenerator.WorkItemTypeToCategoryMapping[m_endWorkItemType];

                    return new LinkRule(startCategory,
                                        LinkedField,
                                        LinkRefernceName,
                                        endCategory,
                                        Description);
                }
                else
                {
                    return null;
                }
            }
        }

        public ObservableCollection<string> AvailableFields
        {
            get;
            private set;
        }

        public string WaterMark
        {
            get
            {
                return m_waterMark;
            }
            set
            {
                m_waterMark = value;
                NotifyPropertyChanged("WaterMark");
            }
        }

        #endregion

        #region Public methods

        public void RemoveAvailableField(string fieldName)
        {
            if (String.CompareOrdinal(fieldName, LinkedField) != 0 &&
                AvailableFields.Contains(fieldName))
            {
                App.CallMethodInUISynchronizationContext(RemoveAvailableLinkedFieldsInUIContext, fieldName);
            }
        }

        public void AddAvailableField(string fieldName)
        {
            if (!AvailableFields.Contains(fieldName))
            {
                App.CallMethodInUISynchronizationContext(AddAvailableLinkedFieldsInUIContext, fieldName);
            }
        }

        #endregion

        #region Private Methods

        private void AddAvailableLinkedFieldsInUIContext(object obj)
        {
            string fieldName = obj as string;
            if (!string.IsNullOrEmpty(fieldName))
            {
                AvailableFields.Add(fieldName);
            }
        }

        private void RemoveAvailableLinkedFieldsInUIContext(object obj)
        {
            string fieldName = obj as string;
            if (!string.IsNullOrEmpty(fieldName) && AvailableFields.Contains(fieldName))
            {
                AvailableFields.Remove(fieldName);
            }
        }

        #endregion
    }
}
