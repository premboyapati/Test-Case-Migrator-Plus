//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.IO;

    /// <summary>
    /// Wizard part responsible for showing the different configurations made by user
    /// and let him decide to take Action on those configurations
    /// </summary>
    internal class ConfirmSettingsPart : BaseWizardPart
    {
        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        public ConfirmSettingsPart()
        {
            Header = Resources.ConfirmSettings_Header;
            Description = Resources.ConfirmSettings_Description;
            CanBack = true;
            WizardPage = WizardPage.ConfirmSettings;
        }

        #endregion

        #region public methods

        public override void Reset()
        {
            Warning = null;
            string exePath = Path.Combine(Directory.GetCurrentDirectory(), "TestCaseMigratorPlus.exe");
            string reportDirectory = Path.GetDirectoryName(m_wizardInfo.Reporter.ReportFile);
            string settingsFilePath = m_wizardInfo.OutputSettingsFilePath;
            if (string.IsNullOrEmpty(settingsFilePath))
            {
                settingsFilePath = Path.Combine(reportDirectory, "Settings.xml");
            }

            if (m_wizardInfo.DataSourceType == DataSourceType.Excel)
            {
                ExcelStorageInfo excelInfo = m_wizardInfo.DataSourceParser.StorageInfo as ExcelStorageInfo;
                m_wizardInfo.CommandUsed = String.Format(System.Globalization.CultureInfo.CurrentCulture,
                                                         @"{0} {1} /{2}:""{3}"" /{4}:""{5}"" /{6}:""{7}"" /{8}:""{9}"" /{10}:""{11}"" /{12}:""{13}"" /{14}:""{15}"" /{16}:""{17}""",
                                                         exePath,
                                                         App.ExcelSwitch,
                                                         App.SourceFileCLISwitch,
                                                         excelInfo.Source,
                                                         App.WorkSheetNameCLISwitch,
                                                         excelInfo.WorkSheetName,
                                                         App.HeaderRowCLISwitch,
                                                         excelInfo.RowContainingFieldNames,
                                                         App.CollectionCLISwitch,
                                                         m_wizardInfo.WorkItemGenerator.Server,
                                                         App.ProjectCLISwitch,
                                                         m_wizardInfo.WorkItemGenerator.Project,
                                                         App.WorkItemTypeCLISwitch,
                                                         m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName,
                                                         App.SettingsCLISwitch,
                                                         settingsFilePath,
                                                         App.ReportCLISwitch,
                                                         reportDirectory);
            }
            else if (m_wizardInfo.DataSourceType == DataSourceType.MHT)
            {
                m_wizardInfo.CommandUsed = String.Format(System.Globalization.CultureInfo.CurrentCulture,
                                                         @"{0} {1} /{2}:""{3}"" /{4}:""{5}"" /{6}:""{7}"" /{8}:""{9}"" /{10}:""{11}"" /{12}:""{13}""",
                                                         exePath,
                                                         App.MHTSwitch,
                                                         App.SourceFileCLISwitch,
                                                         m_wizardInfo.MHTSource,
                                                         App.CollectionCLISwitch,
                                                         m_wizardInfo.WorkItemGenerator.Server,
                                                         App.ProjectCLISwitch,
                                                         m_wizardInfo.WorkItemGenerator.Project,
                                                         App.WorkItemTypeCLISwitch,
                                                         m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName,
                                                         App.SettingsCLISwitch,
                                                         settingsFilePath,
                                                         App.ReportCLISwitch,
                                                         reportDirectory);
            }
        }

        public override bool UpdateWizardPart()
        {
            return true;
        }

        #endregion

        #region protected/private methods

        protected override bool IsInitializationRequired(WizardInfo state)
        {
            return true;
        }

        public override bool ValidatePartState()
        {
            Warning = null;
            return true;
        }

        /// <summary>
        /// If all mandatory fields are maped in Field Mapping only then we can initialize and show this
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
        protected override bool CanInitializeWizardPage(WizardInfo info)
        {
            if (info.WorkItemGenerator == null || info.WorkItemGenerator.TfsNameToFieldMapping == null)
            {
                return false;
            }

            m_canShow = true;

            foreach (IWorkItemField field in info.WorkItemGenerator.TfsNameToFieldMapping.Values)
            {
                if (field.IsMandatory && string.IsNullOrEmpty(field.SourceName))
                {
                    Warning = Resources.MandatoryFieldsNotMappedErrorTitle;
                    m_canShow = false;
                    break;
                }
            }

            if (info.IsLinking && string.IsNullOrEmpty(info.DataSourceParser.StorageInfo.SourceIdFieldName))
            {
                Warning = "Source Id is not specified for Linking";
                m_canShow = false;
            }
            return m_canShow;
        }

        #endregion
    }
}
