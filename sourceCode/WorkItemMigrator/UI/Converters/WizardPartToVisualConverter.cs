//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.Windows.Data;
    using System.Windows.Media;

    /// <summary>
    /// 
    /// </summary>
    class WizardPartToVisualConverter : IValueConverter
    {
        #region Fields

        private Dictionary<IWizardPart, Visual> m_views = new Dictionary<IWizardPart, Visual>();
        
        #endregion

        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            IWizardPart part = value as IWizardPart;
            if (part == null)
            {
                return null;
            }
            if (m_views.ContainsKey(part))
            {
                return m_views[part];
            }
            else
            {
                Visual visual = GetVisual(part);
                m_views.Add(part, visual);
                return visual;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region private methods

        private Visual GetVisual(IWizardPart part)
        {
            WelcomePart welcomePart = part as WelcomePart;
            if (welcomePart != null)
            {
                return new WelcomeView(welcomePart);
            }

            SelectDataSourcePart selectDataSourcePart = part as SelectDataSourcePart;
            if (selectDataSourcePart != null)
            {
                return new SelectDataSourceView(selectDataSourcePart);
            }

            SelectDestinationServerPart selectDestinationServerPart = part as SelectDestinationServerPart;
            if (selectDestinationServerPart != null)
            {
                return new SelectDestinationServerView(selectDestinationServerPart);
            }

            SettingsFilePart mappingsFilePart = part as SettingsFilePart;
            if (mappingsFilePart != null)
            {
                return new SettingsFileView(mappingsFilePart);
            }

            FieldsSelectionPart fieldsSelectionPart = part as FieldsSelectionPart;
            if (fieldsSelectionPart != null)
            {
                return new FieldsSelectionView(fieldsSelectionPart);
            }

            FieldMappingPart fieldMappingPart = part as FieldMappingPart;
            if (fieldMappingPart != null)
            {
                return new FieldMappingView(fieldMappingPart);
            }

            DataMappingPart dataMappingPart = part as DataMappingPart;
            if (dataMappingPart != null)
            {
                return new DataMappingView(dataMappingPart);
            }

            LinkingPart relationshipsMappingPart = part as LinkingPart;
            if (relationshipsMappingPart != null)
            {
                return new LinkingView(relationshipsMappingPart);
            }

            MiscSettingsPart miscSettingsPart = part as MiscSettingsPart;
            if (miscSettingsPart != null)
            {
                return new MiscSettingsView(miscSettingsPart);
            }

            ConfirmSettingsPart confirmSettingsPart = part as ConfirmSettingsPart;
            if (confirmSettingsPart != null)
            {
                return new ConfirmSettingsView(confirmSettingsPart);
            }

            SummaryPart migrationProgressPart = part as SummaryPart;
            if (migrationProgressPart != null)
            {
                return new SummaryView(migrationProgressPart);
            }

            throw new ArgumentException("Invalid Wizard Part");
        }

        #endregion
    }
}
