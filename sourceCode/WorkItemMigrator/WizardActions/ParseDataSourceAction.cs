//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Wizard Action Responsible for Parsing the Data Source
    /// </summary>
    class ParseDataSourceAction : WizardAction
    {
        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="wizardInfo"></param>
        public ParseDataSourceAction(WizardInfo wizardInfo)
            : base(wizardInfo)
        {
            Description = Resources.ParseDataSourceAction_Description;
            ActionName = WizardActionName.ParseDataSource;
        }
        #endregion

        #region Overriden methods

        /// <summary>
        /// The Main Working Method
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        protected override WizardActionState DoWork(DoWorkEventArgs e)
        {
            if (m_wizardInfo.DataSourceType != DataSourceType.Excel)
            {
                Message = "Wrong Data Source Type";
                return WizardActionState.Failed;
            }

            try
            {

                m_wizardInfo.Migrator.SourceWorkItems = new List<ISourceWorkItem>();

                var parser = new ExcelParser(m_wizardInfo.DataSourceParser.StorageInfo as ExcelStorageInfo);

                parser.ParseDataSourceFieldNames();

                parser.FieldNameToFields = m_wizardInfo.Migrator.SourceNameToFieldMapping;

                ISourceWorkItem sourceWorkItem = null;
                while ((sourceWorkItem = parser.GetNextWorkItem()) != null)
                {
                    m_wizardInfo.Migrator.SourceWorkItems.Add(sourceWorkItem);
                }

                m_wizardInfo.DataSourceParser.Dispose();
                parser.Dispose();

                // return success
                return WizardActionState.Success;
            }
            catch (WorkItemMigratorException ex)
            {
                // is some exception occured during parsing then set the error Message and return state as Failed
                Message = ex.Args.Title;
                return WizardActionState.Failed;
            }
        }

        #endregion
    }
}