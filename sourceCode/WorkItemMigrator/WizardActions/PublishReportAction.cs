//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    /// <summary>
    /// Wizard Action reponsible for Publixhing the report
    /// </summary>
    internal class PublishReportAction : WizardAction
    {
        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="wizardInfo"></param>
        public PublishReportAction(WizardInfo wizardInfo)
            : base(wizardInfo)
        {
            Description = Resources.PublishReportAction_Description;
            ActionName = WizardActionName.PublishReport;
        }
        #endregion

        #region Overriden methods

        /// <summary>
        /// Main Working Method
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        protected override WizardActionState DoWork(System.ComponentModel.DoWorkEventArgs e)
        {
            if (m_wizardInfo.ResultWorkItems == null || m_wizardInfo.ResultWorkItems.Count == 0)
            {
                Message = "No Workitems are migrated";
                return WizardActionState.Failed;
            }
            else
            {
                foreach (ISourceWorkItem sourceWorkItem in m_wizardInfo.ResultWorkItems)
                {
                    m_wizardInfo.Reporter.AddEntry(sourceWorkItem);
                }
            }
            m_wizardInfo.Reporter.Publish();

            if (m_wizardInfo.IsLinking)
            {
                m_wizardInfo.LinksManager.PublishReport();
            }
            return WizardActionState.Success;
        }

        #endregion
    }
}
