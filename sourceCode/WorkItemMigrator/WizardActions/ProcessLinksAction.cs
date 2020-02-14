//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    internal class ProcessLinksAction : WizardAction
    {
        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="wizardInfo"></param>
        public ProcessLinksAction(WizardInfo wizardInfo)
            : base(wizardInfo)
        {
            Description = "Process Links";
            ActionName = WizardActionName.Relationships;
        }

        #endregion

        #region Overriden Methods

        /// <summary>
        /// Main Working Method
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        protected override WizardActionState DoWork(System.ComponentModel.DoWorkEventArgs e)
        {
            foreach (var sourceWorkItem in m_wizardInfo.ResultWorkItems)
            {
                PassedSourceWorkItem passedWorkItem = sourceWorkItem as PassedSourceWorkItem;
                if (passedWorkItem != null)
                {
                    foreach (Link link in passedWorkItem.Links)
                    {
                        m_wizardInfo.LinksManager.AddLink(link);
                    }
                }
            }
            m_wizardInfo.LinksManager.Save();
            return WizardActionState.Success;
        }

        #endregion

    }
}
