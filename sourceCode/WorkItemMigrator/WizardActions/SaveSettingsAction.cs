//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    /// <summary>
    /// Wizard Action responsible for Saving Settings
    /// </summary>
    internal class SaveSettingsAction : WizardAction
    {
        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="wizardInfo"></param>
        public SaveSettingsAction(WizardInfo wizardInfo)
            : base(wizardInfo)
        {
            Description = Resources.SaveSettingsAction_Description;
            ActionName = WizardActionName.SaveSettings;
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
            m_wizardInfo.SaveSettings(m_wizardInfo.OutputSettingsFilePath);
            return WizardActionState.Success;
        }

        #endregion
    }
}
