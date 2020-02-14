//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    /// <summary>
    /// Wizard Part that displays the Welcome Screen
    /// </summary>
    class WelcomePart : BaseWizardPart
    {
        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        public WelcomePart()
        {
            Header = Resources.WelcomePageHeader;
            CanBack = false;
            WizardPage = WizardPage.Welcome;
        }

        #endregion

        #region Overriden Methods

        public override void Reset()
        { }

        public override bool UpdateWizardPart()
        {
            return true;
        }

        public override bool ValidatePartState()
        {
            return true;
        }

        protected override bool CanInitializeWizardPage(WizardInfo info)
        {
            m_canShow = true;
            return m_canShow;
        }
        #endregion
    }
}