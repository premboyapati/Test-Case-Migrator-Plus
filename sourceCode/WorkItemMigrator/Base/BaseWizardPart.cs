//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    /// <summary>
    /// Base class for All wizard Parts
    /// </summary>
    internal abstract class BaseWizardPart : NotifyPropertyChange, IWizardPart
    {
        #region Fields

        /// <summary>
        /// Member variables needed for data binding
        /// </summary>
        private string m_header;
        private string m_description;
        private bool m_canBack;
        protected bool m_canNext;
        private string m_warning;
        protected bool m_canShow;
        private bool m_isActiveWizardPart;

        // Used to validate and keep Wizard part's state in synch with other wizard parts at load and update time
        protected WizardPartPrerequisite m_prerequisite;

        // WizardInfo
        protected WizardInfo m_wizardInfo;

        #endregion

        #region Constants

        // Needed for the notifications
        public const string CanConfirmPropertyName = "CanConfirm";
        public const string HeaderPropertyName = "Header";
        public const string DescriptionPropertyName = "Description";
        public const string CanBackPropertyName = "CanBack";
        public const string CanNextPropertyName = "CanNext";

        #endregion

        #region Properties

        /// <summary>
        /// Header of a Wizard Page
        /// </summary>
        public string Header
        {
            get
            {
                return m_header;
            }
            protected set
            {
                m_header = value;
                NotifyPropertyChanged(HeaderPropertyName);
            }
        }

        /// <summary>
        /// Description of the Wizard page
        /// </summary>
        public string Description
        {
            get
            {
                return m_description;
            }
            protected set
            {
                m_description = value;
                NotifyPropertyChanged(DescriptionPropertyName);
            }
        }


        public string Warning
        {
            get
            {
                return m_warning;
            }
            set
            {
                if (value == null)
                {
                    value = string.Empty;
                }
                m_warning = value;
                NotifyPropertyChanged("Warning");
            }
        }

        /// <summary>
        /// Is Back button of Wizard Enabled
        /// </summary>
        public bool CanBack
        {
            get
            {
                return m_canBack;
            }
            protected set
            {
                m_canBack = value;
                NotifyPropertyChanged(CanBackPropertyName);
            }
        }

        /// <summary>
        /// Is Next Button of Wizard enabled
        /// </summary>
        public virtual bool CanNext
        {
            get
            {
                return m_canNext;
            }
            set
            {
                if (m_canNext != value)
                {
                    m_canNext = value;
                    NotifyPropertyChanged(CanNextPropertyName);
                }
            }
        }

        /// <summary>
        /// Whether we can show this WizardPage
        /// </summary>
        public bool CanShow
        {
            get
            {
                return m_canShow;
            }
            set
            {
                m_canShow = value;
                NotifyPropertyChanged("CanShow");
            }
        }

        /// <summary>
        /// Checks that we can initialize Wizard Part or not.
        /// </summary>
        public bool CanInitialize
        {
            get
            {
                if (m_wizardInfo == null)
                {
                    return false;
                }

                CanInitializeWizardPage(m_wizardInfo);
                NotifyPropertyChanged("CanShow");
                return CanShow;
            }
        }

        /// <summary>
        /// The State of Wizard Page. It is unique for each Wizard Page.
        /// </summary>
        public WizardPage WizardPage
        {
            get;
            protected set;
        }

        /// <summary>
        /// Is this Wizard Page in active and currently in process?
        /// </summary>
        public bool IsActiveWizardPart
        {
            get
            {
                return m_isActiveWizardPart;
            }
            set
            {
                m_isActiveWizardPart = value;
                NotifyPropertyChanged("IsActiveWizardPart");
            }
        }

        /// <summary>
        /// Contains all user settings
        /// </summary>
        public WizardInfo WizardInfo
        {
            get
            {
                return m_wizardInfo;
            }
        }

        #endregion

        #region abstract and virtual methods

        /// <summary>
        /// Initializes the Wizard Part and Send Notifications
        /// </summary>
        /// <param name="info"></param>
        public virtual void Initialize(WizardInfo info)
        {
            if (!CanInitializeWizardPage(info))
            {
                return;
            }


            if (IsInitializationRequired(info))
            {
                m_wizardInfo = info;

                if (m_prerequisite == null)
                {
                    m_prerequisite = new WizardPartPrerequisite(info);
                }

                try
                {
                    Reset();
                }
                catch (WorkItemMigratorException ex)
                {
                    Warning = ex.Args.Title;
                    CanShow = false;
                    m_wizardInfo = null;
                }
            }
            m_prerequisite.Save();
            FireStateNotifications();
        }


        public virtual void Clear()
        {
            m_wizardInfo = null;
            m_prerequisite = null;
        }

        /// <summary>
        /// Abstract Method : Resets the wizard part
        /// </summary>
        public abstract void Reset();


        /// <summary>
        /// Abstract Method : Updates the Wizard info with the current Wizard State.
        /// </summary>
        /// <returns>Whether Updation was successful?</returns>
        public abstract bool UpdateWizardPart();


        /// <summary>
        /// Used by Initialize to check whether Initialization is required or not
        /// </summary>
        /// <param name="state"></param>
        /// <returns></returns>
        protected virtual bool IsInitializationRequired(WizardInfo info)
        {
            return (m_wizardInfo == null);
        }

        /// <summary>
        /// Validates whether the Wizard part State is Valid or not
        /// </summary>
        /// <returns></returns>
        public abstract bool ValidatePartState();

        protected virtual bool IsUpdationRequired()
        {
            return true;
        }

        protected abstract bool CanInitializeWizardPage(WizardInfo info);

        /// <summary>
        /// Send The Notifications for Data Binding
        /// </summary>
        public virtual void FireStateNotifications()
        {
            NotifyPropertyChanged(HeaderPropertyName);
            NotifyPropertyChanged(DescriptionPropertyName);
            NotifyPropertyChanged(CanBackPropertyName);
            NotifyPropertyChanged(CanConfirmPropertyName);
            NotifyPropertyChanged("CanShow");
            NotifyPropertyChanged("Warning");
            NotifyPropertyChanged("WizardInfo");
            CanNext = ValidatePartState();
        }

        #endregion
    }
}
