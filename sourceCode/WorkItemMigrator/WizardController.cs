//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Diagnostics;

    /// <summary>
    /// 
    /// Wizard Class reponsible for Navigation between Different Wizard Pages.
    /// 
    /// </summary>
    internal class WizardController : NotifyPropertyChange, IDisposable
    {
        #region Fields

        // Wizard Part responsible for all operation related to a wizard Page.
        private IWizardPart m_wizardPart;

        // All Information collected in the wizard needed for migration of the worktems.
        private WizardInfo m_wizardInfo;

        // Wizard State to Wizard Part Hash Table
        private IDictionary<WizardPage, IWizardPart> m_wizardPageHashTable;

        // An ordered list of Wizard States needed for sequential navgation of Wizard Pages
        private IList<WizardPage> m_wizardPages;

        private bool m_isMigrationCompleted;

        private string m_currentDirectory;

        private WizardPage m_loadWizardPage;

        protected BackgroundWorker m_worker;

        private string m_title;

        private bool m_canStartAgain;

        #endregion

        #region Constants

        public const string IsMigrationCompletedTag = "IsMigrationCompleted";
        private const int FieldsWizardPartPositionInWizardFlow = 4;
        private const int LinkingWizardPartPositionInWizardFlow = 6;

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor to initialize Wizard Info, Wizard Page Hash table and List of Wizard States.
        /// Also loads the first wizard page
        /// </summary>
        public WizardController()
        {
            m_currentDirectory = System.IO.Directory.GetCurrentDirectory();
            Reset();
        }

        #endregion

        #region Properties

        public bool CanStartAgain
        {
            get
            {
                return m_canStartAgain;
            }
            set
            {
                m_canStartAgain = value;
                NotifyPropertyChanged("CanStartAgain");
            }
        }

        public string Title
        {
            get
            {
                return m_title;
            }
            set
            {
                m_title = value;
                NotifyPropertyChanged("Title");
            }
        }

        /// <summary>
        /// The index of Active Wizard part in list of Wizard Parts
        /// </summary>
        private int CurrentWizardIndex
        {
            get
            {
                return WizardPart == null ? 0 : m_wizardPages.IndexOf(WizardPart.WizardPage);
            }
        }

        /// <summary>
        /// Collection of Wizard Parts
        /// </summary>
        public ObservableCollection<IWizardPart> WizardParts
        {
            get;
            set;
        }

        /// <summary>
        /// Needed by The UI to know the Wizard Page to show.
        /// </summary>
        public IWizardPart WizardPart
        {
            get
            {
                return m_wizardPart;
            }
            private set
            {
                if (m_wizardPart != null)
                {
                    m_wizardPart.IsActiveWizardPart = false;
                }

                value.IsActiveWizardPart = true;
                m_wizardPart = value;
                NotifyPropertyChanged("WizardPart");
            }
        }

        /// <summary>
        /// Wizard Information
        /// </summary>
        public WizardInfo WizardInfo
        {
            get
            {
                return m_wizardInfo;
            }
            private set
            {
                m_wizardInfo = value;
                NotifyPropertyChanged("WizardInfo");
            }
        }

        /// <summary>
        /// Is Wizard Part ressponsibel for Wizard Actions is active?
        /// </summary>
        public bool IsSummaryPage
        {
            get
            {
                return WizardPart != null && WizardPart.WizardPage == WizardPage.Summary;
            }
        }

        /// <summary>
        /// Is confirm Wizard Configuration Wizard part Active?
        /// </summary>
        public bool IsConfirmPage
        {
            get
            {
                return WizardPart != null && WizardPart.WizardPage == WizardPage.ConfirmSettings;
            }
        }

        /// <summary>
        /// Is Confirm Button visible
        /// </summary>
        public bool IsConfirmVisible
        {
            get
            {
                return !(IsSummaryPage || IsConfirmPage);
            }
        }

        /// <summary>
        /// Are Wizard Actions Done?
        /// </summary>
        public bool IsMigrationCompleted
        {
            get
            {
                return m_isMigrationCompleted;
            }
            set
            {
                m_isMigrationCompleted = value;
                if (m_isMigrationCompleted)
                {
                    CanStartAgain = true;
                }
                NotifyPropertyChanged(IsMigrationCompletedTag);
            }
        }

        public bool IsLoadingWizardPage
        {
            get;
            private set;
        }

        #endregion

        #region public methods

        /// <summary>
        /// Shows the Help
        /// </summary>
        public void ShowHelp()
        {
            try
            {
                Process.Start(m_currentDirectory + "\\TestCaseMigratorPlus(Excel-MHT)-Readme.doc");
            }
            catch (Win32Exception)
            {
                WizardPart.Warning = "Unable to show help document located at" + m_currentDirectory + "\\TestCaseMigratorPlus(Excel-MHT)-Readme.doc. Please verify that Help document exists there.";
            }
        }

        /// <summary>
        /// Load the Previous Wizard Part
        /// </summary>
        public void GoBack()
        {
            if (CurrentWizardIndex > 0)
            {
                LoadWizardPart(m_wizardPages[CurrentWizardIndex - 1]);
            }
        }

        /// <summary>
        /// Loads the Next Wizard Part
        /// </summary>
        public void GoNext()
        {
            LoadWizardPart(m_wizardPages[CurrentWizardIndex + 1]);
        }

        /// <summary>
        /// Finishes the Wizard and Loads the last Wizard Part which is also responsible for the migration of Workitems
        /// </summary>
        public void SaveAndMigrate()
        {
            LoadWizardPart(WizardPage.Summary);
        }

        /// <summary>
        /// Save Wizard State
        /// </summary>
        public void Save()
        {
            using (new AutoWaitCursor())
            {
                WizardAction saveSettings = new SaveSettingsAction(m_wizardInfo);
                saveSettings.Start();
            }
        }

        /// <summary>
        /// Load Confirm Settings Wizard part
        /// </summary>
        public void LoadConfirmPage()
        {
            LoadWizardPart(WizardPage.ConfirmSettings);
        }

        /// <summary>
        /// Load Wizard Part
        /// </summary>
        /// <param name="part"></param>
        public void LoadWizardPart(IWizardPart part)
        {
            LoadWizardPart(part.WizardPage);
        }

        /// <summary>
        /// Stop Wizard Actions in progress
        /// </summary>
        public void StopWizardActions()
        {
            SummaryPart part = WizardPart as SummaryPart;
            if (part.IsActiveWizardPart)
            {
                part.StopAll();
            }
        }

        public void Reset()
        {
            if (WizardInfo != null)
            {
                foreach (IWizardPart part in m_wizardPageHashTable.Values)
                {
                    part.Clear();
                }

                WizardInfo.PropertyChanged -= new PropertyChangedEventHandler(WizardInfo_PropertyChanged);
                WizardInfo.Dispose();
            }
            WizardInfo = new WizardInfo();
            WizardInfo.PropertyChanged += new PropertyChangedEventHandler(WizardInfo_PropertyChanged);

            if (m_wizardPages == null)
            {
                InitializeWizardPages();

                InitializeWizardHashTable();
            }

            LoadWizardPart(WizardPage.Welcome);
            m_wizardPageHashTable[WizardPage.Welcome].CanShow = true;

            CanStartAgain = false;

            Title = "Test Case Migrator Plus";
        }


        #endregion

        #region private methods

        /// <summary>
        /// Initializes the Wiard State  to Wizard Part hash Table
        /// </summary>
        private void InitializeWizardHashTable()
        {
            m_wizardPageHashTable = new Dictionary<WizardPage, IWizardPart>();
            m_wizardPageHashTable.Add(WizardPage.Welcome, new WelcomePart());
            m_wizardPageHashTable.Add(WizardPage.SelectDataSource, new SelectDataSourcePart());
            m_wizardPageHashTable.Add(WizardPage.SelectDestinationServer, new SelectDestinationServerPart());
            m_wizardPageHashTable.Add(WizardPage.SettingsFile, new SettingsFilePart());
            m_wizardPageHashTable.Add(WizardPage.FieldsSelection, new FieldsSelectionPart());
            m_wizardPageHashTable.Add(WizardPage.FieldMapping, new FieldMappingPart());
            m_wizardPageHashTable.Add(WizardPage.DataMapping, new DataMappingPart());
            m_wizardPageHashTable.Add(WizardPage.Linking, new LinkingPart());
            m_wizardPageHashTable.Add(WizardPage.MiscSettings, new MiscSettingsPart());
            m_wizardPageHashTable.Add(WizardPage.ConfirmSettings, new ConfirmSettingsPart());
            m_wizardPageHashTable.Add(WizardPage.Summary, new SummaryPart());

            WizardParts = new ObservableCollection<IWizardPart>();
            foreach (WizardPage page in m_wizardPages)
            {
                IWizardPart part = m_wizardPageHashTable[page];
                part.PropertyChanged += new PropertyChangedEventHandler(WizardPart_PropertyChanged);

                if (page != WizardPage.Summary)
                {
                    WizardParts.Add(part);
                }
            }
        }

        /// <summary>
        /// Initializes the List of Wizard States
        /// </summary>
        private void InitializeWizardPages()
        {
            m_wizardPages = new List<WizardPage>
            { 
                WizardPage.Welcome,
                WizardPage.SelectDataSource,
                WizardPage.SelectDestinationServer,
                WizardPage.SettingsFile,
                WizardPage.FieldMapping,
                WizardPage.DataMapping,
                WizardPage.Linking,
                WizardPage.MiscSettings,
                WizardPage.ConfirmSettings,
                WizardPage.Summary
            };
        }

        /// <summary>
        /// Listens to the Wizard part's notifications
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void WizardPart_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            IWizardPart wizardPart = sender as IWizardPart;
            if (WizardPart == null || wizardPart.WizardPage != WizardPart.WizardPage)
            {
                return;
            }

            // MigrateWorkitem Part's State notification
            if (String.CompareOrdinal(e.PropertyName, SummaryPart.MigrationStateTag) == 0)
            {
                SummaryPart summaryPart = sender as SummaryPart;
                if (summaryPart.MigrationState != WizardActionState.InProgress)
                {
                    IsMigrationCompleted = true;
                }
            }
            // Wizard Part's CanConfirm Notification
            else if (String.CompareOrdinal(e.PropertyName, BaseWizardPart.CanConfirmPropertyName) == 0 ||
                     String.CompareOrdinal(e.PropertyName, BaseWizardPart.CanNextPropertyName) == 0)
            {
                CanShowOtherWizardParts();
            }
        }

        /// <summary>
        /// Loads Wizard Part for corresponding Wizard State. Also checks whether to update current wizard part or not.
        /// </summary>
        /// <param name="wizardPage"></param>
        /// <param name="updateCurrentWizardPage"></param>
        /// <returns></returns>
        private void LoadWizardPart(WizardPage wizardPage)
        {
            m_loadWizardPage = wizardPage;
            m_worker = new BackgroundWorker();

            // Setting the event handlers of background worker
            m_worker.DoWork += new DoWorkEventHandler(LoadWizardPartInBackgroundThread);
            m_worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Worker_LoadWizardPartInBackgroundThreadCompleted);
            m_worker.RunWorkerAsync(null);
        }

        private void LoadWizardPartInBackgroundThread(object sender, DoWorkEventArgs e)
        {
            IsLoadingWizardPage = true;
            using (new AutoWaitCursor())
            {
                if (WizardPart == null)
                {
                    WizardPart = m_wizardPageHashTable[m_loadWizardPage];
                    WizardPart.Initialize(m_wizardInfo);
                }
                int newWizardPageIndex = m_wizardPages.IndexOf(m_loadWizardPage);
                int currentWizardPageIndex = CurrentWizardIndex;
                if (newWizardPageIndex > CurrentWizardIndex)
                {
                    IWizardPart part = m_wizardPageHashTable[m_wizardPages[currentWizardPageIndex]];
                    while (currentWizardPageIndex < newWizardPageIndex)
                    {
                        if (!part.UpdateWizardPart())
                        {
                            break;
                        }
                        currentWizardPageIndex++;

                        part.IsActiveWizardPart = false;
                        part = m_wizardPageHashTable[m_wizardPages[currentWizardPageIndex]];
                        part.IsActiveWizardPart = true;

                        if (currentWizardPageIndex == newWizardPageIndex && part.CanInitialize)
                        {
                            WizardPart = part;
                        }

                        part.Initialize(m_wizardInfo);
                        if (!part.CanShow)
                        {
                            string warning = part.Warning;
                            currentWizardPageIndex--;
                            part = m_wizardPageHashTable[m_wizardPages[currentWizardPageIndex]];
                            part.Warning = warning;
                            break;
                        }
                    }
                    WizardPart = part;
                    WizardPart.Warning = WizardPart.Warning;
                    CanShowOtherWizardParts();
                }
                else
                {
                    WizardPart = m_wizardPageHashTable[m_loadWizardPage];
                }
            }
            WizardPart.CanNext = WizardPart.ValidatePartState();
            m_wizardPageHashTable[WizardPage.ConfirmSettings].Initialize(m_wizardInfo);
            WizardInfo.CanConfirm = m_wizardPageHashTable[WizardPage.ConfirmSettings].CanInitialize;
            DisableNavigationIfLastPage();

            NotifyPropertyChanged("IsSummaryPage");
            NotifyPropertyChanged("IsConfirmPage");
            NotifyPropertyChanged("IsConfirmVisible");

            IsLoadingWizardPage = false;
        }

        private void Worker_LoadWizardPartInBackgroundThreadCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // If any exception is thrown during the work then sets the state as failed and set the corresponding message.
            if (e.Error != null)
            {
                throw e.Error;
            }
            m_worker.Dispose();
        }

        /// <summary>
        /// Disables Navigation to other Wizard parts if this is last Wizard Part
        /// </summary>
        private void DisableNavigationIfLastPage()
        {
            if (WizardPart.WizardPage == WizardPage.Summary)
            {
                for (int i = 0; i < WizardParts.Count; i++)
                {
                    WizardParts[i].CanShow = false;
                }
            }
        }

        /// <summary>
        /// Checks whether we can show other wizard Parts or not
        /// </summary>
        private void CanShowOtherWizardParts()
        {
            for (int i = CurrentWizardIndex + 1; i < m_wizardPages.Count; i++)
            {
                if (m_wizardPageHashTable[m_wizardPages[i]].CanInitialize)
                {
                    m_wizardPageHashTable[m_wizardPages[i]].CanShow = true;
                }
                else
                {
                    for (; i < m_wizardPages.Count; i++)
                    {
                        m_wizardPageHashTable[m_wizardPages[i]].CanShow = false;
                    }
                }
            }
            if (!WizardPart.CanNext)
            {
                for (int i = CurrentWizardIndex + 1; i < m_wizardPages.Count; i++)
                {
                    m_wizardPageHashTable[m_wizardPages[i]].CanShow = false;
                }
            }
            else if (CurrentWizardIndex < m_wizardPages.Count - 1)
            {

                m_wizardPageHashTable[m_wizardPages[CurrentWizardIndex + 1]].CanShow = true;
                // If Wizard part can confirm
                if (m_wizardInfo.CanConfirm)
                {
                    // then Set CanShow of all Visible Part as true
                    foreach (IWizardPart part in m_wizardPageHashTable.Values)
                    {
                        part.CanShow = true;
                    }
                }
            }
        }

        private void WizardInfo_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (String.CompareOrdinal(e.PropertyName, "DataSourceType") == 0)
            {
                if (m_wizardInfo.DataSourceType == DataSourceType.MHT)
                {
                    App.CallMethodInUISynchronizationContext(ArrangeWizardPartsForMHTFLow, null);
                }
                else if (m_wizardInfo.DataSourceType == DataSourceType.Excel)
                {
                    App.CallMethodInUISynchronizationContext(ArrangeWizardPartsForExcelFLow, null);
                }
            }
            else if (String.CompareOrdinal(e.PropertyName, "CanConfirm") == 0)
            {
                CanShowOtherWizardParts();
            }
            else if (String.CompareOrdinal(e.PropertyName, "WorkItemGenerator") == 0)
            {
                if (m_wizardInfo.WorkItemGenerator != null)
                {
                    m_wizardInfo.WorkItemGenerator.PropertyChanged += new PropertyChangedEventHandler(WorkItemGenerator_PropertyChanged);
                }
            }
        }

        private void WorkItemGenerator_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (String.CompareOrdinal(e.PropertyName, "SelectedWorkItemTypeName") == 0)
            {
                if (!string.IsNullOrEmpty(m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName))
                {
                    Title = "Test Case Migrator Plus - " + m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName;
                }
                else
                {
                    Title = "Test Case Migrator Plus";
                }
            }

        }

        private void ArrangeWizardPartsForExcelFLow(object obj)
        {
            if (m_wizardPages.Contains(WizardPage.FieldsSelection))
            {
                m_wizardPages.Remove(WizardPage.FieldsSelection);
                WizardParts.Remove(m_wizardPageHashTable[WizardPage.FieldsSelection]);
            }
            if (!m_wizardPages.Contains(WizardPage.Linking))
            {
                m_wizardPages.Insert(LinkingWizardPartPositionInWizardFlow, WizardPage.Linking);
                WizardParts.Insert(LinkingWizardPartPositionInWizardFlow, m_wizardPageHashTable[WizardPage.Linking]);
            }
        }

        private void ArrangeWizardPartsForMHTFLow(object page)
        {
            if (!m_wizardPages.Contains(WizardPage.FieldsSelection))
            {
                m_wizardPages.Insert(FieldsWizardPartPositionInWizardFlow, WizardPage.FieldsSelection);
                WizardParts.Insert(FieldsWizardPartPositionInWizardFlow, m_wizardPageHashTable[WizardPage.FieldsSelection]);
            }
            if (m_wizardPages.Contains(WizardPage.Linking))
            {
                m_wizardPages.Remove(WizardPage.Linking);
                WizardParts.Remove(m_wizardPageHashTable[WizardPage.Linking]);
            }
        }

        #endregion

        #region IDisposible Implementation

        /// <summary>
        /// Disposes the Data Source
        /// </summary>
        public void Dispose()
        {
            m_wizardInfo.PropertyChanged -= WizardInfo_PropertyChanged;
            m_wizardInfo.Dispose();
            ExcelParser.Quit();
            MHTParser.Quit();
        }

        #endregion
    }
}