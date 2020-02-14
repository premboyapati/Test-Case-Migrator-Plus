//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.ComponentModel;

    /// <summary>
    /// Base Class for all Wizard Actions
    /// </summary>
    internal abstract class WizardAction : NotifyPropertyChange, IDisposable
    {
        #region Fields

        // Member Variables needed for binding
        private WizardActionState m_state;
        private string m_action;
        private string m_message;

        // Background worker to perform action on background to make the interacting thread to be more responsive
        protected BackgroundWorker m_worker;

        // Wizard info Object needed for Initialization of Wizard Action. It may gets updated after the action
        protected WizardInfo m_wizardInfo;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes the m_wizardInfo and BackgroundWorker. 
        /// Also sets the Current State of WizardAction as pending
        /// </summary>
        /// <param name="wizardInfo"></param>
        public WizardAction(WizardInfo wizardInfo)
        {
            m_wizardInfo = wizardInfo;
            m_state = WizardActionState.Pending;
            InitializeBackgroundWorker();
        }

        #endregion

        #region BackGround Worker

        /// <summary>
        /// Initializes the Backgroundworker
        /// </summary>
        private void InitializeBackgroundWorker()
        {
            m_worker = new BackgroundWorker();

            // Setting the event handlers of background worker
            m_worker.DoWork += new DoWorkEventHandler(Worker_DoWork);
            m_worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Worker_RunWorkerCompleted);
            m_worker.ProgressChanged += new ProgressChangedEventHandler(Worker_ProgressChanged);
        }


        /// <summary>
        /// Will be called when Background worker is asked to perform work
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        [STAThread]
        void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            // Setting the state to be in progress
            State = WizardActionState.InProgress;
            BackgroundWorker worker = sender as BackgroundWorker;

            // Perrforming the work and setting the subsequent result
            e.Result = DoWork(e);
        }


        /// <summary>
        /// Will be called when Background worker will perform its work.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // If any exception is thrown during the work then sets the state as failed and set the corresponding message.
            if (e.Error != null)
            {
                State = WizardActionState.Failed;
                WorkItemMigratorException te = e.Error as WorkItemMigratorException;
                if (te != null)
                {
                    Message = te.Args.Title;
                }
                else
                {
                    Message = e.Error.Message;
                }
            }
            else
            {
                // else just call OnWorkComplete method
                OnWorkComplete(e);
            }
        }

        /// <summary>
        /// This is called when there is some progress during work.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ShowProgress(e);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Action Name:- used to display the name of Action is Actions Pane
        /// </summary>
        public WizardActionName ActionName
        {
            get;
            protected set;
        }

        /// <summary>
        /// The current State of Wizard Action
        /// </summary>
        public WizardActionState State
        {
            get
            {
                return m_state;
            }
            set
            {
                m_state = value;
                NotifyPropertyChanged("State");
                NotifyPropertyChanged("Status");
            }
        }

        /// <summary>
        /// The Wizard Action's Description
        /// </summary>
        public string Description
        {
            get
            {
                return m_action;
            }
            set
            {
                m_action = value;
                NotifyPropertyChanged("Action");
            }
        }

        /// <summary>
        /// The current state of Wizard Action in string format telling its status
        /// </summary>
        public string Status
        {
            get
            {
                switch (State)
                {
                    case WizardActionState.Pending:
                        return Resources.Pending;

                    case WizardActionState.InProgress:
                        return Resources.InProgress;

                    case WizardActionState.Stopped:
                        return Resources.Stopped;

                    case WizardActionState.Success:
                        return Resources.Success;

                    case WizardActionState.Warning:
                        return Resources.Warning;

                    case WizardActionState.Failed:
                        return Resources.Failed;

                    default:
                        return string.Empty;
                }
            }
        }

        /// <summary>
        /// Used to show the current progress of Wizard Action like to show the progress or 
        /// to show the error message when some error occures.
        /// </summary>
        public string Message
        {
            get
            {
                return m_message;
            }
            set
            {
                m_message = value;
                NotifyPropertyChanged("Message");
            }
        }

        #endregion

        #region Abstract / Virtual Methods

        /// <summary>
        /// Asks the background worker to start the worker
        /// </summary>
        public virtual void Start()
        {
            m_worker.RunWorkerAsync(null);
        }

        /// <summary>
        /// Asks Background worker to cancel the Work and sets the state as Stopped
        /// </summary>
        public virtual void Stop()
        {
            Message = Resources.UserInterruptionText;
            if (m_worker.WorkerSupportsCancellation)
            {
                m_worker.CancelAsync();
            }
        }

        /// <summary>
        /// Abstract function which actually do the work and returns the result in form of state
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        protected abstract WizardActionState DoWork(DoWorkEventArgs e);


        /// <summary>
        /// Virtual function to show the progress
        /// </summary>
        /// <param name="e"></param>
        protected virtual void ShowProgress(ProgressChangedEventArgs e)
        {
            return;
        }

        /// <summary>
        /// Virtual function to set the state/message after work is completed
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnWorkComplete(RunWorkerCompletedEventArgs e)
        {
            State = (WizardActionState)e.Result;
        }

        #endregion

        #region IDisposable Implementation

        /// <summary>
        /// Disposes the Background Worker
        /// </summary>
        public void Dispose()
        {
            if (m_worker != null)
            {
                m_worker.Dispose();
            }
        }

        #endregion
    }
}
