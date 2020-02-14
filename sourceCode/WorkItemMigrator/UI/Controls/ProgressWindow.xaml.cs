//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.ComponentModel;
    using System.Windows;

    /// <summary>
    /// back end part of Progress View
    /// </summary>
    internal partial class ProgressWindow : Window, INotifyPropertyChanged
    {
        #region Fields

        private string m_text;
        private string m_header;
        private double m_progressValue;
        private bool m_isClosed;

        #endregion

        #region Constants

        public const string IsClosedPropertyName = "IsClosed";

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor to intialize Title, Text and Progress Status 
        /// </summary>
        /// <param name="title"></param>
        /// <param name="text"></param>
        public ProgressWindow()
        {
            // Removing Wait Cursor
            using (new AutoWaitCursor())
            { }

            InitializeComponent();
            DataContext = this;
            ProgressValue = 0;
            IsClosed = false;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Title of the Progress Overlay Part
        /// </summary>
        public string Header
        {
            get
            {
                return m_header;
            }
            set
            {
                m_header = value;
                NotifyPropertyChanged("Header");
            }
        }

        public string Text
        {
            get
            {
                return m_text;
            }
            set
            {
                m_text = value;
                NotifyPropertyChanged("Text");
            }
        }

        /// <summary>
        /// This is the percentage of work completed.
        /// </summary>
        public double ProgressValue
        {
            get
            {
                return m_progressValue;
            }
            set
            {
                m_progressValue = value;
                NotifyPropertyChanged("ProgressValue");
            }
        }

        public bool IsClosed
        {
            get
            {
                return m_isClosed;
            }
            private set
            {
                m_isClosed = value;
                NotifyPropertyChanged(IsClosedPropertyName);
            }
        }

        #endregion

        #region Event Handlers methods

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsClosed = true;
            Close();
        }

        #endregion

        #region INotifyPropertyChanged Implementation

        /// <summary>
        /// Property Change Event
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Send the notification for Property
        /// </summary>
        /// <param name="propertyName">Property whose change is to be notified</param>
        protected void NotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #endregion

    }
}

