//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.ComponentModel;
    using System.Windows;

    /// <summary>
    /// Delegate Definition for Message Callback on any Message Response.
    /// </summary>
    /// <param name="args"></param>
    internal delegate void MessageCallbackDelegate(MessageEventArgs args);


    /// <summary>
    /// Message Event Args: Used for displaying messages occured during the course of Wizard Configuration and Wizard Actions.
    /// </summary>
    internal class MessageEventArgs : RoutedEventArgs, INotifyPropertyChanged
    {
        #region Fields

        // Member variables used for data binding
        private string m_title;
        private string m_likelyCause;
        private string m_potentialSolution;
        private MessageCategory m_category;

        #endregion

        #region Constructor

        /// <summary>
        /// Basic Constructor
        /// </summary>
        public MessageEventArgs()
        {
            CancelButtonLabel = Resources.CloseButtonLabel;
        }

        /// <summary>
        /// Constructor for complete intialization
        /// </summary>
        /// <param name="category">Message Category</param>
        /// <param name="title">Title of the Message</param>
        /// <param name="likelycause">Likely Cause of the message</param>
        /// <param name="potentialSolution">Potential Solution of the message</param>
        /// <param name="firstButtonLabel">The Label of the first button</param>
        /// <param name="secondButtonLabel">The Label of the Second button</param>
        /// <param name="cancelButtonLabel">The label of the Cancel Button</param>
        /// <param name="callback">The Method to be called after message response</param>
        public MessageEventArgs(MessageCategory category,
                                string title,
                                string likelycause,
                                string potentialSolution,
                                string firstButtonLabel,
                                string secondButtonLabel,
                                string cancelButtonLabel,
                                MessageCallbackDelegate callback)
            : this()
        {
            Category = category;
            Title = title;
            LikelyCause = likelycause;
            PotentialSolution = potentialSolution;
            FirstButtonLabel = firstButtonLabel;
            SecondButtonLabel = secondButtonLabel;

            // Cancel Button's Label can't be empty
            if (!string.IsNullOrEmpty(cancelButtonLabel))
            {
                CancelButtonLabel = cancelButtonLabel;
            }
            Callback = callback;
        }

        #endregion

        #region Properties

        /// <summary>
        /// The Message Type: warning, error or information
        /// </summary>
        public MessageCategory Category
        {
            get
            {
                return m_category;
            }
            set
            {
                m_category = value;
                NotifyPropertyChanged("Category");
            }
        }

        /// <summary>
        /// Title of the message
        /// </summary>
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
        /// The likely cause of the Message
        /// </summary>
        public string LikelyCause
        {
            get
            {
                return m_likelyCause;
            }
            set
            {
                m_likelyCause = value;
                NotifyPropertyChanged("LikelyCause");
            }
        }

        /// <summary>
        /// The Potential Solution os the Message
        /// </summary>
        public string PotentialSolution
        {
            get
            {
                return m_potentialSolution;
            }
            set
            {
                m_potentialSolution = value;
                NotifyPropertyChanged("PotentialSolution");
            }
        }

        /// <summary>
        /// The Label of the First button
        /// </summary>
        public string FirstButtonLabel
        {
            get;
            set;
        }

        /// <summary>
        /// The Label of the Second button
        /// </summary>
        public string SecondButtonLabel
        {
            get;
            set;
        }

        /// <summary>
        /// The label of the Cancel Button
        /// </summary>
        public string CancelButtonLabel
        {
            get;
            set;
        }

        /// <summary>
        /// Message Response Definition (First, Second or Cancel)
        /// </summary>
        public MessageResponseDefinition ResponseDefinition
        {
            get;
            set;
        }

        /// <summary>
        /// The Delegate to be called when message gets some response from the user
        /// </summary>
        public MessageCallbackDelegate Callback
        {
            get;
            set;
        }

        /// <summary>
        /// Message Data Information. Can be used for interaction between callback and message calling method
        /// </summary>
        public object Data
        {
            get;
            set;
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



        public void SetValues(MessageEventArgs messageEventArgs)
        {
            if (messageEventArgs != null)
            {
                Title = messageEventArgs.Title;
                LikelyCause = messageEventArgs.LikelyCause;
                PotentialSolution = messageEventArgs.PotentialSolution;
            }
        }

        internal void Clear()
        {
            Title = string.Empty;
            LikelyCause = string.Empty;
            PotentialSolution = string.Empty;
        }
    }
}
