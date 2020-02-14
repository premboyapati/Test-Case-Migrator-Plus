//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Windows;

    /// <summary>
    /// Interaction logic for MessageBox.xaml
    /// </summary>
    internal partial class MessageBox : Window
    {
        #region Fields
        private MessageEventArgs m_messageArgs;
        #endregion

        #region Constructor

        public MessageBox(MessageEventArgs args)
        {
            InitializeComponent();
            m_messageArgs = args;
            DataContext = args;
        }

        #endregion

        #region private methods

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (sender == FirstButton)
            {
                m_messageArgs.ResponseDefinition = MessageResponseDefinition.FirstResponse;
            }
            else if (sender == SecondButton)
            {
                m_messageArgs.ResponseDefinition = MessageResponseDefinition.SecondResponse;
            }
            else
            {
                m_messageArgs.ResponseDefinition = MessageResponseDefinition.CancelResponse;
            }

            if (m_messageArgs.Callback != null)
            {
                m_messageArgs.Callback(m_messageArgs);
            }
            Close();
        }

        #endregion
    }
}
