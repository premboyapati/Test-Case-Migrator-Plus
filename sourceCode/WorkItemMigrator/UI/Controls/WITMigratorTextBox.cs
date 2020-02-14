//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;

    internal class WITMigratorTextBox : TextBox
    {
        #region Constructor

        static WITMigratorTextBox()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(WITMigratorTextBox), new FrameworkPropertyMetadata(typeof(WITMigratorTextBox)));
        }

        public WITMigratorTextBox()
            : base()
        {
            this.TextChanged += WITMigratorTextBox_TextChanged;
            this.LostKeyboardFocus += new KeyboardFocusChangedEventHandler(WITMigratorTextBox_LostKeyboardFocus);
        }

        #endregion

        #region Dependency Properties

        public bool IsRequired
        {
            get { return (bool)GetValue(IsRequiredProperty); }
            set { SetValue(IsRequiredProperty, value); }
        }

        public static readonly DependencyProperty IsRequiredProperty =
            DependencyProperty.Register("IsRequired", typeof(bool), typeof(WITMigratorTextBox), new UIPropertyMetadata(false));

        public MessageEventArgs ErrorArgs
        {
            get { return (MessageEventArgs)GetValue(ErrorArgsProperty); }
            set { SetValue(ErrorArgsProperty, value); }
        }

        public static readonly DependencyProperty ErrorArgsProperty =
            DependencyProperty.Register("ErrorArgs", typeof(MessageEventArgs), typeof(WITMigratorTextBox), new UIPropertyMetadata(null));


        public event EventHandler<MessageEventArgs> TextChangeAction = delegate { };

        #endregion

        #region private methods

        private void WITMigratorTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!IsFocused)
            {
                SetErrorArgs(sender);
            }
        }

        private void WITMigratorTextBox_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            SetErrorArgs(sender);
        }

        private void SetErrorArgs(object sender)
        {
            ToolTip = null;
            ErrorArgs = new MessageEventArgs();
            TextChangeAction(sender, ErrorArgs);
            if (string.IsNullOrEmpty(ErrorArgs.Title))
            {
                ErrorArgs = null;
            }
            else
            {
                ToolTip = ErrorArgs.Title + "\n" + ErrorArgs.LikelyCause + "\n" + ErrorArgs.PotentialSolution;
            }
        }
        #endregion

    }
}
