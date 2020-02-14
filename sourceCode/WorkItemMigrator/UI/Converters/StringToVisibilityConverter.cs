//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Windows;
    using System.Windows.Data;

    [ValueConversion(typeof(string), typeof(Visibility))]
    public class StringToVisibilityConverter : IValueConverter
    {
        #region IValueConverter Members

        /// <summary>
        /// This is used to collaspe text controls when they have no content
        /// </summary>
        /// <param name="value">String to evaluate</param>
        /// <param name="targetType">System.Windows>Visibility</param>
        /// <param name="parameter">not used</param>
        /// <param name="culture">not used</param>
        /// <returns>If the string value has a lenght, then it returns Visibility.Visible.  
        /// Otherwise it returns Visibility.Collapsed</returns>
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            // Default behavior is colapsed
            Visibility visibility = Visibility.Collapsed;

            if (value != null)
            {
                string information = value as String;
                if (information != null)
                {
                    if (information.Length > 0)
                    {
                        // if the string has length, then show it
                        visibility = Visibility.Visible;
                    }
                }
            }

            return visibility;
        }

        /// <summary>
        /// Not used
        /// </summary>
        /// <param name="value"></param>
        /// <param name="targetType"></param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
