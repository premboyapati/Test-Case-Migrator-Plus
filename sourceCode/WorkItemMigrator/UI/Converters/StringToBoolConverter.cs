//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Windows;
    using System.Windows.Data;

    [ValueConversion(typeof(string), typeof(bool))]
    public class StringToBoolConverter : IValueConverter
    {
        #region IValueConverter Members

        /// <summary>
        /// Returns true if string is not empty
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            bool isStringNotEmpty = false;

            if (value != null)
            {
                string information = value as String;
                if (information != null)
                {
                    if (information.Length > 0)
                    {
                        isStringNotEmpty = true;
                    }
                }
            }
            return isStringNotEmpty;
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

