//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Windows.Data;

    /// <summary>
    /// Converts a boolean value into its complement
    /// </summary>
    [ValueConversion(typeof(bool), typeof(bool))]
    class InverseBooleanConverter : IValueConverter
    {
        #region IValueConverter_Implementation

        /// <summary>
        /// Converts a bool value to its inverse.
        /// </summary>
        /// <param name="value">value to be converted</param>
        /// <param name="targetType">Specifed target type</param>
        /// <param name="parameter">not used</param>
        /// <param name="culture">not used</param>
        /// <returns>Inverse of value</returns>
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return this.GetInverse(value);
        }

        /// <summary>
        /// Converts a bool value to its inverse.
        /// </summary>
        /// <param name="value">value to be converted</param>
        /// <param name="targetType">Specifed target type</param>
        /// <param name="parameter">not used</param>
        /// <param name="culture">not used</param>
        /// <returns>Inverse of value</returns>
        /// <returns></returns>
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return this.GetInverse(value);
        }

        #region Private methods
        /// <summary>
        /// Determines the inverse value of a boolean
        /// </summary>
        /// <param name="targetBool">value to evalaute</param>
        /// <returns>The inverse of the boolean value</returns>
        private bool GetInverse(object targetBool)
        {
            bool converted = false;

            if (targetBool == null)
            {
                converted = true;
            }
            else
            {
                bool boolValue = (bool)targetBool;
                converted = !boolValue;
            }

            return converted;
        }
        #endregion
        #endregion
    }
}
