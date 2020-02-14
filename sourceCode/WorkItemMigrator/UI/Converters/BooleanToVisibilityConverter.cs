// Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Windows;
    using System.Windows.Data;

    [ValueConversion(typeof(bool), typeof(Visibility))]
    public class BooleanToVisibilityConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (targetType != typeof(Visibility))
            {
                throw new NotSupportedException();
            }

            Type valueType = (value == null ? null : value.GetType());

            if (valueType != null)
            {
                if (valueType != typeof(bool?)
                 && valueType != typeof(bool))
                {
                    throw new NotSupportedException();
                }
            }

            bool flag = false;

            if (valueType == typeof(bool))
            {
                flag = (bool)value;
            }
            else if (valueType == typeof(bool?))
            {
                bool? nullable = (bool?)value;
                flag = nullable.HasValue ? nullable.Value : false;
            }

            return (flag ? Visibility.Visible : Visibility.Collapsed);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
