//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.IO;

    internal class ExcelStorageInfo : DataStorageInfoBase
    {
        #region Fields

        private string m_worksheetName;
        private string m_headerRow;

        #endregion

        #region Constructor

        public ExcelStorageInfo(string source)
        {
            if (string.IsNullOrEmpty(source))
            {
                throw new ArgumentNullException("source", "source is null");
            }
            if (!File.Exists(source) ||
                (String.CompareOrdinal(Path.GetExtension(source), ".xls") != 0 &&
                 String.CompareOrdinal(Path.GetExtension(source), ".xlsx") != 0))
            {
                throw new ArgumentException("source is not valid", "source");
            }
            Source = source;
            RowContainingFieldNames = "1";
        }

        #endregion

        #region Properties

        /// <summary>
        /// Percentage of parsing done for excel source
        /// </summary>
        public double ProgressPercentage
        {
            get;
            set;
        }

        /// <summary>
        /// Worksheet to parse
        /// </summary>
        public string WorkSheetName
        {
            get
            {
                return m_worksheetName;
            }
            set
            {
                m_worksheetName = value;
                NotifyPropertyChanged("WorkSheetName");
            }
        }

        /// <summary>
        /// Excel Row containing Field names
        /// </summary>
        public string RowContainingFieldNames
        {
            get
            {
                return m_headerRow;
            }
            set
            {
                m_headerRow = value;
                NotifyPropertyChanged("HeaderRow");
            }

        }
        #endregion
    }
}
