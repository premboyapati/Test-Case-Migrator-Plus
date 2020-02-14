//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.IO;

    internal class MHTStorageInfo : DataStorageInfoBase
    {
        #region Constants

        private const string MHTExtension = ".mht";

        #endregion

        #region Properties

        /// <summary>
        /// List of allowed field names for this MHT File
        /// </summary>
        public IList<string> PossibleFieldNames
        {
            get;
            set;
        }

        /// <summary>
        /// Is First Line of MHT file Title?
        /// </summary>
        public bool IsFirstLineTitle
        {
            get;
            set;
        }

        /// <summary>
        /// whether to take file name as title?
        /// </summary>
        public bool IsFileNameTitle
        {
            get;
            set;
        }

        /// <summary>
        /// this is field name mapped to TFS title field
        /// </summary>
        public string TitleField
        {
            get;
            set;
        }

        /// <summary>
        /// Field Name of Steps Field
        /// </summary>
        public string StepsField
        {
            get;
            set;
        }

        #endregion

        #region Constructor

        public MHTStorageInfo(string mhtSource)
        {
            if (mhtSource == null)
            {
                throw new ArgumentNullException("mhtSource", "mhtSource is null");
            }

            if (!File.Exists(mhtSource) ||
                (String.CompareOrdinal(Path.GetExtension(mhtSource), MHTExtension) != 0 &&
                 String.CompareOrdinal(Path.GetExtension(mhtSource), ".doc") != 0 &&
                 String.CompareOrdinal(Path.GetExtension(mhtSource), ".docx") != 0))
            {
                throw new ArgumentException("mhtSource is not valid", "mhtSource");
            }

            Source = mhtSource;
        }

        #endregion
    }
}