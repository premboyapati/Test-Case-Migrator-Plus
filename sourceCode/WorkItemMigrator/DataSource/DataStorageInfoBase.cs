using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    internal class DataStorageInfoBase : NotifyPropertyChange, IDataStorageInfo
    {
        #region Fields

        private string m_source;
        private string m_startParameterizationDelimeter;
        private string m_endParameterizationDelimeter;
        private bool m_isMultiLineSense;

        #endregion

        #region Properties

        /// <summary>
        /// MHT File Path
        /// </summary>
        public string Source
        {
            get
            {
                return m_source;
            }
            protected set
            {
                m_source = value;
                NotifyPropertyChanged("Source");
            }
        }

        public string StartParameterizationDelimeter
        {
            get
            {
                return m_startParameterizationDelimeter;
            }
            set
            {
                m_startParameterizationDelimeter = value;
                NotifyPropertyChanged("StartParameterizationDelimeter");
            }
        }

        public string EndParameterizationDelimeter
        {
            get
            {
                return m_endParameterizationDelimeter;
            }
            set
            {
                m_endParameterizationDelimeter = value;
                NotifyPropertyChanged("EndParameterizationDelimeter");
            }
        }

        public bool IsMultilineSense
        {
            get
            {
                return m_isMultiLineSense;
            }
            set
            {
                m_isMultiLineSense = value;
                NotifyPropertyChanged("IsMultiLineSense");
            }
        }

        public IList<SourceField> FieldNames
        {
            get;
            set;
        }

        public string SourceIdFieldName
        {
            get;
            set;
        }

        public string TestSuiteFieldName
        {
            get;
            set;
        }

        public IList<ILinkRule> LinkRules
        {
            get;
            protected set;
        }

        #endregion

        #region Constructor

        protected DataStorageInfoBase()
        {
            FieldNames = new List<SourceField>();
            LinkRules = new List<ILinkRule>();
        }

        #endregion

    }
}
