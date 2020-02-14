//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Collections.Generic;
    using System.ComponentModel;
    
    /// <summary>
    /// Storage Information of Data Source Parser
    /// </summary>
    public interface IDataStorageInfo : INotifyPropertyChanged
    {
        /// <summary>
        /// Source Information where Parser will parse
        /// </summary>
        string Source { get; }


        string SourceIdFieldName { get; set; }

        string TestSuiteFieldName { get; set; }

        IList<ILinkRule> LinkRules { get; }

        /// <summary>
        /// Start Delimeter of word which is to be taken as parameter
        /// </summary>
        string StartParameterizationDelimeter { get; set; }


        /// <summary>
        /// End delimeter of word which is to be taken as parameter
        /// </summary>
        string EndParameterizationDelimeter { get; set; }

        
        /// <summary>
        /// Is to take multiple lines present in single block of step to take as multiple lines.
        /// </summary>
        bool IsMultilineSense { get; set; }


        /// <summary>
        /// Field names present in this Data Source
        /// </summary>
        IList<SourceField> FieldNames { get; set; }
    }
}
