//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Responsible for Parsing the data source and all related operations
    /// </summary>
    public interface IDataSourceParser : IDisposable
    {

        /// <summary>
        /// Data Storage Information of Parser
        /// </summary>
        IDataStorageInfo StorageInfo { get; }

        
        /// <summary>
        /// Parse the Data source for field names and updates the 'StorageInfo' with the field names
        /// </summary>
        void ParseDataSourceFieldNames();


        /// <summary>
        /// Mapping from field name to correspording 'IWorkItemField'
        /// </summary>
        IDictionary<string, IWorkItemField> FieldNameToFields { get;  set; }


        /// <summary>
        /// Returns the next source workitem to parse in the data source
        /// </summary>
        /// <returns></returns>
        ISourceWorkItem GetNextWorkItem();


        /// <summary>
        /// List of Source Workitems parsed till now with no parsed settings(parameterization & multiline) applied
        /// </summary>
        IList<ISourceWorkItem> RawSourceWorkItems { get; }


        /// <summary>
        /// List of Source Workitems parsed till now with parsed settings(parameterization & multiline) applied
        /// </summary>
        IList<ISourceWorkItem> ParsedSourceWorkItems { get; }
        
    }
}