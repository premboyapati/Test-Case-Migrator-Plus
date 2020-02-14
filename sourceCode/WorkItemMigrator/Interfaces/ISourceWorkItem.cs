//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Collections.Generic;

    /// <summary>
    /// Represents Workitem at Source 
    /// </summary>
    public interface ISourceWorkItem
    {
        /// <summary>
        /// The Path of Source containing this workitem
        /// </summary>
        string SourcePath { get; set; }


        /// <summary>
        /// ID of Workitem at Source
        /// </summary>
        string SourceId { get; set; }


        /// <summary>
        /// List of suite paths containg this particular workitem(test case)
        /// </summary>
        IList<string> TestSuites { get; }

        /// <summary>
        /// List of Link that this workitem will have with other work items 
        /// </summary>
        IList<ILink> Links { get; }


        /// <summary>
        /// A HashTable of mapping between Field to Value
        /// </summary>
        Dictionary<string, object> FieldValuePairs { get; }

    }
}
