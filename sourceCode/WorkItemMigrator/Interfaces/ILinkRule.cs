//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;

    /// <summary>
    /// Represents linking rule between two work item types
    /// Linking of Start WorkItem Type To EndWorkItem Type is shown Below
    /// 
    ///                               LinkType
    /// Start Work Item Type -------------------------> End WorkItem Type
    ///                                   [Source Field Name of End WorkItem Type]
    ///                                   
    /// </summary>
    public interface ILinkRule
    {
        /// <summary>
        /// TFS Type Name of Start Work Item Type
        /// </summary>
        string StartWorkItemCategory { get; }


        /// <summary>
        /// TFS Link Type Name which is going to be established between the workitems of two workItem Type
        /// </summary>
        string LinkTypeReferenceName { get;}


        /// <summary>
        /// FieldName present in Data Source(Excel) which will have the workitem ids of End Work Item Type
        /// </summary>
        string SourceFieldNameOfEndWorkItemCategory { get; }


        /// <summary>
        /// TFS WorkItem Type Name of End Side of the Link
        /// </summary>
        string EndWorkItemCategory { get; }

        string Description { get; }

    }
}
