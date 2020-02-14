//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;

    /// <summary>
    /// Represents Link between two WorkItems
    /// Two ends of link are named as 'Start WorkItem' and 'End WorkItem'
    /// 
    /// Start WorkItem------------>EndWorkItem
    /// </summary>
    public interface ILink
    {
        /// <summary>
        /// The ID of End WorkItem in Data Source(Excel/MHT)
        /// </summary>
        string EndWorkItemSourceId { get; set; }


        /// <summary>
        /// The TFS ID of End WorkItem
        /// </summary>
        int EndWorkItemTfsId { get; set; }


        /// <summary>
        /// TFS Type Name of End Work Item
        /// </summary>
        string EndWorkItemTfsTypeName { get; set; }

        string EndWorkItemCategory { get; set; }

        Status LinksStatus { get; }

        /// <summary>
        /// Is this link exists in TFS
        /// </summary>
        bool IsExistInTfs { get; set; }


        int SessionId { get; set; }


        /// <summary>
        /// The Type Name of Link
        /// </summary>
        string LinkTypeName { get; set; }


        /// <summary>
        /// The error occured when tried to create the link from stat workitem to tfs workitem
        /// </summary>
        string Message { get; set; }


        /// <summary>
        /// The ID of Start WorkItem in Data Source(Excel/MHT)
        /// </summary>
        string StartWorkItemSourceId { get; set; }


        /// <summary>
        /// The TFS ID of Start WorkItem
        /// </summary>
        int StartWorkItemTfsId { get; set; }

        /// <summary>
        /// TFS Type Name of Start Workitem
        /// </summary>
        string StartWorkItemTfsTypeName { get; set; }

        string StartWorkItemCategory { get; set; }

    }
}
