//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    /// <summary>
    /// Represents Link between two WorkItems
    /// Two ends of link are named as 'Start WorkItem' and 'End WorkItem'
    /// 
    /// Start WorkItem------------>EndWorkItem
    /// </summary>
    internal class Link : ILink
    {
        /// <summary>
        /// TFS Type Name of Start Workitem
        /// </summary>
        public string StartWorkItemTfsTypeName
        {
            get;
            set;
        }

        /// <summary>
        /// The ID of Start WorkItem in Data Source(Excel/MHT)
        /// </summary>
        public string StartWorkItemSourceId
        {
            get;
            set;
        }

        /// <summary>
        /// The TFS ID of Start WorkItem
        /// </summary>
        public int StartWorkItemTfsId
        {
            get;
            set;
        }

        /// <summary>
        /// The Type Name of Link
        /// </summary>
        public string LinkTypeName
        {
            get;
            set;
        }

        /// <summary>
        /// TFS Type Name of End Work Item
        /// </summary>
        public string EndWorkItemTfsTypeName
        {
            get;
            set;
        }

        /// <summary>
        /// The ID of End WorkItem in Data Source(Excel/MHT)
        /// </summary>
        public string EndWorkItemSourceId
        {
            get;
            set;
        }

        /// <summary>
        /// The TFS ID of End WorkItem
        /// </summary>
        public int EndWorkItemTfsId
        {
            get;
            set;
        }

        /// <summary>
        /// Is this link exists in TFS
        /// </summary>
        public bool IsExistInTfs
        {
            get;
            set;
        }

        /// <summary>
        /// The error occured when tried to create the link from stat workitem to tfs workitem
        /// </summary>
        public string Message
        {
            get;
            set;
        }


        public int SessionId
        {
            get;
            set;
        }


        public Status LinksStatus
        {
            get
            {
                if (IsExistInTfs)
                {
                    return Status.Passed;
                }
                else if (string.IsNullOrEmpty(Message))
                {
                    return Status.Warning;
                }
                else
                {
                    return Status.Failed;
                }
            }
        }

        public string EndWorkItemCategory
        {
            get;
            set;
        }

        public string StartWorkItemCategory
        {
            get;
            set;
        }
    }
}
