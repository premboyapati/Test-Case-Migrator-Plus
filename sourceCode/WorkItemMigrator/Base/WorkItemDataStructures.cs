//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Collections.Generic;

    /// <summary>
    /// Encapsulates Field name with the attribute whether this field is deletable or not?
    /// </summary>
    public class SourceField : NotifyPropertyChange
    {
        /// <summary>
        /// name of the field prsent in the source
        /// </summary>
        public string FieldName
        {
            get;
            private set;
        }

        /// <summary>
        /// Can Remove this field from the list of all fields
        /// </summary>
        public bool CanDelete
        {
            get;
            private set;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fieldName"></param>
        /// <param name="canDelete"></param>
        public SourceField(string fieldName, bool canDelete)
        {
            FieldName = fieldName;
            CanDelete = canDelete;
        }
    }


    /// <summary>
    /// Internal Data Structre that represents TFS Test Step Field
    /// </summary>
    internal struct SourceTestStep
    {
        // title of test step
        public string title;

        // expected result of test step
        public string expectedResult;

        // list of attachments prsent in this step
        public List<string> attachments;

        /// <summary>
        /// Contructor
        /// </summary>
        /// <param name="title"></param>
        /// <param name="expectedResult"></param>
        /// <param name="attachments"></param>
        public SourceTestStep(string title, string expectedResult, List<string> attachments)
        {
            this.title = title;
            this.expectedResult = expectedResult;
            this.attachments = attachments;
        }
    }


    /// <summary>
    /// Internal Data Structure that represents a single entiity to be imported into TFS server as a workitem
    /// </summary>
    internal class SourceWorkItem : ISourceWorkItem
    {
        /// <summary>
        /// The Path of Source containing this workitem
        /// </summary>
        public string SourcePath
        {
            get;
            set;
        }

        /// <summary>
        /// ID of Workitem at Source
        /// </summary>
        public string SourceId
        {
            get;
            set;
        }

        /// <summary>
        /// List of suite paths containg this particular workitem(test case)
        /// </summary>
        public IList<string> TestSuites
        {
            get;
            set;
        }

        /// <summary>
        /// List of Link that this workitem will have with other work items 
        /// </summary>
        public IList<ILink> Links
        {
            get;
            set;
        }

        /// <summary>
        /// A HashTable of mapping between Field to Value
        /// </summary>
        public Dictionary<string, object> FieldValuePairs
        {
            get;
            protected set;
        }

        /// <summary>
        /// Basic Constructor to initialize member variables
        /// </summary>
        public SourceWorkItem()
        {
            FieldValuePairs = new Dictionary<string, object>();
            TestSuites = new List<string>();
            Links = new List<ILink>();
        }

        /// <summary>
        /// Creates a clone source work item
        /// </summary>
        /// <param name="dsWorkItem"></param>
        public SourceWorkItem(ISourceWorkItem dsWorkItem)
        {
            SourcePath = dsWorkItem.SourcePath;
            SourceId = dsWorkItem.SourceId;
            TestSuites = dsWorkItem.TestSuites;
            Links = dsWorkItem.Links;
            FieldValuePairs = new Dictionary<string, object>();
            foreach (var kvp in dsWorkItem.FieldValuePairs)
            {
                FieldValuePairs.Add(kvp.Key, kvp.Value);
            }
        }
    }

    /// <summary>
    /// Internal Data Structure for TFS Workitem with an error
    /// It is a failed Data Source WorkItem which is failed to migrate with an error.
    /// </summary>
    internal class FailedSourceWorkItem : SourceWorkItem
    {
        /// <summary>
        /// The error which is caused while migration
        /// </summary>
        public string Error
        {
            get;
            set;
        }

        /// <summary>
        /// Constructor to initialzie workitem from other workitem
        /// </summary>
        /// <param name="dsWorkItem"></param>
        /// <param name="error"></param>
        public FailedSourceWorkItem(ISourceWorkItem dsWorkItem, string error)
            : base(dsWorkItem)
        {
            this.Error = error;
        }
    }

    /// <summary>
    /// Internal Data Structure for TFS Workitem with a warning
    /// It is a passed Data Source WorkItem which has different value from the Data Source in some of its field
    /// </summary>
    internal class WarningSourceWorkItem : PassedSourceWorkItem
    {
        /// <summary>
        /// The error which is caused while migration
        /// </summary>
        public string Warning
        {
            get;
            set;
        }

        /// <summary>
        /// Constructor to initialzie workitem from other workitem
        /// </summary>
        public WarningSourceWorkItem(ISourceWorkItem dsWorkItem, int id, string warning)
            : base(dsWorkItem, id)
        {
            this.Warning = warning;
        }
    }

    /// <summary>
    /// The Data Source WorkItem which is successfully migrated to TFS.
    /// It has additional ID information which is the ID of corresponding migrated workitem in TFS
    /// </summary>
    internal class PassedSourceWorkItem : SourceWorkItem
    {
        /// <summary>
        /// The ID of successfully migrated workitem in the TFS
        /// </summary>
        public int TFSId
        {
            get;
            set;
        }

        /// <summary>
        /// Constructor to initialzie workitem from other workitem
        /// </summary>
        public PassedSourceWorkItem(ISourceWorkItem dsWorkItem, int id)
            : base(dsWorkItem)
        {
            this.TFSId = id;
        }
    }

    internal class SkippedSourceWorkItem : SourceWorkItem
    {
        /// <summary>
        /// Constructor to initialzie workitem from other workitem
        /// </summary>
        public SkippedSourceWorkItem(ISourceWorkItem dsWorkItem)
            : base(dsWorkItem)
        { }
    }
}
