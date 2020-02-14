//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using Microsoft.TeamFoundation.WorkItemTracking.Client;
    using Microsoft.TeamFoundation.Client;

    /// <summary>
    /// Reponsible for all TFS related activities
    /// </summary>
    public interface IWorkItemGenerator : IDisposable, INotifyPropertyChanged
    {
        /// <summary>
        /// TFS Server Collection URL
        /// </summary>
        string Server { get; }


        /// <summary>
        /// Team Project Name
        /// </summary>
        string Project { get; }


        /// <summary>
        /// Gets the name of all workitem types in a Workitem Category
        /// </summary>
        /// <param name="wiCategaory"></param>
        /// <returns></returns>
        IList<string> WorkItemTypeNames { get; }


        /// <summary>
        /// Returns the name of the default workitem type in a workitem category
        /// </summary>
        /// <param name="wiCategaory"></param>
        /// <returns></returns>
        string DefaultWorkItemTypeName { get; }


        /// <summary>
        /// Slected WorkItem Type's Name
        /// </summary>
        string SelectedWorkItemTypeName { get; set; }


        string WorkItemCategory { get; }

        IDictionary<string, string> WorkItemCategoryToDefaultType { get; }

        IDictionary<string, string> WorkItemTypeToCategoryMapping { get; }

        /// <summary>
        /// Whether to add 'Test Step Title' and 'Test Step Expected Result' FIeld. Required for Excel migration.
        /// </summary>
        bool AddTestStepsField { get; set; }


        /// <summary>
        /// Hashtable of WorkItemFields by FieldName
        /// </summary>
        IDictionary<string, IWorkItemField> TfsNameToFieldMapping { get; }


        /// <summary>
        /// Mapping from source field name to 'IWorkItemField'
        /// </summary>
        IDictionary<string, IWorkItemField> SourceNameToFieldMapping { get; set; }


        IList<string> LinkTypeNames { get; }

        /// <summary>
        /// Creates 'IWorkItem'(TFS WorkItem Wrapper) object from Source Workitem
        /// </summary>
        /// <param name="dsWorkItem"></param>
        /// <returns></returns>
        IWorkItem CreateWorkItem(ISourceWorkItem dsWorkItem);


        bool CreateAreaIterationPath { get; set; }


        /// <summary>
        /// Saves TFS Workitem and return Result WorkItem(Passed/Failed/Warning) depending upon result of migration.
        /// </summary>
        /// <param name="workItem"></param>
        /// <param name="dataSourceWorkItem"></param>
        /// <returns></returns>
        ISourceWorkItem SaveWorkItem(IWorkItem workItem, ISourceWorkItem dataSourceWorkItem);

        WorkItem LinkingTask { get; }

        void AddWorkItemToTestSuite(int tfsId, string testSuite);

        void CreateLinksInBatch(IList<ILink> links);

        bool IsWitExists(int witId);

        void RemoveLink(int sourceWorkItemId, int targetWorkItemId);
               
        TfsTeamProjectCollection TeamProjectCollection { get; }

        bool IsTFS2012 { get; }
    }
}
