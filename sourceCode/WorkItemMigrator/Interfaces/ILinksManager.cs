//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Responsible for managing and creating links between different work item typws and accross sessions
    /// </summary>
    public interface ILinksManager : INotifyPropertyChanged
    {
        /// <summary>
        /// Update ID Mapping between Source Id and TFS ID
        /// </summary>
        /// <param name="sourceId"></param>
        /// <param name="id"></param>
        void UpdateIdMapping(string sourceId, WorkItemMigrationStatus id);

        /// <summary>
        /// Add Link in the list of links to create
        /// </summary>
        /// <param name="link"></param>
        void AddLink(ILink link);

        int SessionId { get; }

        IDictionary<string, IDictionary<string, WorkItemMigrationStatus>> WorkItemCategoryToIdMappings { get; }


        /// <summary>
        /// File Path of file containing the status of all links processed/remaining
        /// </summary>
        string LinksFilePath { get; }


        /// <summary>
        /// Creates links and save the status of successful/failed links
        /// </summary>
        void Save();


        void PublishReport();
    }
}
