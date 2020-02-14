//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Collections.Generic;

    /// <summary>
    /// This delegate is called just before saving a tfs workitem by the Migrator.
    /// Gives the handler flexibility to modify the loaded workitem before save
    /// </summary>
    /// <param name="workItem">TFS Workitem</param>
    /// <returns>True if save workitem; false if skip workitem</returns>
    public delegate bool PreMigrationEvent(ISourceWorkItem sourceWorkItem, IWorkItem workItem);


    /// <summary>
    /// This delegate is called after workitem save by the Migrator. used for updating the counters or cancelling the migration
    /// </summary>
    /// <param name="sourceWorkItem"></param>
    /// <returns>false is to stop migration</returns>
    public delegate bool PostMigrationEvent(ISourceWorkItem sourceWorkItem);


    /// <summary>
    /// Interface responsible for Migrating the source workitems into TFS
    /// </summary>
    public interface IMigrator
    {
        /// <summary>
        /// Mapping from source field name to corresponding IWorkitemField
        /// </summary>
        IDictionary<string, IWorkItemField> SourceNameToFieldMapping { get; set; }


        /// <summary>
        /// List of SourceWorkitems without parameterization/multiline settings
        /// </summary>
        IList<ISourceWorkItem> RawSourceWorkItems { get; set; }


        /// <summary>
        /// List of Source Workitem with parsing settings applied. Setter should generate FieldToUniqueValues
        /// </summary>
        IList<ISourceWorkItem> SourceWorkItems { get; set; }


        /// <summary>
        /// Mapping from field to all uniques values found in the source for this field
        /// </summary>
        IDictionary<string, IList<string>> FieldToUniqueValues { get; }


        /// <summary>
        /// This is called just before saving a tfs workitem by the Migrator.
        /// Gives the handler flexibility to modify the loaded workitem before save
        /// </summary>
        PreMigrationEvent PreMigration { get;  set; }

        /// <summary>
        /// Used for updating the counters or cancelling the migration
        /// </summary>
        PostMigrationEvent PostMigration { get; set; }


        /// <summary>
        /// Migrates 'SourceWorkitems' with the help of WorkItemGenerater and return the list of 
        /// result source workitems(passed/failed/warning)
        /// </summary>
        /// <param name="workItemGenerator"></param>
        /// <returns></returns>
        IList<ISourceWorkItem> Migrate(IWorkItemGenerator workItemGenerator);

    }
}
