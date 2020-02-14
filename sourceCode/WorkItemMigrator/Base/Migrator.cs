//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Collections.Generic;

    /// <summary>
    /// Main class reponsible for migrating Work Item from source to TCM.
    /// </summary>
    internal class Migrator : IMigrator
    {
        #region Fields

        // list of source workitems which are to be migrated
        private IList<ISourceWorkItem> m_sourceWorkItems;

        #endregion

        #region Properties

        /// <summary>
        /// Mapping from Source Field Name to corresponding Field
        /// </summary>
        public IDictionary<string, IWorkItemField> SourceNameToFieldMapping
        {
            get;
            set;
        }

        /// <summary>
        /// List of Source Workitems
        /// </summary>
        public IList<ISourceWorkItem> SourceWorkItems
        {
            get
            {
                return m_sourceWorkItems;
            }
            set
            {
                m_sourceWorkItems = value;
                InitializeUniqueValues();
            }
        }

        /// <summary>
        /// List of Source WorkItems with no parsing settings(parameterization/multiline) applied.
        /// </summary>
        public IList<ISourceWorkItem> RawSourceWorkItems
        {
            get;
            set;
        }

        /// <summary>
        /// Mapping from field to all its unique values found in the source
        /// </summary>
        public IDictionary<string, IList<string>> FieldToUniqueValues
        {
            get;
            private set;
        }

        /// <summary>
        /// This delegate is called just before saving the workitem at destination
        /// </summary>
        public PreMigrationEvent PreMigration
        {
            get;
            set;
        }

        /// <summary>
        /// This delegate is called after saving the worrktem after saving the workitem at the server
        /// </summary>
        public PostMigrationEvent PostMigration
        {
            get;
            set;
        }

        #endregion

        #region public Methods

        /// <summary>
        /// Migrates all the workitems with the help of WorkItemGenerator and returns the list of updated source workitems
        /// </summary>
        /// <param name="workItemGenerator"></param>
        /// <returns></returns>
        public IList<ISourceWorkItem> Migrate(IWorkItemGenerator workItemGenerator)
        {
            // The List of source workitems which has to be returned
            var resultSourceWorkItems = new List<ISourceWorkItem>();

            // Now process each source workitems which are to migrate
            foreach (ISourceWorkItem sourceWorkItem in SourceWorkItems)
            {
                if (sourceWorkItem is FailedSourceWorkItem)
                {
                    AddSourceWorkItemInResultSourceWorkItems(resultSourceWorkItems, sourceWorkItem);
                    continue;
                }

                // bool variable to track the return value of Premigration event
                bool preMigrationReturnValue = true;

                // bool variable to track the return value of PostMigration event
                bool postMigrationEventValue = true;

                IWorkItem workItem = null;
                try
                {
                    // Create IWorkitem with the help of WorkItemGenerator
                    workItem = workItemGenerator.CreateWorkItem(sourceWorkItem);
                }
                catch (WorkItemMigratorException ex)
                {
                    var failedWorkItem = new FailedSourceWorkItem(sourceWorkItem, ex.Args.Title);
                    AddSourceWorkItemInResultSourceWorkItems(resultSourceWorkItems, failedWorkItem);
                    continue;
                }

                if (PreMigration != null)
                {
                    preMigrationReturnValue = PreMigration(sourceWorkItem, workItem);
                }

                // If PreMigration allowed to process further then save the workitem at destination
                if (preMigrationReturnValue)
                {
                    // Saves the workitem at Destination and get the updated sourceworkitem
                    var migratedSourceWorkItem = workItemGenerator.SaveWorkItem(workItem, sourceWorkItem);

                    postMigrationEventValue = AddSourceWorkItemInResultSourceWorkItems(resultSourceWorkItems, migratedSourceWorkItem);

                    // If PostMigration event has return false to terminate the further migration
                    if (!postMigrationEventValue)
                    {
                        break;
                    }
                }
                // If preMigration event returned false the skip this sourceworkitem
                else
                {
                    AddSourceWorkItemInResultSourceWorkItems(resultSourceWorkItems, new SkippedSourceWorkItem(sourceWorkItem));
                }
            }

            // returning the list of updated sourceworkitem
            return resultSourceWorkItems;
        }

        private bool AddSourceWorkItemInResultSourceWorkItems(List<ISourceWorkItem> resultSourceWorkItems, ISourceWorkItem sourceWorkItem)
        {
            bool postMigrationEventValue = true;
            resultSourceWorkItems.Add(sourceWorkItem);
            if (PostMigration != null)
            {
                postMigrationEventValue = PostMigration(sourceWorkItem);
            }
            return postMigrationEventValue;
        }

        #endregion

        #region Private methods

        private void InitializeUniqueValues()
        {
            FieldToUniqueValues = new Dictionary<string, IList<string>>();
            foreach (var sourceWorkItem in SourceWorkItems)
            {
                foreach (var kvp in sourceWorkItem.FieldValuePairs)
                {
                    string value = kvp.Value as string;
                    if (string.IsNullOrEmpty(value))
                    {
                        continue;
                    }

                    if (SourceNameToFieldMapping.ContainsKey(kvp.Key) &&
                        SourceNameToFieldMapping[kvp.Key].HasAllowedValues)
                    {
                        if (!FieldToUniqueValues.ContainsKey(kvp.Key))
                        {
                            FieldToUniqueValues.Add(kvp.Key, new List<string>());
                        }
                        if (!FieldToUniqueValues[kvp.Key].Contains(value))
                        {
                            FieldToUniqueValues[kvp.Key].Add(value);
                        }
                    }
                }
            }
        }

        #endregion
    }
}
