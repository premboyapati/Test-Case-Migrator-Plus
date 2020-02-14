//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    /// <summary>
    /// Interface used by WITHelper to import workitems to TFS Project.
    /// It represents a particular TFS Workitem and takes the reponsibility of 
    /// creating, updating and saving it.
    /// </summary>
    public interface IWorkItem
    {
        /// <summary>
        /// Craetes Workitem by the help of the project and workitem type name
        /// </summary>
        /// <param name="project"></param>
        /// <param name="wiTypeName"></param>
        void Create();


        /// <summary>
        /// Updates the workitem. It sets the 'value' to the field having referenceName'fieldReferenceName'
        /// </summary>
        /// <param name="fieldName"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        void UpdateField(string fieldName, object value);


        /// <summary>
        /// Gets the Value of workitem's field having reference name 'fieldReferenceName'
        /// </summary>
        /// <param name="fieldReferenceName"></param>
        /// <returns></returns>
        object GetFieldValue(string fieldName);


        /// <summary>
        /// Saves the Workitem
        /// </summary>
        void Save();


        /// <summary>
        /// The embedded Workitem at TFS Server
        /// </summary>
        object WorkItem { get; }
    }
}
