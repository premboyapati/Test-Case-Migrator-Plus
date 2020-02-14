//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.ComponentModel;

    /// <summary>
    /// Responsible for generating Report file
    /// </summary>
    interface IReporter : IDisposable, INotifyPropertyChanged
    {
        /// <summary>
        /// Path of the report file
        /// </summary>
        string ReportFile { get; set; }


        /// <summary>
        /// Add entry in the report file
        /// </summary>
        /// <param name="sourceWorkItem"></param>
        void AddEntry(ISourceWorkItem sourceWorkItem);


        /// <summary>
        /// Publish the report at specified path
        /// </summary>
        void Publish();
    }
}
