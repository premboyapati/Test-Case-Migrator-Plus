//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Runtime.Serialization;

    /// <summary>
    /// This is used by Model Classes to send exeptions to View model whenever the data provided to them
    /// by ViewModel is incorrect.
    /// </summary>
    [Serializable]
    internal class WorkItemMigratorException : Exception
    {
        #region Properties

        /// <summary>
        /// Message Event Args: Used by View Model to show messages
        /// </summary>
        public MessageEventArgs Args
        {
            get;
            private set;
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Basic Constructor to initialize the Args
        /// </summary>
        public WorkItemMigratorException()
        {
            Args = new MessageEventArgs();
            Args.Category = MessageCategory.Error;
        }

        /// <summary>
        /// Constructor to set the title, likely cause and Potential Solution of the Exception
        /// </summary>
        /// <param name="errorTitle"></param>
        /// <param name="likelyCause"></param>
        /// <param name="potentialSolution"></param>
        public WorkItemMigratorException(string errorTitle, string likelyCause, string potentialSolution)
            : this()
        {
            Args.Title = errorTitle;
            Args.LikelyCause = likelyCause;
            Args.PotentialSolution = potentialSolution;
        }

        protected WorkItemMigratorException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        { }

        #endregion
    }
}
