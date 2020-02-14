//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Collections.Generic;

    /// <summary>
    /// Contains complete relationship information of the workitems which are going to be migrated to the other workitems
    /// and test suite
    /// </summary>
    interface IRelationshipsInfo
    {
        /// <summary>
        /// Field Name Present in Data Source(Excel) which is containg the ID of Workitems
        /// </summary>
        string SourceIdField { get; set; }

        
        /// <summary>
        /// Field Name Present in Data Source(Excel) which is having the list of test suites which will contain this workitem(testcase)
        /// </summary>
        string TestSuiteField { get; set; }


        /// <summary>
        /// List of Link rules which will be applied to the workitems present in the data source
        /// </summary>
        IList<ILinkRule> LinkRules { get; }
    }
}
