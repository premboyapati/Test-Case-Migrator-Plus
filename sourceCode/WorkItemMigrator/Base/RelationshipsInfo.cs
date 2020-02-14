//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;

    internal class RelationshipsInfo : IRelationshipsInfo
    {
        #region Properties
        public string SourceIdField
        {
            get;
            set;
        }

        public string TestSuiteField
        {
            get;
            set;
        }

        public IList<ILinkRule> LinkRules
        {
            get;
            private set;
        }


        public char[] Delimeters
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }
        #endregion

        #region Constructor

        public RelationshipsInfo()
        {
            LinkRules = new List<ILinkRule>();
        }

        #endregion
    }
}
