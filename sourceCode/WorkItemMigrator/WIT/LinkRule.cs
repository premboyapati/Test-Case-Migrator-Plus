//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    internal class LinkRule : ILinkRule
    {
        public string StartWorkItemCategory
        {
            get;
            private set;
        }

        public string SourceFieldNameOfEndWorkItemCategory
        {
            get;
            private set;
        }

        public string LinkTypeReferenceName
        {
            get;
            private set;
        }

        public string EndWorkItemCategory
        {
            get;
            private set;
        }

        public string Description
        {
            get;
            private set;
        }

        public LinkRule(string startworkItemCategory, string sourceFieldName, string linkTypeName, string endWorkItemCategory, string description)
        {
            StartWorkItemCategory = startworkItemCategory;
            SourceFieldNameOfEndWorkItemCategory = sourceFieldName;
            LinkTypeReferenceName = linkTypeName;
            EndWorkItemCategory = endWorkItemCategory;
            Description = description;
        }
    }
}
