//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using Microsoft.TeamFoundation;
    using Microsoft.TeamFoundation.Client;
    using Microsoft.TeamFoundation.Framework.Client;
    using Microsoft.TeamFoundation.TestManagement.Client;
    using Microsoft.TeamFoundation.WorkItemTracking.Client;

    internal class WorkItemGenerator : NotifyPropertyChange, IWorkItemGenerator
    {
        #region Fields

        private string m_serverName;
        private string m_projectName;        
        private ITestManagementTeamProject m_teamProject;
        private string m_workItemTypeName;
        private TestPlanAndSuiteCreator m_suiteCreater;
        private WorkItemLinkTypeEndCollection m_workItemLinkTypeEndCollection;
        private AreaAndIterationPathCreator m_areaAndIterationPathCreator;
        private WorkItem m_linkingTask;

        #endregion

        #region Constants

        // TestCase Category Reference Name
        public const string TestCaseCategory = "Microsoft.TestCaseCategory";

        public const string BugCategory = "Microsoft.BugCategory";

        public const string RequirementCategory = "Microsoft.RequirementCategory";

        // Refernce Name prefix for TCM Fields
        private const string TCMField = "Microsoft.VSTS.TCM";

        // Reference Name for WIT Field: Test Steps
        private const string StepsReferenceName = "Microsoft.VSTS.TCM.Steps";

        public const string TestsLinkReferenceName = "Tests";

        public const string TestedByLinkRefernceName = "Tested By";

        public const string ParentLinkReferenceName = "Parent";

        public const string RelatedLinkReferenceName = "Related";

        public const string LinkingTaskTitle = "<DONOTDELETEORCHANGE><TestCaseMigratorPlus><LinksMetaData><Stored as a Task work item>";

        #endregion

        #region Constructor

        public WorkItemGenerator(string serverName, string projectName)
        {
            if (string.IsNullOrEmpty(serverName))
            {
                throw new ArgumentNullException("serverName", "servername is null");
            }
            if (string.IsNullOrEmpty(projectName))
            {
                throw new ArgumentNullException("projectName", "projectName is null");
            }
         
            TeamProjectCollection = new TfsTeamProjectCollection(new Uri(serverName), new Services.Common.VssCredentials(true)/* new UICredentialsProvider()*/);        
       
            // Get the Test Management Service
            ITestManagementService service = (ITestManagementService)TeamProjectCollection.GetService(typeof(ITestManagementService));

            // Initialize the TestManagement Team Project
            m_teamProject = (ITestManagementTeamProject)service.GetTeamProject(projectName);

            ILocationService locationService = TeamProjectCollection.GetService<ILocationService>();

            IsTFS2012 = !String.IsNullOrEmpty(locationService.LocationForCurrentConnection("IdentityManagementService2",
                                                                                         new Guid("A4CE4577-B38E-49C8-BDB4-B9C53615E0DA")));

            // Set the Properties
            Server = serverName;
            Project = projectName;

            SourceNameToFieldMapping = new Dictionary<string, IWorkItemField>();

            m_suiteCreater = new TestPlanAndSuiteCreator(m_teamProject);
            m_areaAndIterationPathCreator = new AreaAndIterationPathCreator(m_teamProject.WitProject);
            WorkItemCategoryToDefaultType = new Dictionary<string, string>();

            LinkTypeNames = new List<string>();
            m_workItemLinkTypeEndCollection = m_teamProject.WitProject.Store.WorkItemLinkTypes.LinkTypeEnds;
            foreach (WorkItemLinkTypeEnd linkTypeEnd in m_teamProject.WitProject.Store.WorkItemLinkTypes.LinkTypeEnds)
            {
                string linkTypeName = linkTypeEnd.Name;

                LinkTypeNames.Add(linkTypeName);
            }

            WorkItemTypeToCategoryMapping = new Dictionary<string, string>();
            WorkItemTypeNames = new List<string>();

            PopulateWorkItemTypeDetailsFromCategory(m_teamProject.WitProject.Categories, BugCategory);
            PopulateWorkItemTypeDetailsFromCategory(m_teamProject.WitProject.Categories, RequirementCategory);

            Category testCaseCategory = m_teamProject.WitProject.Categories[TestCaseCategory];
            if (testCaseCategory != null)
            {
                DefaultWorkItemTypeName = testCaseCategory.DefaultWorkItemType.Name;
                PopulateWorkItemTypeDetailsFromCategory(m_teamProject.WitProject.Categories, TestCaseCategory);
            }
                        
            CreateAreaIterationPath = true;
        }

        #endregion

        #region Properties

        public bool IsTFS2012
        {
            get;
            private set;
        }

        /// <summary>
        /// TFS Server Collection URL
        /// </summary>
        public string Server
        {
            get
            {
                return m_serverName;
            }
            private set
            {
                m_serverName = value;
                NotifyPropertyChanged("Server");
            }
        }

        /// <summary>
        /// TFS Project Name
        /// </summary>
        public string Project
        {
            get
            {
                return m_projectName;
            }
            private set
            {
                m_projectName = value;
                NotifyPropertyChanged("Project");
            }
        }

        public IList<string> LinkTypeNames
        {
            get;
            private set;
        }

        /// <summary>
        /// List of Name of WorkItem types
        /// </summary>
        public IList<string> WorkItemTypeNames
        {
            get;
            private set;
        }

        public string WorkItemCategory
        {
            get
            {
                return WorkItemTypeToCategoryMapping[SelectedWorkItemTypeName];
            }
        }

        public IDictionary<string, string> WorkItemTypeToCategoryMapping
        {
            get;
            private set;
        }

        public IDictionary<string, string> WorkItemCategoryToDefaultType
        {
            get;
            private set;
        }

        /// <summary>
        /// Default work item type in the selected workitem category
        /// </summary>
        public string DefaultWorkItemTypeName
        {
            get;
            private set;
        }

        /// <summary>
        /// Work Item Type of the TFS workitems to generates
        /// </summary>
        public string SelectedWorkItemTypeName
        {
            get
            {
                return m_workItemTypeName;
            }
            set
            {
                m_workItemTypeName = value;
                InitializeWorkItemFields();
                NotifyPropertyChanged("SelectedWorkItemTypeName");
            }
        }

        /// <summary>
        /// Whether to ass 'test step title' and 'test step expected result' field.
        /// </summary>
        public bool AddTestStepsField
        {
            get;
            set;
        }

        /// <summary>
        /// TFS Name to FIeld Mapping
        /// </summary>
        public IDictionary<string, IWorkItemField> TfsNameToFieldMapping
        {
            get;
            private set;
        }

        /// <summary>
        /// Source Name to FIeld Mapping for selected source Fields 
        /// </summary>
        public IDictionary<string, IWorkItemField> SourceNameToFieldMapping
        {
            get;
            set;
        }

        public WorkItem LinkingTask
        {
            get
            {
                if (m_linkingTask == null)
                {
                    try
                    {
                        string query = "SELECT * FROM WorkItem WHERE [System.Title] = '" + LinkingTaskTitle + "' AND  [System.TeamProject] = '" + Project + "'";
                        WorkItemCollection result = m_teamProject.WitProject.Store.Query(query);

                        if (result != null && result.Count > 0)
                        {
                            m_linkingTask = result[0];
                        }
                        else
                        {
                            m_linkingTask = m_teamProject.WitProject.WorkItemTypes["Task"].NewWorkItem();
                            m_linkingTask.Title = LinkingTaskTitle;
                            m_linkingTask.Save();
                        }
                    }
                    catch (TeamFoundationServerException te)
                    {
                        throw new WorkItemMigratorException(te.Message, null, null);
                    }
                }
                return m_linkingTask;
            }
        }

        public bool CreateAreaIterationPath
        {
            get;
            set;
        }
        
        /// <summary>
        /// TfsTeamProjectCollection object
        /// </summary>
        public TfsTeamProjectCollection TeamProjectCollection
        {
            get;
            private set;
        }

        #endregion

        #region Public Methods
           
        /// <summary>
        /// Creates Workitem from Source Workitem
        /// </summary>
        /// <param name="sourceWorkItem"></param>
        /// <returns></returns>
        public IWorkItem CreateWorkItem(ISourceWorkItem sourceWorkItem)
        {
            if (sourceWorkItem == null)
            {
                throw new ArgumentNullException("sourceWorkItem", "source work item");
            }

            // Initializing the variables
            IWorkItem workItem = null;

            // Initializing the Workitem
            switch (WorkItemTypeToCategoryMapping[SelectedWorkItemTypeName])
            {
                case TestCaseCategory:
                    TestCase testCase = new TestCase(m_teamProject, SelectedWorkItemTypeName, IsTFS2012);
                    workItem = testCase;
                    break;

                default:
                    workItem = new WorkItemBase(m_teamProject.WitProject, SelectedWorkItemTypeName);
                    break;
            }

            workItem.Create();

            // iterating through each field-value pair
            foreach (KeyValuePair<string, object> kvp in sourceWorkItem.FieldValuePairs)
            {
                // Refernce Name of the field which is to be updated
                string fieldName = kvp.Key;

                // value of the field which has to assigned
                object fieldValue = kvp.Value;

                // Gets the correct field reference name
                if (SourceNameToFieldMapping.ContainsKey(kvp.Key))
                {
                    fieldName = SourceNameToFieldMapping[kvp.Key].TfsName;

                    if (CreateAreaIterationPath && SourceNameToFieldMapping[kvp.Key].IsAreaPath)
                    {
                        // Create area path if it does not already exist.
                        fieldValue = m_areaAndIterationPathCreator.Create(Node.TreeType.Area, fieldValue as string);
                        SourceNameToFieldMapping[kvp.Key].ValueMapping.Clear();
                    }

                    if (CreateAreaIterationPath && SourceNameToFieldMapping[kvp.Key].IsIterationPath)
                    {
                        // Create area path if it does not already exist.
                        fieldValue = m_areaAndIterationPathCreator.Create(Node.TreeType.Iteration, fieldValue as string);
                        SourceNameToFieldMapping[kvp.Key].ValueMapping.Clear();
                    }


                    // If it is autogenerated value then just ignore the assignment of value and continue to next iteration
                    if (SourceNameToFieldMapping[kvp.Key].IsAutoGenerated)
                    {
                        continue;
                    }
                }

                string value = kvp.Value as string;

                // value to assign is of type string and it is not null
                if (value != null)
                {
                    // If this value is dataMapped
                    if (SourceNameToFieldMapping.ContainsKey(kvp.Key) && SourceNameToFieldMapping[kvp.Key].ValueMapping.ContainsKey(value))
                    {
                        if (String.Compare(SourceNameToFieldMapping[kvp.Key].ValueMapping[value], Resources.IgnoreLabel, StringComparison.CurrentCulture) == 0)
                        {
                            continue;
                        }
                        else
                        {
                            // else sets the correct data value
                            fieldValue = SourceNameToFieldMapping[kvp.Key].ValueMapping[value];
                        }
                    }
                }

                // Update the workitem with field-value
                workItem.UpdateField(fieldName, fieldValue);
            }
            // return the workitem
            return workItem;
        }

        /// <summary>
        /// Save TFS Workitem and return result workitem(passed,failed or warning) depending upon the result of migration.
        /// </summary>
        /// <param name="workItem"></param>
        /// <param name="sourceWorkItem"></param>
        /// <returns></returns>
        public ISourceWorkItem SaveWorkItem(IWorkItem workItem, ISourceWorkItem sourceWorkItem)
        {
            try
            {
                workItem.Save();
            }
            catch (WorkItemMigratorException te)
            {
                return new FailedSourceWorkItem(sourceWorkItem, te.Args.Title);
            }

            ISourceWorkItem updatedSourceWorkItem = new SourceWorkItem(sourceWorkItem);
            updatedSourceWorkItem.FieldValuePairs.Clear();
            string warning = null;
            int id = -1;

            switch (WorkItemTypeToCategoryMapping[SelectedWorkItemTypeName])
            {
                case TestCaseCategory:
                    id = ((ITestCase)(workItem.WorkItem)).Id;
                    break;

                default:
                    id = ((WorkItem)(workItem.WorkItem)).Id;
                    break;
            }


            foreach (var kvp in sourceWorkItem.FieldValuePairs)
            {
                string previousValue = kvp.Value as string;
                if (!string.IsNullOrEmpty(previousValue))
                {
                    if (SourceNameToFieldMapping[kvp.Key].ValueMapping.ContainsKey(previousValue) &&
                        String.CompareOrdinal(SourceNameToFieldMapping[kvp.Key].ValueMapping[previousValue], Resources.IgnoreLabel) != 0)
                    {
                        previousValue = SourceNameToFieldMapping[kvp.Key].ValueMapping[previousValue];
                    }
                    string tfsFieldName = SourceNameToFieldMapping[kvp.Key].TfsName;
                    string newValue = workItem.GetFieldValue(tfsFieldName).ToString();
                    if (String.CompareOrdinal(previousValue, newValue) != 0 &&
                        !SourceNameToFieldMapping[kvp.Key].IsHtmlField &&
                        !SourceNameToFieldMapping[kvp.Key].IsAreaPath &&
                        !SourceNameToFieldMapping[kvp.Key].IsIterationPath)
                    {
                        warning += String.Format(CultureInfo.CurrentCulture,
                                                 "Value of Field:{0} is modified from \n'{1}'\n to \n'{2}'",
                                                 kvp.Key,
                                                 previousValue,
                                                 newValue);
                    }
                    updatedSourceWorkItem.FieldValuePairs.Add(kvp.Key, newValue);
                }
                else
                {
                    string fieldName = kvp.Key;
                    if (SourceNameToFieldMapping.ContainsKey(kvp.Key))
                    {
                        fieldName = SourceNameToFieldMapping[kvp.Key].TfsName;
                    }

                    List<SourceTestStep> testSteps = kvp.Value as List<SourceTestStep>;
                    if (testSteps != null)
                    {
                        updatedSourceWorkItem.FieldValuePairs.Add(kvp.Key, testSteps);
                    }
                    else
                    {
                        updatedSourceWorkItem.FieldValuePairs.Add(kvp.Key, workItem.GetFieldValue(fieldName));
                    }
                }
            }
            if (string.IsNullOrEmpty(warning))
            {
                return new PassedSourceWorkItem(updatedSourceWorkItem, id);
            }
            else
            {
                return new WarningSourceWorkItem(updatedSourceWorkItem, id, warning);
            }
        }

        public void AddWorkItemToTestSuite(int tfsId, string testSuite)
        {
            m_suiteCreater.Add(tfsId, testSuite);
        }

        public void CreateLinksInBatch(IList<ILink> links)
        {
            Dictionary<int, WorkItem> workItemIdToWorkItem = new Dictionary<int, WorkItem>();
            Dictionary<WorkItem, IList<ILink>> workItemToLinks = new Dictionary<WorkItem, IList<ILink>>();
            foreach (var link in links)
            {
                if (link.StartWorkItemTfsId <= 0
                    || link.EndWorkItemTfsId <= 0 ||
                    link.StartWorkItemTfsId == link.EndWorkItemTfsId)
                {
                    continue;
                }
                if (!workItemIdToWorkItem.ContainsKey(link.StartWorkItemTfsId))
                {
                    workItemIdToWorkItem.Add(link.StartWorkItemTfsId, m_teamProject.WitProject.Store.GetWorkItem(link.StartWorkItemTfsId));
                }

                WorkItem workItem = workItemIdToWorkItem[link.StartWorkItemTfsId];
                if (!workItemToLinks.ContainsKey(workItem))
                {
                    workItemToLinks.Add(workItem, new List<ILink>());
                }

                bool isAlreadyLinked = false;
                foreach (WorkItemLink wiLinkType in workItem.WorkItemLinks)
                {
                    if (String.CompareOrdinal(wiLinkType.LinkTypeEnd.Name, link.LinkTypeName) == 0 &&
                        wiLinkType.TargetId == link.EndWorkItemTfsId)
                    {
                        isAlreadyLinked = true;
                        break;
                    }
                }
                if (isAlreadyLinked)
                {
                    continue;
                }

                try
                {
                    WorkItemLinkTypeEnd workItemLinkTypeEnd = m_workItemLinkTypeEndCollection[link.LinkTypeName];
                    WorkItemLink workItemLink = new WorkItemLink(workItemLinkTypeEnd, link.EndWorkItemTfsId);
                    workItem.WorkItemLinks.Add(workItemLink);
                    workItemToLinks[workItem].Add(link);
                }
                catch (ValidationException)
                { }
            }
            int counter = 0;
            foreach (var kvp in workItemIdToWorkItem)
            {
                WorkItem workItem = kvp.Value;
                try
                {
                    if (workItem.IsDirty)
                    {
                        workItem.Save();
                        if (workItemToLinks.ContainsKey(kvp.Value))
                        {
                            foreach (var link in workItemToLinks[kvp.Value])
                            {
                                link.IsExistInTfs = true;
                            }
                        }
                    }
                }
                catch (ValidationException)
                {
                    CreateLinks(workItem, workItemToLinks[workItem]);
                }
                catch (TeamFoundationServerException)
                {
                    CreateLinks(workItem, workItemToLinks[workItem]);
                }
                counter++;
                Console.Write("\rCreating Links...{0}% completed", (counter * 100 / workItemIdToWorkItem.Count));
            }
            Console.Write("\rCreating Links...100% completed   ");
        }

        public bool IsWitExists(int witId)
        {
            try
            {
                m_teamProject.WitProject.Store.GetWorkItem(witId);
                return true;
            }
            catch (TeamFoundationServerException)
            {
                return false;
            }
        }

        public void RemoveLink(int sourceWorkItemId, int targetWorkItemId)
        {
            try
            {
                WorkItem workItem = m_teamProject.WitProject.Store.GetWorkItem(sourceWorkItemId);
                var deletedLinks = new List<WorkItemLink>();
                foreach (WorkItemLink link in workItem.WorkItemLinks)
                {
                    if (link.TargetId == targetWorkItemId)
                    {
                        deletedLinks.Add(link);
                    }
                }
                foreach (var link in deletedLinks)
                {
                    workItem.WorkItemLinks.Remove(link);
                }
                workItem.Save();
            }
            catch (TeamFoundationServerException)
            { }
        }

        #endregion

        #region private methods

        private void PopulateWorkItemTypeDetailsFromCategory(CategoryCollection categories, string categoryName)
        {
            if (categories.Contains(categoryName))
            {
                Category cat = categories[categoryName];
                WorkItemCategoryToDefaultType.Add(cat.ReferenceName, cat.DefaultWorkItemType.Name);
                
                foreach (WorkItemType workItemType in cat.WorkItemTypes)
                {
                    if (!WorkItemTypeToCategoryMapping.ContainsKey(workItemType.Name))
                    {
                        WorkItemTypeToCategoryMapping.Add(workItemType.Name, cat.ReferenceName);
                        WorkItemTypeNames.Add(workItemType.Name);
                    }
                }
            }
        }

        private void CreateLinks(WorkItem workItem, IList<ILink> links)
        {
            workItem.Reset();
            foreach (Link link in links)
            {
                try
                {
                    WorkItemLinkTypeEnd workItemLinkTypeEnd = m_workItemLinkTypeEndCollection[link.LinkTypeName];
                    WorkItemLink workItemLink = new WorkItemLink(workItemLinkTypeEnd, link.EndWorkItemTfsId);
                    workItem.WorkItemLinks.Add(workItemLink);
                    workItem.Save();
                    link.IsExistInTfs = true;
                }
                catch (ValidationException valEx)
                {
                    link.Message = valEx.Message;
                    workItem.Reset();
                }
                catch (TeamFoundationServerException tfsEx)
                {
                    link.Message = tfsEx.Message;
                    workItem.Reset();
                }
            }
        }

        private void InitializeWorkItemFields()
        {
            TfsNameToFieldMapping = new Dictionary<string, IWorkItemField>();

            if (string.IsNullOrEmpty(SelectedWorkItemTypeName))
            {
                return;
            }
            WorkItemType workItemType = m_teamProject.WitProject.WorkItemTypes[SelectedWorkItemTypeName];

            WorkItem wi = workItemType.NewWorkItem();

            // iterating through each field present in the work item type
            foreach (Field field in wi.Fields)
            {
                // Steps Field needs special handling
                if (String.CompareOrdinal(field.FieldDefinition.ReferenceName, StepsReferenceName) == 0)
                {
                    if (AddTestStepsField)
                    {
                        WorkItemField testStepTitleField = new TestStepTitleField();
                        WorkItemField testStepExpectedResultField = new TestStepExpectedResultField();
                        TfsNameToFieldMapping.Add(testStepTitleField.TfsName, testStepTitleField);
                        TfsNameToFieldMapping.Add(testStepExpectedResultField.TfsName, testStepExpectedResultField);
                    }
                    else
                    {
                        IWorkItemField wiField = new WorkItemField(field, m_teamProject.WitProject);
                        TfsNameToFieldMapping.Add(wiField.TfsName, wiField);
                    }

                }
                // else if it is not a TCM Field
                else if (String.CompareOrdinal(WorkItemTypeToCategoryMapping[SelectedWorkItemTypeName], TestCaseCategory) != 0 ||
                         !field.FieldDefinition.ReferenceName.Contains(TCMField))
                {
                    IWorkItemField wiField = new WorkItemField(field, m_teamProject.WitProject);
                    TfsNameToFieldMapping.Add(wiField.TfsName, wiField);
                }
            }
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose()
        {
            try
            {
                if (TeamProjectCollection != null)
                {
                    TeamProjectCollection.Dispose();
                }
            }
            catch (TeamFoundationServerException)
            { }
        }

        #endregion

    }
}
