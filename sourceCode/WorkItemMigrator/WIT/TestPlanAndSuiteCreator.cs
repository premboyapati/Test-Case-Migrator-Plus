//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using Microsoft.TeamFoundation.TestManagement.Client;

    internal class TestPlanAndSuiteCreator
    {
        #region Fields

        private ITestManagementTeamProject m_project;
        private Dictionary<string, ITestPlan> m_createdPlans;

        #endregion

        #region Constructor
        public TestPlanAndSuiteCreator(ITestManagementTeamProject project)
        {
            m_project = project;
            m_createdPlans = new Dictionary<string, ITestPlan>();
        }
        #endregion

        #region Public methods

        public void Add(int tfsId, string path)
        {
            string[] names = path.Split('\\');
            ITestPlan testPlan = null;

            // Create test plan from the first entry in the array.
            if ((names != null) && (names.Length > 0))
            {
                testPlan = CreateTestPlan(names[0]);
            }

            // Create test suites inside the plan.
            IStaticTestSuite parentTestSuite = null;
            if (testPlan != null)
            {
                parentTestSuite = testPlan.RootSuite;
            }

            for (int i = 1; i < names.Length; i++)
            {
                if (!string.IsNullOrEmpty(names[i]))
                {
                    string name = names[i];
                    ITestSuiteBase newSuite = FindSuite(parentTestSuite, name);
                    if (newSuite == null)
                    {
                        newSuite = CreateSuite(name);
                        if (newSuite != null)
                        {
                            parentTestSuite.Entries.Add(newSuite);
                        }
                    }

                    parentTestSuite = newSuite as IStaticTestSuite;
                }
            }

            // Add test case to the suite.
            AddTestToSuite(tfsId, parentTestSuite);
        }

        #endregion

        #region Private Methods

        private void AddTestToSuite(int tfsId, IStaticTestSuite parentTestSuite)
        {
            if (parentTestSuite != null)
            {
                ITestCase testCase = m_project.TestCases.Find(tfsId);
                if ((testCase != null) && !parentTestSuite.TestCases.Contains(testCase))
                {
                    parentTestSuite.Entries.AddCases(new List<ITestCase>() { testCase }, false);
                }
            }
        }

        private ITestPlan CreateTestPlan(string testPlanName)
        {
            ITestPlan testPlan = null;
            if (!string.IsNullOrEmpty(testPlanName))
            {
                testPlan = FindTestPlan(testPlanName);
                if (testPlan == null)
                {
                    // Create a new test plan
                    testPlan = m_project.TestPlans.Create();
                    testPlan.Name = testPlanName;
                    testPlan.Save();
                }
            }
            if (!m_createdPlans.ContainsKey(testPlanName))
            {
                m_createdPlans.Add(testPlanName, testPlan);
            }

            return testPlan;
        }

        private ITestPlan FindTestPlan(string testPlanName)
        {
            if (m_createdPlans.ContainsKey(testPlanName))
            {
                return m_createdPlans[testPlanName];
            }

            string query = string.Format("SELECT * FROM TestPlan WHERE PlanName = '{0}'", testPlanName);
            foreach (ITestPlan plan in m_project.TestPlans.Query(query))
            {
                // Return the first entry.
                return plan;
            }

            return null;
        }

        private static ITestSuiteBase FindSuite(IStaticTestSuite parentTestSuite, string name)
        {
            parentTestSuite.Refresh();
            foreach (ITestSuiteEntry testSuite in parentTestSuite.Entries)
            {
                if (string.Equals(name, testSuite.Title, StringComparison.CurrentCultureIgnoreCase))
                {
                    return testSuite.TestSuite;
                }
            }

            return null;
        }

        private IStaticTestSuite CreateSuite(string name)
        {
            IStaticTestSuite newSuite = m_project.TestSuites.CreateStatic();
            newSuite.Title = name;
            newSuite.Description = string.Empty;
            return newSuite;
        }
        #endregion
    }
}
