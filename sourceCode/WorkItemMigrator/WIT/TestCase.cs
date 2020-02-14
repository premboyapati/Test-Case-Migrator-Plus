//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using Microsoft.TeamFoundation.TestManagement.Client;

    /// <summary>
    /// Wrapper over TCM Test Case
    /// </summary>
    internal class TestCase : WorkItemBase
    {
        #region Fields

        private ITestManagementTeamProject m_tcmProject;

        private string m_testCaseTypeName;

        private string m_stepsFieldName;

        // ITestCase object to save workitem as TCM Testcase
        private ITestCase m_testCase;
        
        private bool m_areRichSteps;
        #endregion

        #region Constants

        // The character used by TCM for identifying parameter
        private const string TCMParameterizationCharacter = "@";

        // White space character
        private const string WhiteSpaceCharacter = " ";

        #endregion

        #region Constructor

        public TestCase(ITestManagementTeamProject project, string testCaseTypeName, bool areRichSteps)
        {
            m_tcmProject = project;
            m_testCaseTypeName = testCaseTypeName;
            m_areRichSteps = areRichSteps;
        }

        #endregion

        #region Properties

        public override object WorkItem
        {
            get
            {
                return m_testCase;
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// creates TCM Testcase
        /// </summary>
        public override void Create()
        {
            m_testCase = m_tcmProject.TestCases.Create(m_tcmProject.WitProject.WorkItemTypes[m_testCaseTypeName]);
            base.m_workItem = m_testCase.WorkItem;
        }

        /// <summary>
        /// Updates testcase's field with corresponding value
        /// </summary>
        /// <param name="fieldName"></param>
        /// <param name="value"></param>
        public override void UpdateField(string fieldName, object value)
        {
            // If value is null then just take the default valuea nd return the default value
            if (value == null)
            {
                return;
            }

            List<SourceTestStep> sourceSteps = value as List<SourceTestStep>;

            // If the value is a List of test steps then Update th testcase with test steps
            if (sourceSteps != null)
            {
                m_stepsFieldName = fieldName;

                foreach (SourceTestStep sourceStep in sourceSteps)
                {
                    // Update the Test step's text with correct parameters
                    string title = sourceStep.title;
                    string expectedResult = sourceStep.expectedResult;

                    if (m_areRichSteps)
                    {
                        // This is temporary. Work around for product issue.
                        title = title.Replace("\r\n", "<P>").Replace("\n", "<P>").Replace("\r", "<P>");
                        expectedResult = expectedResult.Replace("\r\n", "<P>").Replace("\n", "<P>").Replace("\r", "<P>");
                    }

                    // Creating TCM Test Step and filling testcase with them
                    ITestStep step = m_testCase.CreateTestStep();
                    step.Title = title;

                    // Set the TestStepType properly for validated steps
                    if (!String.IsNullOrEmpty(expectedResult))
                    {
                        step.ExpectedResult = expectedResult;
                        step.TestStepType = TestStepType.ValidateStep;
                    }
                    
                    if (sourceStep.attachments != null)
                    {
                        foreach (string filePath in sourceStep.attachments)
                        {
                            ITestAttachment attachment = step.CreateAttachment(filePath);
                            step.Attachments.Add(attachment);
                        }
                    }
                    m_testCase.Actions.Add(step);

                }
            }
            // else if it is a normal tfs field then just updates the tfs' testcase field
            else
            {
                base.UpdateField(fieldName, value);
            }
        }

        /// <summary>
        /// Get corresponding Field value for field name
        /// </summary>
        /// <param name="fieldName"></param>
        /// <returns></returns>
        public override object GetFieldValue(string fieldName)
        {
            if (String.CompareOrdinal(fieldName, m_stepsFieldName) == 0)
            {
                List<SourceTestStep> sourceSteps = new List<SourceTestStep>();
                foreach (ITestAction action in m_testCase.Actions)
                {
                    ITestStep step = action as ITestStep;
                    if (step != null)
                    {
                        SourceTestStep sourceStep = new SourceTestStep();
                        sourceStep.title = step.Title;
                        sourceStep.expectedResult = step.ExpectedResult;
                        sourceStep.attachments = new List<string>();

                        foreach (var attachment in step.Attachments)
                        {
                            sourceStep.attachments.Add(attachment.Name);
                        }
                        sourceSteps.Add(sourceStep);
                    }
                }
                return sourceSteps;
            }
            else
            {
                return base.GetFieldValue(fieldName);
            }
        }

        /// <summary>
        /// Saves TCM tetcase
        /// </summary>
        public override void Save()
        {
            base.Save();
            m_testCase.Save();
        }
        #endregion
    }
}
