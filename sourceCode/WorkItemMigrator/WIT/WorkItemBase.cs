//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.TeamFoundation;
    using Microsoft.TeamFoundation.WorkItemTracking.Client;

    internal class WorkItemBase : IWorkItem
    {
        #region Fields

        protected WorkItem m_workItem;
        private Project m_project;
        private string m_workItemTypeName;
        private string m_historyFieldValue;


        #endregion

        #region Constructor

        protected WorkItemBase()
        { }

        public WorkItemBase(Project project, string workItemTypeName)
        {
            m_project = project;
            m_workItemTypeName = workItemTypeName;
        }
        #endregion

        #region Properties

        public virtual object WorkItem
        {
            get
            {
                return m_workItem;
            }
        }

        #endregion

        #region Public Methods

        public virtual void Create()
        {
            m_workItem = m_project.WorkItemTypes[m_workItemTypeName].NewWorkItem();
        }

        public virtual void UpdateField(string fieldName, object value)
        {
            // If value is null then just take the default valuea nd return the default value
            if (value == null)
            {
                return;
            }

            if (String.CompareOrdinal(fieldName, "History") == 0)
            {
                m_historyFieldValue += value.ToString();
            }
            else if (m_workItem.Fields[fieldName].FieldDefinition.SystemType == typeof(DateTime))
            {
                DateTime date;
                if (DateTime.TryParse(value.ToString(), out date))
                {
                    value = date;
                }
            }
            else
            {
                // Sets Value to Work Item's field
                m_workItem.Fields[fieldName].Value = value;
            }
        }

        public virtual object GetFieldValue(string fieldName)
        {
            if (String.CompareOrdinal(fieldName, "History") == 0)
            {
                if (m_workItem.Revisions.Count > 1)
                {
                    int revisonCount = m_workItem.Revisions.Count - 2;
                    return m_workItem.Revisions[revisonCount].Fields[fieldName].Value;
                }
                else
                {
                    return m_historyFieldValue;
                }
            }
            else
            {
                return m_workItem.Fields[fieldName].Value;
            }
        }

        public virtual void Save()
        {
            try
            {
                if (!string.IsNullOrEmpty(m_historyFieldValue))
                {
                    if (m_historyFieldValue.Contains("<div><img src='"))
                    {
                        HistoryBuilder history = CreateHistoryBuilder();
                        m_workItem.Save();
                        m_workItem.Fields["History"].Value = history.Text;
                        m_workItem.Save();

                    }
                    else
                    {
                        m_workItem.Fields["History"].Value = m_historyFieldValue;
                        m_workItem.Save();
                    }
                }
                else
                {
                    m_workItem.Save();
                }
            }
            catch (TeamFoundationServerException)
            {
                string error = string.Empty;
                foreach (Field field in m_workItem.Fields)
                {
                    if (field.Status != FieldStatus.Valid)
                    {
                        error += "TFS Field: " + field.Name + "-->" + field.Status + "\n";
                    }
                }
                throw new WorkItemMigratorException(error, null, null);
            }
        }
        #endregion

        #region Private Methods

        private HistoryBuilder CreateHistoryBuilder()
        {
            HistoryBuilder history = new HistoryBuilder();

            while (m_historyFieldValue.Contains("<div><img src='"))
            {
                history.Append<string>(m_historyFieldValue.Substring(0, m_historyFieldValue.IndexOf("<div><img src='", StringComparison.Ordinal) + 15));

                m_historyFieldValue = m_historyFieldValue.Substring(m_historyFieldValue.IndexOf("<div><img src='", StringComparison.Ordinal) + 15);

                string attachmentPath = m_historyFieldValue.Substring(0, m_historyFieldValue.IndexOf("' /></div>", StringComparison.Ordinal));
                Attachment attachment = new Attachment(attachmentPath);
                m_workItem.Attachments.Add(attachment);
                history.Append<Attachment>(attachment);

                m_historyFieldValue = m_historyFieldValue.Substring(m_historyFieldValue.IndexOf("' /></div>", StringComparison.Ordinal));
            }
            history.Append<string>(m_historyFieldValue);
            return history;
        }

        #endregion
    }

    internal class HistoryBuilder
    {
        #region Fields

        private List<object> m_history = new List<object>();

        #endregion

        #region Properties

        public string Text
        {
            get
            {
                StringBuilder builder = new StringBuilder();
                foreach (object o in m_history)
                {
                    Attachment attachment = o as Attachment;
                    if (attachment != null)
                    {
                        builder.Append(attachment.Uri.ToString());
                    }
                    else
                    {
                        builder.Append(o.ToString());
                    }
                }
                return builder.ToString();
            }
        }

        #endregion

        #region public methods

        public void Append<Type>(Type part)
        {
            m_history.Add(part);
        }

        #endregion
    }
}
