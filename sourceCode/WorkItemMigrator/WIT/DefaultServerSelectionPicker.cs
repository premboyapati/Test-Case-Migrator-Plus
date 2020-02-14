//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using Microsoft.TeamFoundation.Client;

    /// <summary>
    /// Class that loads and saves last Server/Project Collection settings used by Team Project Picker
    /// </summary>
    internal class DefaultServerSelectionPicker : ITeamProjectPickerDefaultSelectionProvider
    {
        #region Fields

        // Singleton instance
        private static DefaultServerSelectionPicker m_instance;

        // Server URI
        private Uri m_serverUri;

        // List of Default Projects to be selected. In our case only one defualt project is required
        private List<string> m_defaultProjects;

        // The guid of Default Project Collection
        private Guid? m_defaultCollectionId;

        #endregion

        #region Properties

        /// <summary>
        /// Singleton instance of the Class
        /// </summary>
        public static DefaultServerSelectionPicker Instance
        {
            get
            {
                // Initiate a new instance if not already created
                if (m_instance == null)
                {
                    m_instance = new DefaultServerSelectionPicker();
                    m_instance.Load();
                }
                return m_instance;
            }
        }

        #endregion

        #region public Methods

        /// <summary>
        /// Server URI
        /// </summary>
        /// <returns></returns>
        public Uri GetDefaultServerUri()
        {
            return m_serverUri;
        }

        /// <summary>
        /// Sets Server URI
        /// </summary>
        /// <returns></returns>
        public void SetDefaultServerUri(Uri serverUri)
        {
            m_serverUri = serverUri;
        }

        /// <summary>
        /// Default Project Collection's Guid
        /// </summary>
        /// <param name="instanceUri"></param>
        /// <returns></returns>
        public Guid? GetDefaultCollectionId(Uri instanceUri)
        {
            m_serverUri = instanceUri;
            return m_defaultCollectionId;
        }

        /// <summary>
        /// Returns the Default Project to be selected
        /// </summary>
        /// <param name="collectionId"></param>
        /// <returns></returns>
        public IEnumerable<string> GetDefaultProjects(Guid collectionId)
        {
            m_defaultCollectionId = collectionId;
            return m_defaultProjects;
        }

        /// <summary>
        ///  Saves the Server/Project settings to the config file
        /// </summary>
        /// <param name="workitemPicker"></param>
        public void Save(string tfsCollectionURL, string defaultProject)
        {
            m_defaultProjects = new List<string>();
            m_defaultProjects.Add(defaultProject);

            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            config.AppSettings.Settings.Remove("ServerUri");
            config.AppSettings.Settings.Remove("CollectionURL");
            config.AppSettings.Settings.Remove("DefaultCollectionId");
            config.AppSettings.Settings.Remove("DefaultProject");

            config.AppSettings.Settings.Add("ServerUri", m_serverUri.ToString());
            config.AppSettings.Settings.Add("CollectionURL", tfsCollectionURL);
            config.AppSettings.Settings.Add("DefaultCollectionId", m_defaultCollectionId.ToString());
            config.AppSettings.Settings.Add("DefaultProject", defaultProject);

            config.Save(ConfigurationSaveMode.Full);
            ConfigurationManager.RefreshSection("appSettings");
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Loads Server/ Project Settings from the config file
        /// </summary>
        private void Load()
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            if (config.AppSettings.Settings.Count == 0)
            {
                return;
            }
            if (!string.IsNullOrEmpty(config.AppSettings.Settings["ServerUri"].Value))
            {
                m_serverUri = new Uri(config.AppSettings.Settings["ServerUri"].Value);
            }

            Guid guid;
            if (Guid.TryParse(config.AppSettings.Settings["DefaultCollectionId"].Value, out guid))
            {
                m_defaultCollectionId = guid;
            }

            if (!string.IsNullOrEmpty(config.AppSettings.Settings["DefaultProject"].Value))
            {
                m_defaultProjects = new List<string>();
                m_defaultProjects.Add(config.AppSettings.Settings["DefaultProject"].Value);
            }
        }

        #endregion

    }
}
