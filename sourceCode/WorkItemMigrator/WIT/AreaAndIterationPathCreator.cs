//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.Diagnostics;
    using System.Security;
    using Microsoft.TeamFoundation;
    using Microsoft.TeamFoundation.Client;
    using Microsoft.TeamFoundation.Core.WebApi;
    using Microsoft.TeamFoundation.Server;
    using Microsoft.TeamFoundation.WorkItemTracking.Client;

    internal class AreaAndIterationPathCreator
    {
        #region Fields

        private WorkItemStore m_store;
        private Project m_project;
        private ICommonStructureService m_css;
        private string m_areaNodeUri;
        private string m_iterationNodeUri;
        private bool m_hasRootNodes;

        private static readonly char[] InvalidChars = new char[] { '\\', '/', '$', '?', '*', ':', '"', '&', '>', '<', '#', '%', '|' };

        #endregion

        #region Constructor
        public AreaAndIterationPathCreator(Project project)
        {
            m_project = project;
            m_store = project.Store;
            m_css = (ICommonStructureService)m_store.TeamProjectCollection.GetService(typeof(ICommonStructureService));
            GetRootNodes();
        }

        #endregion

        #region Public methods
        public string Create(
           Node.TreeType type,
           string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return m_project.Name;
            }

            string[] names = path.Split('\\');
            SanitizeNames(names);
            NodeCollection nc = type == Node.TreeType.Area ? m_project.AreaRootNodes : m_project.IterationRootNodes;
            Node n = null;
            int i = 0;
            if (string.CompareOrdinal(m_project.Name, names[0]) == 0)
            {
                i++;
            }
            for (; i < names.Length; i++)
            {
                string name = names[i];
                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }
                try
                {
                    n = nc[name];
                    nc = n.ChildNodes;
                    continue;
                }
                catch (DeniedOrNotExistException)
                {
                    // Ignore this exception as we need to create the node if it does not exist.
                }

                string parentUri;
                if (n == null)
                {
                    parentUri = type == Node.TreeType.Area ? m_areaNodeUri : m_iterationNodeUri;
                }
                else
                {
                    parentUri = n.Uri.ToString();
                }

                try
                {
                    path = CreatePath(type, parentUri, names, i);
                }
                catch (SecurityException se)
                {
                    throw new WorkItemMigratorException(se.Message + "\n\nNode::" + names[i] + "in " + path, null, null);
                }

                Debug.WriteLine(
                    "Created path '{0}' in the TFS Work Item store '({1})'",
                    m_project.Name + "\\" + path,
                    m_project.Name);

                return path;
            }

            return n.Path;
        }

        #endregion

        #region private methods

        private void SanitizeNames(string[] names)
        {
            for (int i = 0; i < names.Length; i++)
            {
                if (names[i].Length > 255)
                {
                    // Limit the area node name to 255 characters only.
                    names[i] = names[i].Substring(0, 255);
                }

                foreach (char invalidChar in InvalidChars)
                {
                    names[i] = names[i].Replace(invalidChar, '_');
                }

                foreach (char invalidChar in System.IO.Path.GetInvalidPathChars())
                {
                    names[i] = names[i].Replace(invalidChar, '_');
                }
            }
        }

        /// <summary>
        /// Creates path.
        /// </summary>
        /// <param name="type">Type of the node to be created</param>
        /// <param name="parentUri">Parent node</param>
        /// <param name="nodes">Node names</param>
        /// <param name="first">Index of the first node to create</param>
        /// <returns>Id of the node</returns>
        private string CreatePath(
            Node.TreeType type,
            string parentUri,
            string[] nodes,
            int first)
        {
            Debug.Assert(first < nodes.Length, "Nothing to create!");

            // Step 1: create in CSS
            for (int i = first; i < nodes.Length; i++)
            {
                string node = nodes[i];
                if (!string.IsNullOrEmpty(node))
                {
                    try
                    {
                        parentUri = m_css.CreateNode(node, parentUri);
                    }
                    catch (CommonStructureSubsystemException cssEx)
                    {
                        if (cssEx.Message.Contains("TF200020"))
                        {
                            // TF200020 may be thrown if the tree node metadata has been propagated
                            // from css to WIT cache. In this case, we will wait for the node id
                            //   Microsoft.TeamFoundation.Server.CommonStructureSubsystemException: 
                            //   TF200020: The parent node already has a child node with the following name: {0}. 
                            //   Child nodes must have unique names.
                            Node existingNode = WaitForTreeNodeId(type, nodes, i);
                            if (existingNode == null)
                            {
                                throw;
                            }
                            else
                            {
                                parentUri = existingNode.Uri.AbsoluteUri;
                            }
                        }
                    }
                }
            }

            // Step 2: locate in the cache
            // Syncing nodes into WIT database is an asynchronous process, and there's no way to tell
            // the exact moment. 
            Node newNode = WaitForTreeNodeId(type, nodes, -1);
            if (newNode == null)
            {
                return m_project.Name;
            }
            else
            {
                return newNode.Path;
            }
        }


        private Node WaitForTreeNodeId(
            Node.TreeType type,
            string[] nodes,
            int first)
        {
            int[] TIMEOUTS = { 100, 500, 1000, 5000 };
            int[] RetryTimes = { 1, 2, 70, 36 };

            for (int i = 0; i < TIMEOUTS.Length; ++i)
            {
                for (int k = 0; k < RetryTimes[i]; ++k)
                {
                    System.Threading.Thread.Sleep(TIMEOUTS[i]);
                    Debug.Write(string.Format("Wake up from {0} millisec sleep for polling CSS node Id", TIMEOUTS[i]));

                    m_store.RefreshCache();
                    Project p = m_store.Projects[m_project.Name];
                    NodeCollection nc = type == Node.TreeType.Area ? p.AreaRootNodes : p.IterationRootNodes;
                    Node n = null;
                    int numNodesToCheck = nodes.Length - 1;
                    if (first != -1)
                    {
                        numNodesToCheck = first;
                    }

                    try
                    {
                        for (int j = 0; j <= numNodesToCheck; j++)
                        {
                            string name = nodes[j];
                            if (!string.IsNullOrEmpty(name))
                            {
                                n = nc[name];
                                nc = n.ChildNodes;
                            }
                        }

                        return n;
                    }
                    catch (DeniedOrNotExistException)
                    {
                        // The node is not there yet. Try one more time...
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Initializes URI of root area and iteration nodes.
        /// </summary>
        private void GetRootNodes()
        {
            if (!m_hasRootNodes)
            {
                NodeInfo[] nodes = m_css.ListStructures(m_project.Uri.ToString());
                string areaUri = null;
                string iterationUri = null;
                for (int i = 0; i < nodes.Length; i++)
                {
                    NodeInfo n = nodes[i];
                    if (TFStringComparer.CssStructureType.Equals(n.StructureType, "ProjectLifecycle"))
                    {
                        iterationUri = n.Uri;
                    }
                    else if (TFStringComparer.CssStructureType.Equals(n.StructureType, "ProjectModelHierarchy"))
                    {
                        areaUri = n.Uri;
                    }
                }

                m_areaNodeUri = areaUri;
                m_iterationNodeUri = iterationUri;
                m_hasRootNodes = true;
            }
        }

        #endregion
    }
}
