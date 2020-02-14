//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Net;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Xml;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.TeamFoundation.WorkItemTracking.Client;
    using Microsoft.TeamFoundation.WorkItemTracking.Proxy;
    /// <summary>
    /// Responsible for managing and creating links between different work item typws and accross sessions
    /// </summary>
    internal class LinksManager : NotifyPropertyChange, ILinksManager
    {
        #region Fields

        // Wizard info
        private WizardInfo m_wizardInfo;

        // List of All links that are either already created or has to create
        private IList<ILink> m_links;

        private int m_currentRow;

        #endregion

        #region Constants

        // Node Names of Links File
        private const string RootXMLElementName = "WorkItemMigratorLinks";
        public const string WorkItemTypeXMLElementName = "WorkItemType";
        public const string NameXMLElementName = "Name";
        public const string PassedLinksXMLElementName = "PassedLinks";
        public const string WarningLinksXMLElementName = "WarningLinks";
        public const string FailedLinksXMLElementName = "FailedLinks";
        public const string LinkIDGroupXMLElementName = "LinkIDGroup";
        public const string IDXMLElementName = "ID";
        public const string LinkXMLElementName = "Link";
        public const string LinksXMLElementName = "Links";
        public const string StartWorkItemTypeXMLElementName = "StartWorkItemType";
        public const string StartWorkItemCategoryElementName = "StartWorkItemCategory";
        public const string EndWorkItemCategoryElementName = "EndWorkItemCategory";
        public const string StartWorkItemSourceIDXMLElementName = "StartWorkItemSourceID";
        public const string StartWorkItemTFSIdXMLElementName = "StartWorkItemTfsId";
        public const string EndWorkItemTfsTypeXMLElementName = "EndWorkItemTfsType";
        public const string EndWorkItemSourceIdXMLElementName = "EndWorkItemSourceId";
        public const string EndWorkItemTFSIdXMLElementName = "EndWorkItemTfsId";
        public const string MessageXMLElementName = "Message";
        public const string StatusXMLElementName = "Status";
        public const string IsExistsInTFSXMLElementName = "IsExistsInTfs";
        public const string LinkTypeXMLElementName = "LinkType";
        public const string IDMappingXMLElementName = "IDMapping";
        public const string SourceIDXMLElementName = "SourceID";
        public const string TFSIDXMLElementName = "TFSID";
        public const string SessionIdElementName = "SessionID";

        public const string XLReportFileName = "LinksMapping-ReportFile.xls";
        #endregion

        #region Constructor

        /// <summary>
        /// Constructor to initailize links and IDMapping
        /// </summary>
        /// <param name="m_wizardInfo"></param>
        public LinksManager(WizardInfo wizardInfo)
        {
            m_wizardInfo = wizardInfo;
            WorkItemCategoryToIdMappings = new Dictionary<string, IDictionary<string, WorkItemMigrationStatus>>();
            m_links = new List<ILink>();
            SessionId = 1;

            try
            {
                // Try to load links
                LoadLinksAndIDMapping();
            }
            // Links file may be corrupted in that case ignore the links file
            catch (XmlException)
            { }
        }

        #endregion

        #region Properties

        public int SessionId
        {
            get;
            private set;
        }

        /// <summary>
        /// Path of File containing information about ID Mapping and links 
        /// </summary>
        public string LinksFilePath
        {
            get;
            set;
        }

        // Mapping from Work Item type Name to IDMapping Dictionary where
        // IDmapping Dictionary has 1-1 mapping from source id of workitem to TFS ID
        public IDictionary<string, IDictionary<string, WorkItemMigrationStatus>> WorkItemCategoryToIdMappings
        {
            get;
            private set;
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Updates ID Mapping Dictionary with one new entry
        /// </summary>
        /// <param name="sourceId"></param>
        /// <param name="id"></param>
        public void UpdateIdMapping(string sourceId, WorkItemMigrationStatus status)
        {
            string category = m_wizardInfo.WorkItemGenerator.WorkItemTypeToCategoryMapping[m_wizardInfo.WorkItemGenerator.SelectedWorkItemTypeName];
            if (!WorkItemCategoryToIdMappings.ContainsKey(category))
            {
                WorkItemCategoryToIdMappings.Add(category, new Dictionary<string, WorkItemMigrationStatus>());
            }
            if (!WorkItemCategoryToIdMappings[category].ContainsKey(sourceId))
            {
                WorkItemCategoryToIdMappings[category].Add(sourceId, status);
            }
        }

        /// <summary>
        /// Adds new link into the list of links
        /// </summary>
        /// <param name="linkToAdd"></param>
        public void AddLink(ILink linkToAdd)
        {
            // If Invalid link(not having source id of end/start workitem) then returns
            if (string.IsNullOrEmpty(linkToAdd.StartWorkItemSourceId) ||
                string.IsNullOrEmpty(linkToAdd.EndWorkItemSourceId))
            {
                return;
            }

            // Child link creation is not possible so if link is of type 'child' then change it appropraitely to 'Parent' type
            if (String.CompareOrdinal(linkToAdd.LinkTypeName, "Child") == 0)
            {
                string id = linkToAdd.StartWorkItemSourceId;
                string type = linkToAdd.StartWorkItemTfsTypeName;
                linkToAdd.StartWorkItemSourceId = linkToAdd.EndWorkItemSourceId;
                linkToAdd.StartWorkItemTfsTypeName = linkToAdd.EndWorkItemTfsTypeName;
                linkToAdd.EndWorkItemSourceId = id;
                linkToAdd.EndWorkItemTfsTypeName = type;
                linkToAdd.LinkTypeName = "Parent";
            }

            // Check that if the entry of link is already added then register this link for retrying
            foreach (Link link in m_links)
            {
                if (String.CompareOrdinal(link.StartWorkItemCategory, linkToAdd.StartWorkItemCategory) == 0 &&
                    String.CompareOrdinal(link.StartWorkItemSourceId, linkToAdd.StartWorkItemSourceId) == 0 &&
                    String.CompareOrdinal(link.LinkTypeName, linkToAdd.LinkTypeName) == 0 &&
                    String.CompareOrdinal(link.EndWorkItemCategory, linkToAdd.EndWorkItemCategory) == 0 &&
                    String.CompareOrdinal(link.EndWorkItemSourceId, linkToAdd.EndWorkItemSourceId) == 0)
                {
                    return;
                }
            }

            // Reset the TFS Status of link
            linkToAdd.StartWorkItemTfsId = -1;
            linkToAdd.EndWorkItemTfsId = -1;
            linkToAdd.IsExistInTfs = false;
            linkToAdd.Message = string.Empty;

            // Add the link in the list of links
            m_links.Add(linkToAdd);
        }

        /// <summary>
        /// Creates all remaining links and save the status in the links file
        /// </summary>
        public void Save()
        {
            // list of links which are going to be created
            List<ILink> linksToCreate = new List<ILink>();

            // Iterate through all links
            foreach (var link in m_links)
            {
                // Process the link only if it does not exist in the TFS
                if (!link.IsExistInTfs)
                {
                    FillWorkItemTypesOfLink(link);

                    // Update TFS ID of Start Workitem of link
                    string startCategory = link.StartWorkItemCategory;
                    if (WorkItemCategoryToIdMappings.ContainsKey(startCategory) &&
                        WorkItemCategoryToIdMappings[startCategory].ContainsKey(link.StartWorkItemSourceId) &&
                        WorkItemCategoryToIdMappings[startCategory][link.StartWorkItemSourceId].TfsId != -1)
                    {
                        link.StartWorkItemTfsId = WorkItemCategoryToIdMappings[startCategory][link.StartWorkItemSourceId].TfsId;
                    }

                    // Update TFS ID of End WorkItem
                    string endCategory = link.EndWorkItemCategory;
                    if (WorkItemCategoryToIdMappings.ContainsKey(endCategory) &&
                        WorkItemCategoryToIdMappings[endCategory].ContainsKey(link.EndWorkItemSourceId) &&
                        WorkItemCategoryToIdMappings[endCategory][link.EndWorkItemSourceId].TfsId != -1)
                    {
                        link.EndWorkItemTfsId = WorkItemCategoryToIdMappings[endCategory][link.EndWorkItemSourceId].TfsId;
                    }

                    // process the link only if both workitem ends of the link are migrated
                    if (link.StartWorkItemTfsId > 0 && link.EndWorkItemTfsId > 0)
                    {
                        linksToCreate.Add(link);
                    }
                }
                else
                {
                    link.Message = null;
                }
            }

            // Command WorkItemGenertor to create all links
            m_wizardInfo.WorkItemGenerator.CreateLinksInBatch(linksToCreate);

            foreach (ILink link in linksToCreate)
            {
                if (link.IsExistInTfs)
                {
                    link.SessionId = SessionId;
                    link.Message = null;
                }
            }

            // Folder path containg links file
            string linksFileFolderPath = Path.GetDirectoryName(m_wizardInfo.Reporter.ReportFile);

            // Get the links file path
            string linksFile = GetLinksFileName();
            LinksFilePath = Path.Combine(linksFileFolderPath, linksFile);

            // Save links File
            SaveLinksFile();

            SaveLinksFileAtTask(linksFile);
        }

        public void PublishReport()
        {
            string tempPath = Path.GetTempPath();
            string tempFile = Path.Combine(tempPath, XLReportFileName);
            Application app = null;
            try
            {

                Utilities.CopyFileLocatedAtAssemblyPathToDestinationFolder(XLReportFileName, tempPath);

                app = new Application();
                Workbook workBook = app.Workbooks.Open(tempFile);
                Worksheet workSheet = workBook.Worksheets[1];

                m_currentRow = 2;

                FillIDMappingsInReport(workSheet);

                FillLinksRawDataInReport(workSheet);

                foreach (Worksheet sheet in workBook.Worksheets)
                {
                    foreach (PivotTable pivot in sheet.PivotTables(Type.Missing))
                    {
                        pivot.RefreshTable();
                    }
                }

                workBook.Save();
                workBook.Close();
                workBook = null;
                app.Workbooks.Close();
                app.Quit();
                app = null;
            }
            catch (COMException comEx)
            {
                throw new WorkItemMigratorException(comEx.Message, null, null);
            }
            catch (InvalidCastException inEx)
            {
                throw new WorkItemMigratorException(inEx.Message, null, null);
            }
            catch (IOException ioEx)
            {
                throw new WorkItemMigratorException(ioEx.Message, null, null);
            }
            finally
            {
                if (app != null)
                {
                    try
                    {
                        app.Quit();
                        app = null;
                    }
                    catch (COMException)
                    { }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                string linksFileFolderPath = Path.GetDirectoryName(LinksFilePath);
                string destinationReportFilePath = Path.Combine(linksFileFolderPath, XLReportFileName);
                File.Move(tempFile, destinationReportFilePath);           
            }
        }

        #endregion

        #region Private Methods
        /// <summary>
        /// Sanitize fileName
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return string.Empty;

            char[] invalidChars = Path.GetInvalidFileNameChars();

            //Replace all occurences of % with %{0:x}
            fileName = ReplaceCharToAscii(fileName, '%');

            StringBuilder fileNameBuilder = new StringBuilder();
            int startIndex = 0;
            int invalidIndex = -1;
            while ((startIndex < fileName.Length) &&
                   (invalidIndex = fileName.IndexOfAny(invalidChars, startIndex)) != -1)
            {
                if (startIndex != invalidIndex)
                {
                    //When the beginning character is not an invalid character
                    fileNameBuilder.Append(fileName.Substring(startIndex, invalidIndex - startIndex));
                }
                fileNameBuilder.Append(GetInvalidCharacterFormat(fileName[invalidIndex]));

                startIndex = invalidIndex + 1;
            }

            if (startIndex < fileName.Length)
            {
                //Append if any characters exits
                fileNameBuilder.Append(fileName.Substring(startIndex, fileName.Length - startIndex));
            }

            return fileNameBuilder.ToString();
        }

        /// <summary>
        /// Replaces char with Ascii Value
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private static string ReplaceCharToAscii(String name, char c)
        {

            string replacement = GetInvalidCharacterFormat(c);
            string org = new String(c, 1);
            return name.Replace(org, replacement);
        }

        /// <summary>
        /// constructs the format for invaid character
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        private static string GetInvalidCharacterFormat(char c)
        {
            return String.Format(CultureInfo.CurrentCulture, "%{0:X}", Convert.ToInt32(c));
        }
        /// <summary>
        /// Get the  input links file.
        /// </summary>
        /// <returns>The latest links file that exists.</returns>
        private string GetLinksFileName()
        {
            string serverName = m_wizardInfo.WorkItemGenerator.Server;
            string projectName = m_wizardInfo.WorkItemGenerator.Project;
            string fileName = string.Format(CultureInfo.InvariantCulture, "LinksFile{0}{1}.xml", serverName, projectName);
            return SanitizeFileName(fileName).ToLowerInvariant();
        }

        /// <summary>
        /// Load Links and ID Mapping from Links file
        /// </summary>
        private void LoadLinksAndIDMapping()
        {
            LoadLinkingFileFromTask();

            // First time there will not be any links file
            if (!File.Exists(LinksFilePath))
            {
                return;
            }

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(LinksFilePath);

            // get the root Node of the xml Document
            XmlElement root = (XmlElement)xmlDocument.LastChild;

            var deletedWITs = new List<WorkItemMigrationStatus>();

            // Traverse each child node of the root node
            foreach (XmlElement node in root.ChildNodes)
            {
                switch (node.Name)
                {
                    case SessionIdElementName:
                        SessionId = int.Parse(node.InnerXml);
                        if (SessionId > 0)
                        {
                            SessionId++;
                        }
                        else
                        {
                            SessionId = 1;
                        }
                        break;

                    // If this is ID Mapping Node
                    case IDMappingXMLElementName:
                        foreach (XmlElement workItemTypeNode in node.ChildNodes)
                        {
                            string workItemType = workItemTypeNode.Attributes[NameXMLElementName].Value;
                            string workItemCategory = m_wizardInfo.WorkItemGenerator.WorkItemTypeToCategoryMapping[workItemType];
                            WorkItemCategoryToIdMappings.Add(workItemCategory, new Dictionary<string, WorkItemMigrationStatus>());
                            foreach (XmlElement idMappingNode in workItemTypeNode.ChildNodes)
                            {
                                string sourceId = idMappingNode.Attributes[SourceIDXMLElementName].Value;

                                WorkItemMigrationStatus status = new WorkItemMigrationStatus();
                                status.WorkItemType = workItemType;
                                status.TfsId = int.Parse(idMappingNode.Attributes[TFSIDXMLElementName].Value);
                                status.SourceId = sourceId;
                                status.Message = idMappingNode.Attributes[MessageXMLElementName].Value;
                                status.Status = (Status)Enum.Parse(typeof(Status), idMappingNode.Attributes[StatusXMLElementName].Value, true);
                                status.SessionId = int.Parse(idMappingNode.Attributes[SessionIdElementName].Value);

                                if (m_wizardInfo.WorkItemGenerator.IsWitExists(status.TfsId))
                                {
                                    WorkItemCategoryToIdMappings[workItemCategory].Add(sourceId, status);
                                }
                                else
                                {
                                    deletedWITs.Add(status);
                                }
                            }
                        }
                        break;

                    case LinksXMLElementName:
                        foreach (XmlElement workItemTypeNode in node.ChildNodes)
                        {
                            string workItemType = workItemTypeNode.Attributes[NameXMLElementName].Value;
                            foreach (XmlElement linkCategoryNode in workItemTypeNode.ChildNodes)
                            {
                                foreach (XmlElement linkIDGroupNode in linkCategoryNode.ChildNodes)
                                {
                                    int id = int.Parse(linkIDGroupNode.Attributes[IDXMLElementName].Value);
                                    foreach (XmlElement linkNode in linkIDGroupNode.ChildNodes)
                                    {
                                        Link link = new Link();
                                        link.StartWorkItemTfsTypeName = linkNode.Attributes[StartWorkItemTypeXMLElementName].Value;
                                        link.StartWorkItemSourceId = linkNode.Attributes[StartWorkItemSourceIDXMLElementName].Value;
                                        link.StartWorkItemTfsId = int.Parse(linkNode.Attributes[StartWorkItemTFSIdXMLElementName].Value);
                                        link.LinkTypeName = linkNode.Attributes[LinkTypeXMLElementName].Value;
                                        link.EndWorkItemTfsTypeName = linkNode.Attributes[EndWorkItemTfsTypeXMLElementName].Value;
                                        link.EndWorkItemSourceId = linkNode.Attributes[EndWorkItemSourceIdXMLElementName].Value;
                                        link.EndWorkItemTfsId = int.Parse(linkNode.Attributes[EndWorkItemTFSIdXMLElementName].Value);
                                        link.IsExistInTfs = bool.Parse(linkNode.Attributes[IsExistsInTFSXMLElementName].Value);
                                        link.Message = linkNode.Attributes[MessageXMLElementName].Value;
                                        link.StartWorkItemCategory = linkNode.Attributes[StartWorkItemCategoryElementName].Value;
                                        link.EndWorkItemCategory = linkNode.Attributes[EndWorkItemCategoryElementName].Value;
                                        int sessionId;
                                        int.TryParse(linkNode.Attributes[SessionIdElementName].Value, out sessionId);
                                        link.SessionId = sessionId;
                                        m_links.Add(link);
                                    }
                                }
                            }
                        }
                        break;

                    default:
                        break;
                }
            }
            RemoveLinksofDeletedWITs(deletedWITs);
        }

        private void RemoveLinksofDeletedWITs(List<WorkItemMigrationStatus> deletedWITs)
        {
            foreach (var deletedWit in deletedWITs)
            {
                if (deletedWit.TfsId > 0)
                {
                    int linkNumber = 0;
                    while (linkNumber < m_links.Count)
                    {
                        ILink link = m_links[linkNumber];
                        if (string.CompareOrdinal(deletedWit.WorkItemType, link.StartWorkItemTfsTypeName) == 0 &&
                            string.CompareOrdinal(deletedWit.SourceId, link.StartWorkItemSourceId) == 0)
                        {
                            m_links.Remove(link);
                            if (link.IsExistInTfs)
                            {
                                m_wizardInfo.WorkItemGenerator.RemoveLink(link.EndWorkItemTfsId, link.StartWorkItemTfsId);
                            }
                            continue;
                        }
                        else if (string.CompareOrdinal(deletedWit.WorkItemType, link.EndWorkItemTfsTypeName) == 0 &&
                            string.CompareOrdinal(deletedWit.SourceId, link.EndWorkItemSourceId) == 0)
                        {
                            if (link.IsExistInTfs)
                            {
                                m_wizardInfo.WorkItemGenerator.RemoveLink(link.StartWorkItemTfsId, link.EndWorkItemTfsId);
                            }
                            link.EndWorkItemTfsId = -1;
                            link.IsExistInTfs = false;
                        }
                        linkNumber++;
                    }
                }
            }
        }
        
        private void LoadLinkingFileFromTask()
        {
            string linksFolder = Path.GetDirectoryName(m_wizardInfo.Reporter.ReportFile);
            string linksFileName = GetLinksFileName();
            LinksFilePath = Path.Combine(linksFolder, linksFileName);

            WorkItem task = m_wizardInfo.WorkItemGenerator.LinkingTask;
            if (task != null)
            {
                foreach (Attachment attachment in task.Attachments)
                {
                    if (String.CompareOrdinal(linksFileName, attachment.Name) == 0)
                    {
                        if (!Directory.Exists(linksFolder))
                        {
                            Directory.CreateDirectory(linksFolder);
                            using (new StreamWriter(LinksFilePath))
                            { }
                        }
                        
                        WorkItemServer wiServer = (WorkItemServer)(m_wizardInfo.WorkItemGenerator.TeamProjectCollection.GetService(typeof(WorkItemServer)));
                        if (wiServer != null)
                        {
                            int attachmentId = GetFileIdFromAttachmentUri(attachment.Uri);
                            if (attachmentId > 0)
                            {
                                string downloadedFilePath = wiServer.DownloadFile(attachmentId);
                                if (!string.IsNullOrEmpty(downloadedFilePath))
                                    File.Move(downloadedFilePath, LinksFilePath);
                            }
                        }                     
                     
                        break;
                    }
                }
            }
        }

        private int GetFileIdFromAttachmentUri(Uri attachmentUri)
        {
            int fileId = -1;
                     
            attachmentUri = new Uri(attachmentUri.AbsoluteUri);
          
            string uriQuery = attachmentUri.GetComponents(UriComponents.Query, UriFormat.UriEscaped);

            if (!string.IsNullOrEmpty(uriQuery))
            {
                string[] queryClauses = uriQuery.Split('&');

                foreach (string clause in queryClauses)
                {
                    if (clause.StartsWith("FileID", true, CultureInfo.InvariantCulture))
                    {
                        string strFileId = clause.Substring(clause.LastIndexOf('=') + 1);

                        if (string.IsNullOrEmpty(strFileId) || (!int.TryParse(strFileId, out fileId)))
                        {
                            fileId = -1;                         
                        }

                        break;
                    }
                }
            }

            return fileId;
        }
        
        private void SaveLinksFile()
        {
            if (string.IsNullOrEmpty(LinksFilePath))
            {
                return;
            }

            var workItemTypeToIDMapping = GetWorkItemTypeIDmapping();

            XmlDocument xmlDocument = new XmlDocument();

            XmlDeclaration dec = xmlDocument.CreateXmlDeclaration("1.0", null, null);
            xmlDocument.AppendChild(dec);

            var rootNode = xmlDocument.CreateElement(RootXMLElementName);
            xmlDocument.AppendChild(rootNode);

            // Saving Session Id Node
            var sessionIdNode = xmlDocument.CreateElement(SessionIdElementName);
            sessionIdNode.InnerXml = SessionId.ToString(CultureInfo.InvariantCulture);
            rootNode.AppendChild(sessionIdNode);

            // Saving IDMappings
            var idMappingNode = xmlDocument.CreateElement(IDMappingXMLElementName);
            foreach (var workItemTypeToIDMappingsKVP in workItemTypeToIDMapping)
            {
                var workItemTypeNode = xmlDocument.CreateElement(WorkItemTypeXMLElementName);
                Utilities.AppendAttribute(workItemTypeNode, NameXMLElementName, workItemTypeToIDMappingsKVP.Key);

                foreach (var idMapping in workItemTypeToIDMappingsKVP.Value)
                {
                    var idNode = xmlDocument.CreateElement(IDXMLElementName);
                    Utilities.AppendAttribute(idNode, SourceIDXMLElementName, idMapping.Key);
                    Utilities.AppendAttribute(idNode, TFSIDXMLElementName, idMapping.Value.TfsId.ToString(CultureInfo.InvariantCulture));
                    Utilities.AppendAttribute(idNode, StatusXMLElementName, idMapping.Value.Status.ToString());
                    Utilities.AppendAttribute(idNode, SessionIdElementName, idMapping.Value.SessionId.ToString());
                    Utilities.AppendAttribute(idNode, MessageXMLElementName, idMapping.Value.Message);
                    Utilities.AppendAttribute(idNode, WorkItemTypeXMLElementName, idMapping.Value.WorkItemType);
                    workItemTypeNode.AppendChild(idNode);
                }
                idMappingNode.AppendChild(workItemTypeNode);
            }
            rootNode.AppendChild(idMappingNode);

            // Saving links

            var workItemTypeToPassedLinks = new Dictionary<string, IDictionary<string, IList<Link>>>();
            var workItemTypeToWarningLinks = new Dictionary<string, IDictionary<string, IList<Link>>>();
            var workItemTypeToFailedLinks = new Dictionary<string, IDictionary<string, IList<Link>>>();
            foreach (Link link in m_links)
            {
                if (link.IsExistInTfs)
                {
                    AddLink(workItemTypeToPassedLinks, link);
                }
                else if (link.EndWorkItemTfsId < 0)
                {
                    AddLink(workItemTypeToWarningLinks, link);
                }
                else
                {
                    AddLink(workItemTypeToFailedLinks, link);
                }
            }

            var linksNode = xmlDocument.CreateElement(LinksXMLElementName);
            foreach (string workItemType in workItemTypeToIDMapping.Keys)
            {
                var workItemNode = xmlDocument.CreateElement(WorkItemTypeXMLElementName);
                Utilities.AppendAttribute(workItemNode, NameXMLElementName, workItemType);
                AppendLinksToWorkItemTypeNode(xmlDocument, workItemNode, PassedLinksXMLElementName, workItemType, workItemTypeToPassedLinks);
                AppendLinksToWorkItemTypeNode(xmlDocument, workItemNode, WarningLinksXMLElementName, workItemType, workItemTypeToWarningLinks);
                AppendLinksToWorkItemTypeNode(xmlDocument, workItemNode, FailedLinksXMLElementName, workItemType, workItemTypeToFailedLinks);
                linksNode.AppendChild(workItemNode);
            }
            rootNode.AppendChild(linksNode);

            string linksFolder = Path.GetDirectoryName(LinksFilePath);
            if (!Directory.Exists(linksFolder))
            {
                Directory.CreateDirectory(linksFolder);
            }
            xmlDocument.Save(LinksFilePath);
        }

        private Dictionary<string, Dictionary<string, WorkItemMigrationStatus>> GetWorkItemTypeIDmapping()
        {
            var workItemTypeToIDMapping = new Dictionary<string, Dictionary<string, WorkItemMigrationStatus>>();
            foreach (var dict in WorkItemCategoryToIdMappings.Values)
            {
                foreach (var kvp in dict)
                {
                    if (!workItemTypeToIDMapping.ContainsKey(kvp.Value.WorkItemType))
                    {
                        workItemTypeToIDMapping.Add(kvp.Value.WorkItemType, new Dictionary<string, WorkItemMigrationStatus>());
                    }
                    workItemTypeToIDMapping[kvp.Value.WorkItemType].Add(kvp.Value.SourceId, kvp.Value);
                }
            }
            return workItemTypeToIDMapping;
        }

        private void AddLink(Dictionary<string, IDictionary<string, IList<Link>>> workItemTypeToLinks, Link link)
        {
            if (!workItemTypeToLinks.ContainsKey(link.StartWorkItemTfsTypeName))
            {
                workItemTypeToLinks.Add(link.StartWorkItemTfsTypeName, new Dictionary<string, IList<Link>>());
            }
            if (!workItemTypeToLinks[link.StartWorkItemTfsTypeName].ContainsKey(link.StartWorkItemSourceId))
            {
                workItemTypeToLinks[link.StartWorkItemTfsTypeName].Add(link.StartWorkItemSourceId, new List<Link>());
            }
            workItemTypeToLinks[link.StartWorkItemTfsTypeName][link.StartWorkItemSourceId].Add(link);
        }

        private void AppendLinksToWorkItemTypeNode(XmlDocument document,
                                                   XmlElement workItemNode,
                                                   string nodeName,
                                                   string workItemTypeName,
                                                   IDictionary<string, IDictionary<string, IList<Link>>> workItemTypeToLinks)
        {
            if (!workItemTypeToLinks.ContainsKey(workItemTypeName))
            {
                return;
            }
            var node = document.CreateElement(nodeName);

            foreach (var kvp in workItemTypeToLinks[workItemTypeName])
            {
                var linkIDGroupNode = document.CreateElement(LinkIDGroupXMLElementName);
                Utilities.AppendAttribute(linkIDGroupNode, IDXMLElementName, kvp.Key);
                foreach (var link in kvp.Value)
                {
                    var linkNode = document.CreateElement(LinkXMLElementName);
                    Utilities.AppendAttribute(linkNode, StartWorkItemTypeXMLElementName, link.StartWorkItemTfsTypeName);
                    Utilities.AppendAttribute(linkNode, StartWorkItemSourceIDXMLElementName, link.StartWorkItemSourceId);
                    Utilities.AppendAttribute(linkNode, StartWorkItemTFSIdXMLElementName, link.StartWorkItemTfsId.ToString(CultureInfo.InvariantCulture));
                    Utilities.AppendAttribute(linkNode, LinkTypeXMLElementName, link.LinkTypeName);
                    Utilities.AppendAttribute(linkNode, EndWorkItemTfsTypeXMLElementName, link.EndWorkItemTfsTypeName);
                    Utilities.AppendAttribute(linkNode, EndWorkItemSourceIdXMLElementName, link.EndWorkItemSourceId);
                    Utilities.AppendAttribute(linkNode, EndWorkItemTFSIdXMLElementName, link.EndWorkItemTfsId.ToString(CultureInfo.InvariantCulture));
                    Utilities.AppendAttribute(linkNode, IsExistsInTFSXMLElementName, link.IsExistInTfs.ToString(CultureInfo.InvariantCulture));
                    Utilities.AppendAttribute(linkNode, SessionIdElementName, link.SessionId.ToString(CultureInfo.InvariantCulture));
                    Utilities.AppendAttribute(linkNode, MessageXMLElementName, link.Message);
                    Utilities.AppendAttribute(linkNode, StartWorkItemCategoryElementName, link.StartWorkItemCategory);
                    Utilities.AppendAttribute(linkNode, EndWorkItemCategoryElementName, link.EndWorkItemCategory);

                    linkIDGroupNode.AppendChild(linkNode);
                }
                node.AppendChild(linkIDGroupNode);
            }
            workItemNode.AppendChild(node);
        }

        private void FillIDMappingsInReport(Worksheet sheet)
        {
            var wiTypeToIdMapping = GetWorkItemTypeIDmapping();
            foreach (var workItemTypeToIDMapping in wiTypeToIdMapping)
            {
                foreach (var idMapping in workItemTypeToIDMapping.Value)
                {
                    ExcelReporter.WriteValueAt(sheet, m_currentRow, 1, m_currentRow - 1);
                    ExcelReporter.WriteValueAt(sheet, m_currentRow, 2, idMapping.Value.SessionId);
                    ExcelReporter.WriteValueAt(sheet, m_currentRow, 3, idMapping.Value.Status);
                    ExcelReporter.WriteValueAt(sheet, m_currentRow, 5, workItemTypeToIDMapping.Key);
                    ExcelReporter.WriteValueAt(sheet, m_currentRow, 6, idMapping.Key);
                    ExcelReporter.WriteValueAt(sheet, m_currentRow, 7, idMapping.Value.TfsId);
                    ExcelReporter.WriteValueAt(sheet, m_currentRow, 13, idMapping.Value.Message);
                    m_currentRow++;
                }
            }
        }

        private void FillLinksRawDataInReport(Worksheet sheet)
        {
            foreach (ILink link in m_links)
            {
                ExcelReporter.WriteValueAt(sheet, m_currentRow, 1, m_currentRow - 1);
                if (link.SessionId > 0)
                {
                    ExcelReporter.WriteValueAt(sheet, m_currentRow, 2, link.SessionId);
                }
                ExcelReporter.WriteValueAt(sheet, m_currentRow, 4, link.LinksStatus);
                ExcelReporter.WriteValueAt(sheet, m_currentRow, 5, link.StartWorkItemTfsTypeName);
                ExcelReporter.WriteValueAt(sheet, m_currentRow, 6, link.StartWorkItemSourceId);
                ExcelReporter.WriteValueAt(sheet, m_currentRow, 7, link.StartWorkItemTfsId);
                ExcelReporter.WriteValueAt(sheet, m_currentRow, 8, link.LinkTypeName);
                ExcelReporter.WriteValueAt(sheet, m_currentRow, 9, link.EndWorkItemTfsTypeName);
                ExcelReporter.WriteValueAt(sheet, m_currentRow, 10, link.EndWorkItemSourceId);
                ExcelReporter.WriteValueAt(sheet, m_currentRow, 11, link.EndWorkItemTfsId);
                ExcelReporter.WriteValueAt(sheet, m_currentRow, 12, link.IsExistInTfs);

                string message = link.Message != null ? link.Message : string.Empty;
                if (!link.IsExistInTfs && string.IsNullOrEmpty(link.Message))
                {
                    message = "Target work item is not migrated yet.";
                }
                ExcelReporter.WriteValueAt(sheet, m_currentRow, 13, message);
                m_currentRow++;
            }
        }

        private void SaveLinksFileAtTask(string linksFile)
        {
            WorkItem task = m_wizardInfo.WorkItemGenerator.LinkingTask;
            if (task == null)
            {
                throw new WorkItemMigratorException("Unable to create task containg links metadata", null, null);
            }

            Attachment oldAttachment = null;
            foreach (Attachment attachment in task.Attachments)
            {
                if (String.CompareOrdinal(linksFile, attachment.Name) == 0)
                {
                    oldAttachment = attachment;
                    break;
                }
            }
            if (oldAttachment != null)
            {
                task.Attachments.Remove(oldAttachment);
            }
            Attachment newAttachment = new Attachment(LinksFilePath);
            task.Attachments.Add(newAttachment);
            task.Save();
        }

        private void FillWorkItemTypesOfLink(ILink link)
        {
            if (WorkItemCategoryToIdMappings.ContainsKey(link.StartWorkItemCategory) &&
                WorkItemCategoryToIdMappings[link.StartWorkItemCategory].ContainsKey(link.StartWorkItemSourceId))
            {
                link.StartWorkItemTfsTypeName = WorkItemCategoryToIdMappings[link.StartWorkItemCategory][link.StartWorkItemSourceId].WorkItemType;
            }
            if (WorkItemCategoryToIdMappings.ContainsKey(link.EndWorkItemCategory) &&
                WorkItemCategoryToIdMappings[link.EndWorkItemCategory].ContainsKey(link.EndWorkItemSourceId))
            {
                link.EndWorkItemTfsTypeName = WorkItemCategoryToIdMappings[link.EndWorkItemCategory][link.EndWorkItemSourceId].WorkItemType;
            }
        }

        #endregion

    }

    public class WorkItemMigrationStatus
    {
        public string SourceId
        {
            get;
            set;
        }

        public string WorkItemType
        {
            get;
            set;
        }

        public int TfsId
        {
            get;
            set;
        }

        public string Message
        {
            get;
            set;
        }

        public int SessionId
        {
            get;
            set;
        }

        public Status Status
        {
            get;
            set;
        }
    }

    public enum Status
    {
        Passed,
        Warning,
        Failed
    }
}
