//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.IO;
    using System.Reflection;
    using System.Security;
    using System.Xml;

    internal static class Utilities
    {
        ///<summary>
        ///Utility Function to append an attribute in a XML Node
        ///</summary>
        ///<param name="document">XMLDocument that contains the Node</param>
        ///<param name="node">XML Node at which attribute is going to be append</param>
        /// <param name="attributeName">The anme of the Attribute</param>
        /// <param name="value">Value of The attribute</param>
        public static void AppendAttribute(XmlNode node, string attributeName, string value)
        {
            if (value == null)
            {
                value = string.Empty;
            }
            XmlDocument document = node.OwnerDocument;
            value = SecurityElement.Escape(value);
            var attribute = document.CreateAttribute(attributeName);
            attribute.InnerXml = value;
            node.Attributes.Append(attribute);
        }


        public static void CopyFileLocatedAtAssemblyPathToDestinationFolder(string fileName, string destinationFilePath)
        {
            try
            {
                string directory = Path.GetDirectoryName(destinationFilePath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                Assembly assembly = Assembly.GetExecutingAssembly();
                string fullProcessPath = assembly.Location;
                string processDir = Path.GetDirectoryName(fullProcessPath);
                string sourceFilePath = Path.Combine(processDir, fileName);

                if (File.Exists(sourceFilePath))
                {
                    string destination = Path.Combine(directory, fileName);
                    if (File.Exists(destination))
                    {
                        File.Delete(destination);
                    }
                    File.Copy(sourceFilePath, destination);
                    File.SetAttributes(destination, FileAttributes.Normal);
                }
            }
            catch (IOException)
            { }
        }

        public static void Copy(string sourcePath, string destinationPath)
        {
            try
            {
                if (File.Exists(sourcePath))
                {
                    if (File.Exists(destinationPath))
                    {
                        File.Delete(destinationPath);
                    }
                    File.Copy(sourcePath, destinationPath);
                }
            }
            catch (IOException ioEx)
            {
                throw new WorkItemMigratorException(ioEx.Message, null, null);
            }
        }
    }
}
