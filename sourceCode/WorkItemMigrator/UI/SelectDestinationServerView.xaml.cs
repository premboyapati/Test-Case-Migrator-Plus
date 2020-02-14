//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.IO;
    using System.Windows;
    using System.Windows.Controls;
    using Microsoft.TeamFoundation.Client;
    using WizardResource = Microsoft.VisualStudio.TestTools.WorkItemMigrator.Resources;

    /// <summary>
    /// Interaction logic for SelectDestinationServer.xaml
    /// </summary>
    internal partial class SelectDestinationServerView : ContentControl
    {
        #region Fields

        SelectDestinationServerPart m_part;

        #endregion

        #region Constructor

        public SelectDestinationServerView(SelectDestinationServerPart part)
        {
            InitializeComponent();
            m_part = part;
            DataContext = part;
        }

        #endregion

        #region Private methods

        private void ServerButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {                
                TeamProjectPicker workitemPicker = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false, new UICredentialsProvider());
                workitemPicker.AcceptButtonText = WizardResource.Select;
                workitemPicker.Text = WizardResource.SelectServerProject;
                workitemPicker.SetDefaultSelectionProvider(DefaultServerSelectionPicker.Instance);
                workitemPicker.ShowDialog();
                if (workitemPicker.SelectedProjects != null || workitemPicker.SelectedProjects.Length > 0)
                {
                    m_part.Server = workitemPicker.SelectedTeamProjectCollection.Uri.ToString();
                    m_part.Project = workitemPicker.SelectedProjects[0].Name;

                    if (DefaultServerSelectionPicker.Instance.GetDefaultServerUri() == null)
                    {
                        DefaultServerSelectionPicker.Instance.SetDefaultServerUri(workitemPicker.SelectedTeamProjectCollection.Uri);
                    }

                    DefaultServerSelectionPicker.Instance.Save(m_part.Server, workitemPicker.SelectedProjects[0].Name);                    
                }
            }
            catch (System.NullReferenceException)
            { }
            catch (ArgumentOutOfRangeException)
            { }
            catch (IOException ex)
            {
                m_part.Warning = "Unable to load Server/Project." + ex.Message;
            }
        }

        #endregion
    }
}
