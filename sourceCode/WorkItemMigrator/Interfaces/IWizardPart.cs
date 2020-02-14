//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System.ComponentModel;

    /// <summary>
    /// Interface for Wizard Part. It make sure that implementing Wizard Part must support all of the given functionalities
    /// and properties.
    /// </summary>
    internal interface IWizardPart : INotifyPropertyChanged
    {
        /// <summary>
        /// Initializes the Wizard part with Wizard Info
        /// </summary>
        /// <param name="state"></param>
        void Initialize(WizardInfo state);

        void Clear();

        /// <summary>
        /// Resets the Wizard Part to its initial state
        /// </summary>
        void Reset();

        bool ValidatePartState();

        /// <summary>
        /// Update the current state of Wizard Info with the properties of the wizard part
        /// </summary>
        /// <returns></returns>
        bool UpdateWizardPart();

        /// <summary>
        /// This is the Error Arguments which gets set whenever there is any error during Initialization/Updation of WizardPart
        /// </summary>
        string Warning { get; set; }


        /// <summary>
        /// The Corresponding Unique Name of the current Wizard part. Used to distinguish one Wizard part from other.
        /// </summary>
        WizardPage WizardPage { get; }



        /// <summary>
        /// The Header of WizardPart. Used By View to display the name of Wizard Part.
        /// </summary>
        string Header { get; }


        /// <summary>
        /// The short description which tells about the Wizard Part.
        /// </summary>
        string Description { get; }


        /// <summary>
        /// Can Control go to the previous Wizard Part in sequence.
        /// </summary>
        bool CanBack { get; }


        /// <summary>
        /// Can Control go the next wizard Part in Sequence.
        /// </summary>
        bool CanNext { get; set; }


        /// <summary>
        /// Returns true if all prerequisites are met to intialize the wizard page.
        /// </summary>
        bool CanInitialize { get; }


        /// <summary>
        /// Can Control show this Wizard Part and make it active WizardPart.
        /// </summary>
        bool CanShow { get; set; }


        /// <summary>
        /// Is this wizardPart Active and currently having the control.
        /// </summary>
        bool IsActiveWizardPart { get; set; }
    }
}
