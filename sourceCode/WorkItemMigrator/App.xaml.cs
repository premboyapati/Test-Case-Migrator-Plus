//  Copyright (c) Microsoft Corporation.  All rights reserved.

namespace Microsoft.VisualStudio.TestTools.WorkItemMigrator
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Threading;
    using System.Windows;

    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    internal partial class App : Application
    {

        #region Native Methods

        /// <summary>
        /// Native method to attach App with Console 
        /// </summary>
        /// <param name="dwProcessId"></param>
        /// <returns></returns>
        [DllImport("kernel32.dll")]
        static extern bool AttachConsole(int dwProcessId);
        private const int ATTACH_PARENT_PROCESS = -1;

        #endregion

        #region Fields

        private static App m_application;
        private int m_processedWorkItemCount;
        private int m_passedWorkItemCount;
        private int m_warningWorkItemCount;
        private int m_failedWorkItemCount;
        private static Dictionary<String, CommandLineSwitch> s_commandSwitchLookup = new Dictionary<string, CommandLineSwitch>(StringComparer.OrdinalIgnoreCase);

        #endregion

        #region Constants

        public const string ExcelSwitch = "excel";
        public const string MHTSwitch = "mht";
        public const string CollectionCLISwitch = "collection";
        public const string ProjectCLISwitch = "project";
        public const string WorkItemTypeCLISwitch = "workitemtype";
        public const string SourceFileCLISwitch = "source";
        public const string WorkSheetNameCLISwitch = "worksheet";
        public const string HeaderRowCLISwitch = "headerrow";
        public const string SettingsCLISwitch = "settings";
        public const string ReportCLISwitch = "report";

        public const string CommandLineUsageFormat =
@"Commands:
testcasemigratorplus excel        Migrates workitems from excel source to TCM

testcasemigratorplus mht          Migrates workitems from mht/word source to TCM";

        public const string ExcelCommandLineUsageFormat =
@"testcasemigratorplus excel [" + SourceFileCLISwitch + @": ExcelSourcePath]
                            [" + WorkSheetNameCLISwitch + @": WorkSheetName]
                            [" + HeaderRowCLISwitch + @": HeaderRow]
                            [" + CollectionCLISwitch + @": TeamProjectCollectionUrl]
                            [" + ProjectCLISwitch + @": TeamProject]
                            [" + WorkItemTypeCLISwitch + @": WorkItemType]
                            [" + SettingsCLISwitch + @": SettingsFilePath]
                            [" + ReportCLISwitch + @": ReportFolderPath]


HeaderRow: Excel row containing field names";

        public const string MHTCommandLineUsageFormat =
@"testcasemigratorplus mht [" + SourceFileCLISwitch + @": MHTSourcePath]
                          [" + CollectionCLISwitch + @": TeamProjectCollectionUrl]
                          [" + ProjectCLISwitch + @": TeamProject]
                          [" + WorkItemTypeCLISwitch + @": WorkItemType]
                          [" + SettingsCLISwitch + @": SettingsFilePath]
                          [" + ReportCLISwitch + @": ReportFolderPath]


MHTSourcePath: Text file containing list of mht/word files path";


        #endregion

        #region Properties

        /// <summary>
        /// UI Thread Context. Required to do Ui related stuff in UI thread
        /// </summary>
        public static SynchronizationContext UISynchronizationContext
        {
            get;
            set;
        }

        /// <summary>
        /// Current Application
        /// </summary>
        public static App Application
        {
            get
            {
                if (m_application == null)
                {
                    m_application = new App();
                }
                return m_application;
            }
        }

        /// <summary>
        /// Command Line arguments
        /// </summary>
        public static Dictionary<CommandLineSwitch, string> CommandLineArguments
        {
            get;
            private set;
        }

        /// <summary>
        /// Wizard Info
        /// </summary>
        public WizardInfo WizardInfo
        {
            get;
            set;
        }

        #endregion

        #region Static Constructor

        static App()
        {
            // Initialize the command switches
            s_commandSwitchLookup[ExcelSwitch] = CommandLineSwitch.Excel;
            s_commandSwitchLookup[MHTSwitch] = CommandLineSwitch.MHT;
            s_commandSwitchLookup[CollectionCLISwitch] = CommandLineSwitch.TFSCollection;
            s_commandSwitchLookup[ProjectCLISwitch] = CommandLineSwitch.Project;
            s_commandSwitchLookup[WorkItemTypeCLISwitch] = CommandLineSwitch.WorkItemType;
            s_commandSwitchLookup[SourceFileCLISwitch] = CommandLineSwitch.SourcePath;
            s_commandSwitchLookup[WorkSheetNameCLISwitch] = CommandLineSwitch.WorkSheetName;
            s_commandSwitchLookup[HeaderRowCLISwitch] = CommandLineSwitch.HeaderRow;
            s_commandSwitchLookup[SettingsCLISwitch] = CommandLineSwitch.SettingsPath;
            s_commandSwitchLookup[ReportCLISwitch] = CommandLineSwitch.ReportPath;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Application Entry Point.
        /// </summary>
        [STAThread]
        public static void Main(String[] args)
        {
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(Application.CurrentDomain_UnhandledException);
            if (args.Length > 0)
            {
                Application.ParseCLIArguments(args);
            }
            else
            {
                Application.InitializeComponent();
                Application.Run();
            }
        }

        /// <summary>
        /// Call passed delegate in UIThread Context if available
        /// </summary>
        /// <param name="del"></param>
        /// <param name="value"></param>
        public static void CallMethodInUISynchronizationContext(SendOrPostCallback del, object value)
        {
            if (del != null)
            {
                if (UISynchronizationContext != null)
                {
                    UISynchronizationContext.Send(new System.Threading.SendOrPostCallback(del), value);
                }
                else
                {
                    del(value);
                }
            }
        }

        #endregion

        #region private methods


        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException -= new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

            MessageHelper.ShowMessageWindow = null;
            if (WizardInfo != null)
            {
                WizardInfo.Dispose();
            }
            ExcelParser.Quit();
            MHTParser.Quit();
        }

        [SuppressMessageAttribute("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        private void ParseCLIArguments(string[] args)
        {
            Dictionary<CommandLineSwitch, string> arguments;

            SetConsoleMode();

            try
            {

                if (ParseArguments(args, out arguments))
                {
                    VerifyArguments(arguments);

                    WizardInfo = new WizardInfo();

                    Console.Write("Loading Data Source...");

                    if (arguments.ContainsKey(CommandLineSwitch.Excel))
                    {
                        InitializeExcelDataSource(arguments);
                    }
                    else if (arguments.ContainsKey(CommandLineSwitch.MHT))
                    {
                        InitializeMHTDataSource(arguments);
                    }

                    Console.Write("\nInitializing TFS Server Connection...");

                    WizardInfo.WorkItemGenerator = new WorkItemGenerator(arguments[CommandLineSwitch.TFSCollection], arguments[CommandLineSwitch.Project]);
                    if (WizardInfo.DataSourceType == DataSourceType.MHT)
                    {
                        WizardInfo.WorkItemGenerator.AddTestStepsField = false;
                    }
                    else if (WizardInfo.DataSourceType == DataSourceType.Excel)
                    {
                        WizardInfo.WorkItemGenerator.AddTestStepsField = true;
                    }

                    if (!WizardInfo.WorkItemGenerator.WorkItemTypeNames.Contains(arguments[CommandLineSwitch.WorkItemType]))
                    {
                        throw new WorkItemMigratorException("Wrong Type Name:" + arguments[CommandLineSwitch.WorkItemType] + " is given", null, null);
                    }

                    WizardInfo.WorkItemGenerator.SelectedWorkItemTypeName = arguments[CommandLineSwitch.WorkItemType];

                    Console.Write("\nLoading Settings...");
                    WizardInfo.LoadSettings(arguments[CommandLineSwitch.SettingsPath]);


                    if (WizardInfo.DataSourceType == DataSourceType.Excel)
                    {
                        WizardInfo.Reporter.ReportFile = Path.Combine(arguments[CommandLineSwitch.ReportPath], "Report.xls");
                    }
                    else
                    {
                        WizardInfo.Reporter.ReportFile = Path.Combine(arguments[CommandLineSwitch.ReportPath], "Report.xml");
                    }

                    WizardInfo.LinksManager = new LinksManager(WizardInfo);

                    Console.WriteLine("\n\nStarting Migration:\n");

                    WizardInfo.Migrator.PostMigration = ShowMigrationStatus;

                    var resultSourceWorkItems = WizardInfo.Migrator.Migrate(WizardInfo.WorkItemGenerator);
                    WizardInfo.ResultWorkItems = resultSourceWorkItems;

                    Console.Write("\n\nPublishing Report:");

                    foreach (var dsWorkItem in resultSourceWorkItems)
                    {
                        WizardInfo.Reporter.AddEntry(dsWorkItem);
                    }
                    WizardInfo.Reporter.Publish();

                    if (WizardInfo.DataSourceType == DataSourceType.Excel)
                    {
                        ProcessLinks();
                    }
                }
                else
                {
                    if (arguments.Count > 0)
                    {
                        if (arguments.ContainsKey(CommandLineSwitch.MHT))
                        {
                            DisplayMHTUsage();
                        }
                        else if (arguments.ContainsKey(CommandLineSwitch.Excel))
                        {
                            DisplayExcelUsage();
                        }
                        else
                        {
                            DisplayGenericUsage();
                        }
                    }
                    else
                    {
                        DisplayGenericUsage();
                    }
                }
            }
            catch (WorkItemMigratorException te)
            {
                Console.WriteLine(te.Args.Title);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                if (WizardInfo != null)
                {
                    WizardInfo.Dispose();
                }
                MHTParser.Quit();
                ExcelParser.Quit();
            }
            Console.Write("\n\nPress Enter to exit...\n");

        }

        private void ProcessLinks()
        {
            Console.Write("\n\nProcessing Links:\n\n");


            foreach (var resultWorkItem in WizardInfo.ResultWorkItems)
            {
                if (resultWorkItem is SkippedSourceWorkItem)
                {
                    string category = WizardInfo.WorkItemGenerator.WorkItemTypeToCategoryMapping[WizardInfo.WorkItemGenerator.SelectedWorkItemTypeName];
                    int tfsId = WizardInfo.LinksManager.WorkItemCategoryToIdMappings[category][resultWorkItem.SourceId].TfsId;

                    AddWorkItemInTestSuites(tfsId, resultWorkItem.TestSuites);
                    continue;
                }

                var workItemStatus = new WorkItemMigrationStatus();
                workItemStatus.SourceId = resultWorkItem.SourceId;
                workItemStatus.SessionId = WizardInfo.LinksManager.SessionId;
                workItemStatus.WorkItemType = WizardInfo.WorkItemGenerator.SelectedWorkItemTypeName;

                PassedSourceWorkItem passedWorkItem = resultWorkItem as PassedSourceWorkItem;

                if (passedWorkItem != null)
                {
                    workItemStatus.Status = Status.Passed;
                    workItemStatus.TfsId = passedWorkItem.TFSId;
                }
                else
                {
                    var failedWorkItem = resultWorkItem as FailedSourceWorkItem;
                    if (failedWorkItem != null)
                    {
                        workItemStatus.Status = Status.Failed;
                        workItemStatus.TfsId = -1;
                        workItemStatus.Message = failedWorkItem.Error;
                    }
                }
                if (!string.IsNullOrEmpty(workItemStatus.SourceId))
                {
                    WizardInfo.LinksManager.UpdateIdMapping(workItemStatus.SourceId, workItemStatus);
                }

                if (passedWorkItem != null)
                {
                    foreach (Link link in passedWorkItem.Links)
                    {
                        WizardInfo.LinksManager.AddLink(link);
                    }
                    foreach (string testSuite in passedWorkItem.TestSuites)
                    {
                        WizardInfo.WorkItemGenerator.AddWorkItemToTestSuite(passedWorkItem.TFSId, testSuite);
                    }
                }
            }
            WizardInfo.LinksManager.Save();
            WizardInfo.LinksManager.PublishReport();
        }

        private bool ShowMigrationStatus(ISourceWorkItem sourceWorkItem)
        {
            m_processedWorkItemCount++;
            Type type = sourceWorkItem.GetType();

            if (type == typeof(WarningSourceWorkItem))
            {
                m_warningWorkItemCount++;
            }
            else if (type == typeof(PassedSourceWorkItem))
            {
                m_passedWorkItemCount++;
            }
            else if (type == typeof(FailedSourceWorkItem))
            {
                m_failedWorkItemCount++;
            }
            Console.Write("\rProcessing {0} of {1} work item: {2} Migrated successfully, {3} Migrated successfully with warning, {4} Failed to migrate",
                           m_processedWorkItemCount,
                           WizardInfo.Migrator.SourceWorkItems.Count,
                           m_passedWorkItemCount,
                           m_warningWorkItemCount,
                           m_failedWorkItemCount);

            PassedSourceWorkItem passedWorkItem = sourceWorkItem as PassedSourceWorkItem;
            if (passedWorkItem != null)
            {
                AddWorkItemInTestSuites(passedWorkItem.TFSId, passedWorkItem.TestSuites);
            }

            return true;
        }

        private void AddWorkItemInTestSuites(int tfsId, IList<string> testSuites)
        {
            foreach (string testSuite in testSuites)
            {
                WizardInfo.WorkItemGenerator.AddWorkItemToTestSuite(tfsId, testSuite);
            }
        }

        private void VerifyArguments(Dictionary<CommandLineSwitch, string> arguments)
        {
            Dictionary<CommandLineSwitch, bool> requiredSwitches = new Dictionary<CommandLineSwitch, bool>();
            requiredSwitches.Add(CommandLineSwitch.TFSCollection, false);
            requiredSwitches.Add(CommandLineSwitch.Project, false);
            requiredSwitches.Add(CommandLineSwitch.WorkItemType, false);
            requiredSwitches.Add(CommandLineSwitch.SettingsPath, false);
            requiredSwitches.Add(CommandLineSwitch.ReportPath, false);
            requiredSwitches.Add(CommandLineSwitch.SourcePath, false);

            Dictionary<CommandLineSwitch, string> cliSwitches = new Dictionary<CommandLineSwitch, string>();
            cliSwitches.Add(CommandLineSwitch.TFSCollection, CollectionCLISwitch);
            cliSwitches.Add(CommandLineSwitch.Project, ProjectCLISwitch);
            cliSwitches.Add(CommandLineSwitch.WorkItemType, WorkItemTypeCLISwitch);
            cliSwitches.Add(CommandLineSwitch.SettingsPath, SettingsCLISwitch);
            cliSwitches.Add(CommandLineSwitch.ReportPath, ReportCLISwitch);
            cliSwitches.Add(CommandLineSwitch.SourcePath, SourceFileCLISwitch);


            foreach (var kvp in arguments)
            {
                if (requiredSwitches.ContainsKey(kvp.Key))
                {
                    requiredSwitches[kvp.Key] = true;
                }

                switch (kvp.Key)
                {
                    case CommandLineSwitch.Excel:
                        if (!(arguments.ContainsKey(CommandLineSwitch.SourcePath) &&
                             arguments.ContainsKey(CommandLineSwitch.WorkSheetName) &&
                             arguments.ContainsKey(CommandLineSwitch.HeaderRow)))
                        {
                            throw new WorkItemMigratorException(ExcelCommandLineUsageFormat, null, null);
                        }
                        break;

                    case CommandLineSwitch.MHT:
                        if (!arguments.ContainsKey(CommandLineSwitch.SourcePath))
                        {
                            throw new WorkItemMigratorException(MHTCommandLineUsageFormat, null, null);
                        }
                        break;

                    case CommandLineSwitch.SourcePath:
                        if (string.IsNullOrEmpty(kvp.Value))
                        {
                            throw new WorkItemMigratorException(String.Format(System.Globalization.CultureInfo.CurrentCulture, "Source Path is not specified"), null, null);
                        }
                        if (!File.Exists(kvp.Value) && !Directory.Exists(kvp.Value))
                        {
                            throw new WorkItemMigratorException(String.Format(System.Globalization.CultureInfo.CurrentCulture, "Source Paths does not exist at {0}", kvp.Value), null, null);
                        }
                        break;

                    case CommandLineSwitch.SettingsPath:
                        if (string.IsNullOrEmpty(kvp.Value))
                        {
                            throw new WorkItemMigratorException(String.Format(System.Globalization.CultureInfo.CurrentCulture, "Settings file path is not specified"), null, null);
                        }
                        if (!File.Exists(kvp.Value))
                        {
                            throw new WorkItemMigratorException(String.Format(System.Globalization.CultureInfo.CurrentCulture, "Settings file is not found at {0}", kvp.Value), null, null);
                        }
                        if (String.CompareOrdinal(Path.GetExtension(kvp.Value), ".xml") != 0)
                        {
                            throw new WorkItemMigratorException("Pleas provide valid settings file(xml)", null, null);
                        }
                        break;

                    case CommandLineSwitch.ReportPath:
                        if (string.IsNullOrEmpty(kvp.Value))
                        {
                            throw new WorkItemMigratorException("Report folder path is not specified", null, null);
                        }
                        try
                        {
                            if (!Directory.Exists(kvp.Value))
                            {
                                Directory.CreateDirectory(kvp.Value);
                            }
                        }
                        catch (IOException ioEx)
                        {
                            throw new WorkItemMigratorException("Unable to create Report Folder at " + kvp.Value + "\n\n" + ioEx.ToString(), null, null);
                        }
                        catch (ArgumentException argEx)
                        {
                            throw new WorkItemMigratorException("Unable to create Report Folder at " + kvp.Value + "\n\n" + argEx.ToString(), null, null);
                        }
                        break;

                    default:
                        break;
                }
            }
            string error = string.Empty;
            foreach (var kvp in requiredSwitches)
            {
                if (!kvp.Value)
                {
                    error += "'" + cliSwitches[kvp.Key] + "' switch is not specified\n";
                }
            }
            if (!string.IsNullOrEmpty(error))
            {
                throw new WorkItemMigratorException(error, null, null);
            }
        }

        private bool IsMHTFile(string SampleMHTFilePath)
        {
            string extension = Path.GetExtension(SampleMHTFilePath);
            if (String.CompareOrdinal(extension, ".mht") == 0 ||
                String.CompareOrdinal(extension, ".mhtml") == 0 ||
                String.CompareOrdinal(extension, ".doc") == 0 ||
                String.CompareOrdinal(extension, ".docx") == 0)
            {
                return true;
            }
            return false;
        }

        private void InitializeMHTDataSource(Dictionary<CommandLineSwitch, string> arguments)
        {
            WizardInfo.DataSourceType = DataSourceType.MHT;
            string sourcePath = arguments[CommandLineSwitch.SourcePath];

            WizardInfo.DataStorageInfos = new List<IDataStorageInfo>();
            if (File.Exists(sourcePath))
            {
                if (String.CompareOrdinal(Path.GetExtension(sourcePath), ".txt") != 0)
                {
                    throw new WorkItemMigratorException("MHT Source should be a text file", null, null);
                }
                using (var tw = new StreamReader(sourcePath))
                {
                    while (!tw.EndOfStream)
                    {
                        string mhtFilePath = tw.ReadLine();
                        try
                        {
                            var mhtInfo = new MHTStorageInfo(mhtFilePath);
                            WizardInfo.DataStorageInfos.Add(mhtInfo);
                        }
                        catch (ArgumentException)
                        { }
                    }
                }
            }
            else if (Directory.Exists(sourcePath))
            {
                foreach (string mhtFile in Directory.GetFiles(sourcePath, "*", SearchOption.AllDirectories))
                {
                    try
                    {
                        if (IsMHTFile(mhtFile) && !Path.GetFileNameWithoutExtension(mhtFile).StartsWith("~", StringComparison.Ordinal))
                        {
                            MHTStorageInfo info = new MHTStorageInfo(mhtFile);
                            WizardInfo.DataStorageInfos.Add(info);
                        }
                    }
                    catch (FileFormatException)
                    { }
                }
            }

            if (WizardInfo.DataStorageInfos.Count == 0)
            {
                throw new WorkItemMigratorException("MHT source is not having any valid mht file path", null, null);
            }
            WizardInfo.DataSourceParser = new MHTParser(WizardInfo.DataStorageInfos[0] as MHTStorageInfo);
        }

        private void InitializeExcelDataSource(Dictionary<CommandLineSwitch, string> arguments)
        {
            WizardInfo.DataSourceType = DataSourceType.Excel;
            ExcelStorageInfo excelInfo = null;
            try
            {
                excelInfo = new ExcelStorageInfo(arguments[CommandLineSwitch.SourcePath]);
            }
            catch (ArgumentException)
            {
                throw new WorkItemMigratorException("Incorrect excel file type", null, null);
            }
            excelInfo.WorkSheetName = arguments[CommandLineSwitch.WorkSheetName];
            excelInfo.RowContainingFieldNames = arguments[CommandLineSwitch.HeaderRow];
            WizardInfo.DataStorageInfos = new List<IDataStorageInfo> { excelInfo };
            WizardInfo.DataSourceParser = new ExcelParser(excelInfo);
            WizardInfo.DataSourceParser.ParseDataSourceFieldNames();
        }

        private void SetConsoleMode()
        {
            AttachConsole(ATTACH_PARENT_PROCESS);
            AutoWaitCursor.IsConsoleMode = true;

            Console.Write("\r                                                                                                                                                             \r");
        }

        private static bool ParseArguments(string[] args, out Dictionary<CommandLineSwitch, string> arguments)
        {
            arguments = new Dictionary<CommandLineSwitch, string>();
            // Iterate through all the args from the command line
            foreach (string argument in args)
            {
                // Check for null string
                if (String.IsNullOrEmpty(argument))
                {
                    continue;
                }

                string paramName = String.Empty;
                string paramValue = String.Empty;

                // Try and split the /foo:bar pair into a meaningful value
                GetArgumentParts(argument.Trim(), out paramName, out paramValue);

                if (paramValue != null)
                {
                    paramValue = paramValue.Replace("\r", "");
                }

                if (String.CompareOrdinal(paramName, "/?") == 0)
                {
                    return false;
                }

                if (arguments.Count > 0)
                {
                    if (!paramName.StartsWith("/", StringComparison.Ordinal))
                    {
                        return false;
                    }
                    paramName = paramName.Substring(1);
                    if (String.CompareOrdinal(paramName, MHTSwitch) == 0 ||
                        String.CompareOrdinal(paramName, ExcelSwitch) == 0)
                    {
                        return false;
                    }
                }
                else
                {
                    if (String.CompareOrdinal(paramName, MHTSwitch) != 0 &&
                        String.CompareOrdinal(paramName, ExcelSwitch) != 0)
                    {
                        return false;
                    }
                }

                CommandLineSwitch argumentType = CommandLineSwitch.Unknown;

                try
                {
                    s_commandSwitchLookup.TryGetValue(paramName, out argumentType);
                }
                catch (ArgumentException)
                {
                    throw new WorkItemMigratorException("Switch:" + paramName + " is not valid", null, null);
                }

                if (argumentType == CommandLineSwitch.Unknown)
                {
                    throw new WorkItemMigratorException("Unknown Switch:" + paramName, null, null);
                }
                else if (arguments.ContainsKey(argumentType))
                {
                    throw new WorkItemMigratorException("Switch:" + paramName + " is passed multiple timess", null, null);
                }
                arguments.Add(argumentType, paramValue);
            }
            return true;
        }

        private static void GetArgumentParts(string argument, out string paramName, out string paramValue)
        {
            paramName = null;
            paramValue = null;

            int splitPosition = argument.IndexOf(':');

            // Make sure the position of the split is somewhere meaningful
            if ((splitPosition < 1) || (splitPosition == argument.Length))
            {
                paramName = argument;
                return;
            }

            paramName = argument.Substring(0, splitPosition);
            paramValue = argument.Substring(splitPosition + 1);

            // see if it starts AND ends with "'s and remove them
            if (paramValue.Length > 1)
            {
                if (paramValue.StartsWith("\"", StringComparison.Ordinal) && paramValue.EndsWith("\"", StringComparison.Ordinal))
                {
                    paramValue = paramValue.Substring(1, paramValue.Length - 1);
                }
            }
        }

        private static void DisplayGenericUsage()
        {
            Console.WriteLine(CommandLineUsageFormat);

        }

        private static void DisplayExcelUsage()
        {
            Console.WriteLine(ExcelCommandLineUsageFormat);

        }

        private static void DisplayMHTUsage()
        {
            Console.WriteLine(MHTCommandLineUsageFormat);

        }

        #endregion
    }

    internal enum CommandLineSwitch
    {
        Unknown,
        Excel,
        MHT,
        SourcePath,
        WorkSheetName,
        HeaderRow,
        TFSCollection,
        Project,
        WorkItemType,
        SettingsPath,
        ReportPath
    }
}
