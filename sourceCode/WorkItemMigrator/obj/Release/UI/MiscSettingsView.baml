   M S B A M L   `   `   `  ���� [  WTestCaseMigratorPlus, Version=1.0.0.0, Culture=neutral, PublicKeyToken=9d26ba8cb6a39082u?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator1Microsoft.VisualStudio.TestTools.WorkItemMigrator  H   BMicrosoft.VisualStudio.TestTools.WorkItemMigrator.MiscSettingsView   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     \WorkItemMigrator?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator       � 	Resources  T   NMicrosoft.VisualStudio.TestTools.WorkItemMigrator.BooleanToVisibilityConverter    BooleanToVisibilityConverter��%   &	          .���  � RowDefinitions �� $
��Auto ��� $
��Auto ��� $
��Auto �.+�� $��
10,10,30,0q�$��0��y��� *AreMiscSettingsVisible+ �� 	Converter# [  .+��� 
�������$	��0.5q�$	��Top=� �� Padding$ 25q�._��  �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto �.+��� 4  .System.Windows.Automation.AutomationProperties   AutomationId$- 'MiscSettingsView.MultiLineSenseCheckBox��   Name$ MultiLineSenseCheckBox��$��0��+'MiscSettingsView.MultiLineSenseCheckBox�� d� 	IsChecked �� *IsMultiLineSenseEnabled+.��TRMultiple test steps are present in single cell separated by new line character(\n)�� -    ��Name$ ParameterizationCheckbox��$/ )MiscSettingsView.ParameterizationCheckBox��$ ParameterizationCheckBox��-)MiscSettingsView.ParameterizationCheckBox��$��0,20,0,0q�$��1�� �� *IsParameterizationChecked+.��CAParse parameters present in test steps with following delimeters:� $��2��$��20,3,0,0q�{���  �� ElementName$ ParameterizationCheckbox��	 �� Path$	 	IsChecked�
 � ColumnDefinitions
 �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto �.+��� StartDelimeter��$��Center=�$��0��$��Start Delimeter:���� 3/MiscSettingsView.StartParameterizationDelimeter��$5 /MiscSettingsView.StartParameterizationDelimeter��$$ StartParameterizationDelimeter��$��1��$��30��$��3,0,0,0q����� * StartParameterizationDelimeter+�� EndDelimeter��$��Center=�$��2��$��15,0,0,0q�$��End Delimeter:���� 1-MiscSettingsView.EndParameterizationDelimeter��$3 -MiscSettingsView.EndParameterizationDelimeter��$" EndParameterizationDelimeter��$��30��$��3,0,0,0q�$��3������ *EndParameterizationDelimeter+�� $��3��$���|Example: @{john} from test step to extract 'john' as parameter where '@{' and '}' are start and end delimeters respectively.��$
��Wrap=�$��10,5,0,0q��� $��4��$����Note: Each parameter parsed will be created as a separate column within the parameter table. Data for parameters is not migrated.��$
��Wrap=�$��10,6,0,0q��� &"MiscSettingsView.MiscSettingsLabel��$��
10,-14,0,0q�$
��Left=�  WizardPageBackgroundColor#���  $	��Top=�.��'%Miscellaneous settings for Test Steps� $��1��$��
25,30,50,0q� �� $
��Auto ��� $
��Auto �
 �� $��* ��� $
��Auto �.+��� $��0��$��2��OutputMappingsFileText��$��0,0,0,5q�$
��Wrap=�$]��WSave mappings(fields, data and links) and miscellaneous settings to this file location:���� $0 *MiscSettingsView.OutputMappingsFileTextBox��$ OutputMappingsFileTextBox��.*MiscSettingsView.OutputMappingsFileTextBox�� �� 
IsReadOnly . 
�������$��1��$��0������ *OutputMappingsFile+�� -    ��Name$$ OutputMappingsFileBrowseButton��$5 /MiscSettingsView.OutputMappingsFileBrowseButton��$$ OutputMappingsFileBrowseButton��$��1��$��1��$
��_...��� $��2��$��
25,30,50,0q� �� $
��Auto ��� $
��Auto �
 �� $��* ��� $
��Auto �.+��� $��0��$��2��$��0,0,0,5q�$
��Wrap=�OutputFileText��$/��)Save log files to the following location:���� $* $MiscSettingsView.ReportFolderTextBox��$ ReportFolderTextBox��($MiscSettingsView.ReportFolderTextBox��$��1�� . $��0��
����������� *ReportFolder+�� -   $ ReportFolderBrowseButton��$/ )MiscSettingsView.ReportFolderBrowseButton��$ ReportFolderBrowseButton��$��1��$��1��$
��_...��