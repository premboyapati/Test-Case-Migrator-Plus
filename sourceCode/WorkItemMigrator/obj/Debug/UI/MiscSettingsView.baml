   M S B A M L   `   `   `  ����[  WTestCaseMigratorPlus, Version=1.0.0.0, Culture=neutral, PublicKeyToken=9d26ba8cb6a39082u?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator1Microsoft.VisualStudio.TestTools.WorkItemMigrator  5   "   H   BMicrosoft.VisualStudio.TestTools.WorkItemMigrator.MiscSettingsView   5      \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   "   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   "   \WorkItemMigrator?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator     5   "     � 	Resources  5      T   NMicrosoft.VisualStudio.TestTools.WorkItemMigrator.BooleanToVisibilityConverter    BooleanToVisibilityConverter��%   &	          5   
   68   5      .��5	      �  � RowDefinitions 5
   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   .+�5   
   � $��
10,10,30,0q�6   $��0��5      y�5      �� *AreMiscSettingsVisible+ �� 	Converter# [  .+�5      �� 
�������$	��0.5q�5      $	��Top=�5       �� Padding$ 25q�5      ._�5      �  5      �� 5      $
��Auto �6(   �� 5      $
��Auto �6(   �� 5      $
��Auto �6(   �� 5      $
��Auto �6(   �� 5      $
��Auto �6(   .+�5       �� 4  .System.Windows.Automation.AutomationProperties   AutomationId$- 'MiscSettingsView.MultiLineSenseCheckBox��6      Name$ MultiLineSenseCheckBox��5!      $��0��5"      +'MiscSettingsView.MultiLineSenseCheckBox�� d� 	IsChecked 5$      �� *IsMultiLineSenseEnabled+.��6M   TRMultiple test steps are present in single cell separated by new line character(\n)5&      �� 5(      -   5         ��Name$ ParameterizationCheckbox��5-      $/ )MiscSettingsView.ParameterizationCheckBox��5(      $ ParameterizationCheckBox��5)      -)MiscSettingsView.ParameterizationCheckBox��$��0,20,0,0q�5+      $��1��5,       5.      �� *IsParameterizationChecked+.��6O   CAParse parameters present in test steps with following delimeters:50      � 52      $��2��6   $��20,3,0,0q�54      {�53      ��  �� ElementName$ ParameterizationCheckbox��	 �� Path$	 	IsChecked�
 � ColumnDefinitions
 55      �� 56      $
��Auto �6/   �� 57      $
��Auto �6/   �� 58      $
��Auto �6/   �� 59      $
��Auto �6/   .+�5=      �� StartDelimeter��$��Center=�5>   $   $��0��5?   $   $��Start Delimeter:��5@   $   �� 5A      3/MiscSettingsView.StartParameterizationDelimeter��$5 /MiscSettingsView.StartParameterizationDelimeter��5B   "   $$ StartParameterizationDelimeter��5C   "   $��1��5D   "   $��30��5E   "   $��3,0,0,0q�5F   "   ��5G   "   �� * StartParameterizationDelimeter+�� 5I      EndDelimeter��$��Center=�5J   $   $��2��5K   $   $��15,0,0,0q�5L   $   $��End Delimeter:��5M   $   �� 5N      1-MiscSettingsView.EndParameterizationDelimeter��$3 -MiscSettingsView.EndParameterizationDelimeter��5O   "   $" EndParameterizationDelimeter��5P   "   $��30��5Q   "   $��3,0,0,0q�5R   "   $��3��5S   "   ��5T   "   �� *EndParameterizationDelimeter+5U      �� 5W      $��3��6    $���|Example: @{john} from test step to extract 'john' as parameter where '@{' and '}' are start and end delimeters respectively.��5X       $
��Wrap=�5Y       $��10,5,0,0q�5Z       �� 5\      $��4��6    $����Note: Each parameter parsed will be created as a separate column within the parameter table. Data for parameters is not migrated.��5]       $
��Wrap=�5^       $��10,6,0,0q�5_       5a      5b      �� 5d      &"MiscSettingsView.MiscSettingsLabel��$��
10,-14,0,0q�5e      $
��Left=�5f        WizardPageBackgroundColor#���  5g      $	��Top=�5h      .��6,   '%Miscellaneous settings for Test Steps5j      5k      � 5m   
   $��1��6   $��
25,30,50,0q�5n       5o      �� 5p      $
��Auto �6    �� 5q      $
��Auto �6    
 5t      �� 5u      $��* �6#   �� 5v      $
��Auto �6#   .+�5y      �� $��0��6   $��2��5z      OutputMappingsFileText��$��0,0,0,5q�5|      $
��Wrap=�5}      $]��WSave mappings(fields, data and links) and miscellaneous settings to this file location:��5~      �� 5�      $0 *MiscSettingsView.OutputMappingsFileTextBox��6   $ OutputMappingsFileTextBox��5�      .*MiscSettingsView.OutputMappingsFileTextBox�� �� 
IsReadOnly . 
�������$��1��5�      $��0��5�      ��5�      �� *OutputMappingsFile+�� 5�      -   5         ��Name$$ OutputMappingsFileBrowseButton��5�      $5 /MiscSettingsView.OutputMappingsFileBrowseButton��5�      $$ OutputMappingsFileBrowseButton��5�      $��1��5�      $��1��5�      $
��_...��5�      5�      5�      � 5�   
   $��2��6   $��
25,30,50,0q�5�       5�      �� 5�      $
��Auto �6    �� 5�      $
��Auto �6    
 5�      �� 5�      $��* �6#   �� 5�      $
��Auto �6#   .+�5�      �� $��0��6   $��2��5�      $��0,0,0,5q�5�      $
��Wrap=�5�      OutputFileText��$/��)Save log files to the following location:��5�      �� 5�      $* $MiscSettingsView.ReportFolderTextBox��6   $ ReportFolderTextBox��5�      ($MiscSettingsView.ReportFolderTextBox��$��1��5�       . $��0��5�      
���������5�      �� *ReportFolder+�� 5�      -   5        $ ReportFolderBrowseButton��5�      $/ )MiscSettingsView.ReportFolderBrowseButton��5�      $ ReportFolderBrowseButton��5�      $��1��5�      $��1��5�      $
��_...��5�      5�      5�      5�      5�      