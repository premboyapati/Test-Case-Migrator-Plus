   M S B A M L   `   `   `  ����[  WTestCaseMigratorPlus, Version=1.0.0.0, Culture=neutral, PublicKeyToken=9d26ba8cb6a39082u?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator1Microsoft.VisualStudio.TestTools.WorkItemMigrator  5   "   H   BMicrosoft.VisualStudio.TestTools.WorkItemMigrator.SettingsFileView   5      \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   "   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   "   \WorkItemMigrator?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator     5   "     � 	Resources  5      O  IMicrosoft.VisualStudio.TestTools.WorkItemMigrator.InverseBooleanConverter   InverseBooleanConverter��%   &	          5   
   63   5      .��5
      � $��15,20,100,0q�6    � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   .+�5   
   	� 4  .System.Windows.Automation.AutomationProperties   AutomationId$1 +SettingsFileView.NotLoadSettingsRadioButton��6      Name$  NotLoadSettingsRadioButton��5      $��0��5       	� 	GroupName$" LoadSettingsRadioButtonGroup��5       d� 	IsChecked 5      ��  �� ElementName$ LoadSettingsRadioButton�� �� Path$ 	IsChecked� �� 	Converter# [  .��5      �� NotLoadSettingText��$
��Wrap=�5      $��Create new settings file��5      5      	� 5   
   -   5        	 	�Name$	 LoadSettingsRadioButton��5      $. (SettingsFileView.LoadSettingsRadioButton��5      $ LoadSettingsRadioButton��5      $��1��5      $" LoadSettingsRadioButtonGroup��5      $��0,15,0,0q�5        5!      �� *LoadSettings+.��5"      �� LoadSettingText��$
��Wrap=�5#      $*��$Load from an existing settings file:��5$      5%      � 5'   
   $��2��6   $��15,3,0,0q�5)      {�5(      �� *LoadSettings+
 � ColumnDefinitions
 5*      �� 5+      �� 5,      $
��Auto �6#   .+�5/      J  DMicrosoft.VisualStudio.TestTools.WorkItemMigrator.WITMigratorTextBox  J  DMicrosoft.VisualStudio.TestTools.WorkItemMigrator.WITMigratorTextBox   TextChangeAction$) #WITMigratorTextBox_TextChangeAction��55   2   $% SettingsFileView.SettingTextBox��5/   2   $ SettingTextBox��50   2   #SettingsFileView.SettingTextBox��$��0��53   2   ��52   2   �� *SettingsFilePath+   
IsRequired 54   2   ��  �� RelativeSource �� *Self+$ 	IsEnabled�55   2   �� 57      -   5        $( "SettingsFileView.LoadSettingButton��57      $ LoadSettingButton��58      $
��_...��59      $��1��5:      5;      5<      5=      5>      