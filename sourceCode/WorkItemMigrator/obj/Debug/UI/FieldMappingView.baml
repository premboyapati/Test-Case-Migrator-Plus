   M S B A M L   `   `   `  ����[  WTestCaseMigratorPlus, Version=1.0.0.0, Culture=neutral, PublicKeyToken=9d26ba8cb6a39082u?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator1Microsoft.VisualStudio.TestTools.WorkItemMigrator  5   "   H   BMicrosoft.VisualStudio.TestTools.WorkItemMigrator.FieldMappingView   5      \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   "   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   "   \WorkItemMigrator?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator     5   "     � 	Resources  5      T   NMicrosoft.VisualStudio.TestTools.WorkItemMigrator.BooleanToVisibilityConverter    BooleanToVisibilityConverter[   UMicrosoft.VisualStudio.TestTools.WorkItemMigrator.InverseBooleanToVisibilityConverter ' #InverseBooleanToVisibilityConverter  DSFieldTemplate � ToolTip �� 
TargetType �� Property �� Value �� Triggers �� Binding �� RelativeSource �� AncestorType)  #System.Windows.Controls.DataGridRow	 �� Path
 �� Value  WIFieldTemplate �� Mode �� 	Converter ��Name �� SelectedValue��%i  &	        &	      &	 &     &	 �    7  7     5   
   68     5   
   6?   �� 5
   
   .R�5      �� $��CharacterEllipsis=�5      $��Center=�5      ��5      �� *DataSourceField+ 5      �� *DataSourceField+��5      �� 5      " ��6   .�5      ��  � ��
 �   �67    5      �� 5       6*   ��  �� "  $	 
IsSelected�
 True5   *   .a�5   "   ��  � ��
 �����6?   5      5      5      5      �� 5   
   .R�5      � .+�5      �� $��Center=�5#      ��5      �� *
TFSField+$ OneWay=�y�5       ��  �� "  $ OneWay=�$	 
IsSelected�8   ��5$      �� 5%      " ��6    .�5&      ��  � ��
 �   �6;    5'      �� 5(   "    6.   �� *IsValidFieldMapping+	
 False5)   .   .a�5*   &   ��  � ��
 �����6C   5+   #   5-      5/      �� 51      $ WIFieldComboBox��6   -   5         52      �� *
TFSField+$ OneWay=�y�53      ��  �� "  $ OneWay=�$	 
IsSelected�8  ��57      �� *AvailableTFSFields+58      59      5;      .��5<      �  � RowDefinitions 5=   
   �� 5>      $��* �6   �� 5?      $
��Auto �6   .+�5B   
   � $��0��6    5C      �� 5D      $
��Auto �6    �� 5E      $��* �6    .+�5H      � $��0��6   $��0,3,0,0q�5I      .+�5J      �� $��Center=�6   �� 5K        WizardPageBackgroundColor#���  6   $��Center=�5L      $
��Left=�5M      ._�5N      �� $��Fields mapping��6     �� Padding$ 3,0,3,0q�5O      $��Center=�5P      5Q      5R      � 5U      $��1��6   $��10,5,0,0q�5V       5W      �� 5X      $��* �6$   �� 5Y      $
��Auto �6$   �� 5Z      $
��Auto �6$   �� 5[      $
��Auto �6$   .+�5^      &   System.Windows.Controls.DataGrid  -   5          Name$ FieldMappingControl��5_         AutoGenerateColumns .     	RowHeight$ 25Q�5a      $��0,5,0,0q�5b         GridLinesVisibility$
 None=�5c         CanUserAddRows .     CanUserDeleteRows .     CanUserResizeRows .     CanUserReorderColumns .     SelectionMode$ Single=�5h        DataGridRowStyle   RowStyle# �  5i      $��0��5j      ��5^      �� *FieldMappingRows+   Columns 5k      4  .System.Windows.Controls.DataGridTemplateColumn  5l         CellTemplate# [ 61   ,  &System.Windows.Controls.DataGridColumn   Header$ Source Field��5m   )     DataGridHeaderTextStyle   HeaderStyle# �  5n   )      Width5  /System.Windows.Controls.DataGridLengthConverter$ * 5o   )     5q      $ * 61   $ Destination Field��5r   )   # �  5s   )   # [ 5t   )   5v      � 5x      $��1��6   ��5y      �� 5z      " �6    .�5{      ��  � y�$ 	Collapsed=�6;    5|      �� 5}   "   F  @Microsoft.VisualStudio.TestTools.WorkItemMigrator.DataSourceType
    MHT#
 Z  5~   &    5}   .   �� *WizardInfo.DataSourceType+.a�5   &   ��  � y�$ Visible=�6C   5�   #   5�       5�      �� 5�      $
��Auto �6(   �� 5�      $
��Auto �6(   .+�5�      �� -   5        ! ��Name$! IsFirstLineTitleCheckBox��5�      $��0��5�      $��0,4,0,0q�5�      -)FieldMappingView.IsFirstLineTitleCheckBox��4	  .System.Windows.Automation.AutomationProperties" 	  AutomationId$/" )FieldMappingView.IsFirstLineTitleCheckBox��5�      # 	  Name$# FirstLineTitleCheckBox��5�      $ d� 	IsChecked$ 5�      �� *IsFirstLineTitle+.��6>   .,Use the first line of mht/word file as title5�      �� 5�      -   5        $! IsFileTitleCheckBox��5�      $��1��5�      $��0,4,0,0q�5�      ($FieldMappingView.IsFileTitleCheckBox��$*" $FieldMappingView.IsFileTitleCheckBox��5�      $# IsFileTitleCheckBox��5�      $ 5�      �� *IsFileNameTitle+.��6=   -+Use the file name of mht/word file as title5�      5�      �� 5�      $c��]* Mandatory fields on the destination; all such fields must have the field mapping specified.��6   $��0,7,0,0q�5�      $��2��5�      $
��Wrap=�5�      �� 5�      $���z+ Fields with values that are auto-generated by the server. The values corresponding to these fields will not be migrated.��6   $��0,5,0,0q�5�      $��3��5�      $
��Wrap=�5�      5�      5�      � 5�   
   $��1��6   $��	0,20,0,10q�5�      y�5�      �� *IsTestSuiteVisible+# [   5�      �� 5�      $
��Auto �6    �� 5�      $
��Auto �6    .+�5�      � $��0��6   .+�5�      �� $��Center=�6   �� 5�      #���  6   $��Center=�5�      $
��Left=�5�      ._�5�      �� $��Test suites��6    $ 3,0,3,0q�5�      $��Center=�5�      5�      5�      � 5�      $��1��6   $��10,7,0,0q�5�      % � ColumnDefinitions% 5�      �� 5�      $
��Auto �6'   �� 5�      $��* �6'   .+�5�      �� $��0��6   $��Center=�5�       $,��&Test suites that test case belongs to:��5�       �� 5�      $��1��6   $��5,0,0,0q�5�      ��5�      �� *TestSuiteAvailableFields+& �� SelectedItem& 5�      �� *TestSuiteField+5�      5�      5�      5�      