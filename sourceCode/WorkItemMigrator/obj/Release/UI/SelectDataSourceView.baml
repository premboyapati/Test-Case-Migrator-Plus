   M S B A M L   `   `   `  ���� [  WTestCaseMigratorPlus, Version=1.0.0.0, Culture=neutral, PublicKeyToken=9d26ba8cb6a39082u?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator1Microsoft.VisualStudio.TestTools.WorkItemMigrator  O Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem L   FMicrosoft.VisualStudio.TestTools.WorkItemMigrator.SelectDataSourceView   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     \WorkItemMigrator?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator     6sys&clr-namespace:System;assembly=mscorlib     $��7q�  � 	Resources  S   MMicrosoft.VisualStudio.TestTools.WorkItemMigrator.StringToVisibilityConverter   StringToVisibilityConverterO  IMicrosoft.VisualStudio.TestTools.WorkItemMigrator.InverseBooleanConverter  InverseBooleanConverterT   NMicrosoft.VisualStudio.TestTools.WorkItemMigrator.BooleanToVisibilityConverter   BooleanToVisibilityConverter[   UMicrosoft.VisualStudio.TestTools.WorkItemMigrator.InverseBooleanToVisibilityConverter ' #InverseBooleanToVisibilityConverter��%<   &	        &	      &	 
     &	              .���  � RowDefinitions �� $
��Auto ��� $��* �.+�� $��0��$��0,0,0,10q� � ColumnDefinitions �� $
��Auto ��� $��* �.+��� SelectDataSourceLabel��$��Select Data Source Type:��$��0��$��Center=��� DataSourceComboBox��$��Center=�$��10,0,0,0q�$��1������ *DataSources+ �� SelectedItem �� *SelectedDataSource+� $��1������  �� 
TargetType" �.���  �� Property � y� �� Value$ 	Collapsed=� �� Triggers �� F  @Microsoft.VisualStudio.TestTools.WorkItemMigrator.DataSourceType
   MHT	 �� Value#	 Z 
 �� Binding
 �� *DataSourceType+.a���  � y�$ Visible=� �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto �.+�	� MHTFolderRadioButton��$��0��$��0,12,0,0q� 	� 	GroupName$ MHTInputSelection�� d� 	IsChecked �� *IsMHTFolder+.��*(Select folder containing mht/word files:� $��1��$��15,7,0,0q�{��� *IsMHTFolder+ �� $��* ��� $
��Auto �.+�J  DMicrosoft.VisualStudio.TestTools.WorkItemMigrator.WITMigratorTextBox  -     Name$ MHTFolderPathTextBox��J  DMicrosoft.VisualStudio.TestTools.WorkItemMigrator.WITMigratorTextBox   TextChangeAction$$ MHTFolderPath_TextChangeAction��$��0��"SelectDataSource.MHTFolderPath��4  .System.Windows.Automation.AutomationProperties   AutomationId$$ SelectDataSource.MHTFolderPath��   Name$ MHTFolderPath�� �� 
IsReadOnly .  ���� *MHTFolderPath+   
IsRequired ��  �� RelativeSource �� *Self+ �� Path$ 	IsEnabled��� -    ��Name$ MHTFolderBrowseButton��$��1��$��10,0,0,0q�*&SelectDataSource.MHTFolderBrowseButton��$, &SelectDataSource.MHTFolderBrowseButton��$ MHTFolderBrowseButton��$
��_...��� $��2��$��0,3,0,0q�y��� *IsMHTFolder+ �� 	Converter# [  �� $
��Auto ��� $
��Auto �.+��� $��0��$��
Loading...��y��� *IsLoadingFiles+# [ �� $��1������ $ MHTCount� �� StringFormat$ {0} documents found.��	� !ListOfMHTsFilePathRadioButton��$��3��$��0,12,0,0q�$ MHTInputSelection�� �� *IsMHTFolder+# [ .��DBSelect text file containing list of mht/word files with full path:� $��4��$��15,7,0,0q�{��� *IsMHTFolder+# [  �� $��* ��� $
��Auto �.+�  -   $ ListOfMHTsFilePath��$) #ListOfMHTsFilePath_TextChangeAction��$��0��'#SelectDataSource.ListOfMHTsFilePath��$) #SelectDataSource.ListOfMHTsFilePath��$ ListOfMHTsFilePath�� .  ���� *ListOfMHTsFilePath+ ��  �� *Self+$ 	IsEnabled��� -   $$ ListOfMHTsFilePathBrowseButton��$��1��$��15,0,0,0q�3/SelectDataSource.ListOfMHTsFilePathBrowseButton��$5 /SelectDataSource.ListOfMHTsFilePathBrowseButton��$$ ListOfMHTsFilePathBrowseButton��$
��_...��� $��5��$��0,3,0,0q�y��� *IsMHTFolder+# [  �� $
��Auto ��� $
��Auto �.+��� $��0��$��
Loading...��y��� *IsLoadingFiles+# [ �� $��1������ $ MHTCount�$ {0} documents found.��� $��1������ " �.���  � y�$ 	Collapsed=� ��    Excel#	 Z 
 �� *DataSourceType+.a���  � y�$ Visible=� �� $
��Auto ��� $
��Auto ��� $��* �.+�� $��0��$��0,0,10,0q� �� $
��Auto ��� $
��Auto ��� $��18 �.+��� '#SelectDataSource.ExcelFilePathLabel��$��0��$��0,5,0,0q�$+��%Specify file path for excel workbook:��� $��1��$��0,7,0,0q� �� $��* ��� $
��Auto �.+�  -   $ ExcelFileName��$$ ExcelFileName_TextChangeAction��$��0��%!SelectDataSource.ExcelFileTextBox��$' !SelectDataSource.ExcelFileTextBox��$ ExcelFileTextBox�� .   . ���� *ExcelFilePath+�� -   $! ExcelSourceFileBrowseButton��$��1��$��10,0,0,0q�0,SelectDataSource.ExcelSourceFileBrowseButton��$2 ,SelectDataSource.ExcelSourceFileBrowseButton��$! ExcelSourceFileBrowseButton��$
��_...���� $��2��$��0,3,0,0q�y��� *ExcelFilePath+# [  ���� $ ExcelSheetsCount�$ {0} worksheets found.��� $��	0,10,30,0q�$��1�� �� $
��Auto ��� $
��Auto ��� $
��Auto �.+��� "SelectDataSource.ExcelWorkSeet��$��0��$��0,7,0,0q�$&�� Select the worksheet to migrate:���� -    ��Name$ ExcelSheets��$��1��$��0,7,26,0q�$" SelectDataSource.ExcelSheets��$ ExcelSheets������ *ExcelSheets+ �� *SelectedExcelSheet+� $��2��$��0,15,0,0q� �� $
��Auto ��� $
��Auto �.+��� $,��&Enter row number that has field names:��$��0��$��Center=�SelectDataSource.EnterRow��  -   $ ExcelHeaderRow��$% ExcelHeaderRow_TextChangeAction��$��1��$, &SelectDataSource.ExcelHeaderRowTextBox��$ ExcelHeaderRowTextBox��$��Center=�$��10,0,0,0q�$��70�� . ���� *ExcelHeaderRow+ �� UpdateSourceTrigger$ PropertyChanged=�� $��2��$��0,10,0,0q� �� $
��Auto ��� $��* � �� $��* ��� $
��Auto ��� $
��Auto �.+��� $��Preview:��$��0��$��0���� $��16��$��16��$��0��$��1��$��Center=�$��0,0,3,0q�  Excel_32#���  �� $��0��$��2���� -	   $$ SelectDataSource.PreviewButton��$ PreviewButton��"SelectDataSource.PreviewButton�� �� 	IsEnabled �� *IsPreviewEnabled+.�	Open file���� " ��.���  � ��  HyperLinkBrush# �   P�  P� Property �@��	IsEnabled P� Value .  . ���  � ��
 �����.�&	   System.Windows.Controls.DataGrid	  -
    	 Name$ ColumnMappingControl��$��0,5,0,0q�$
��Auto=� 	  AutoGenerateColumns .    	  GridLinesVisibility$
  None=�! 	  CanUserAddRows! .  " 	  CanUserDeleteRows" .  # 	  CanUserResizeRows# .  $ 	  CanUserReorderColumns$ .  % 	  CanUserSortColumns% .  $��1��$��3������ *Fields+& 	  Columns& 0
  *System.Windows.Controls.DataGridTextColumn
  ,  &System.Windows.Controls.DataGridColumn'   Header$' Field Names��  DataGridHeaderTextStyle(   HeaderStyle#( �  )   Width5  /System.Windows.Controls.DataGridLengthConverter$) * 1  +System.Windows.Controls.DataGridBoundColumn*   Binding* �� 