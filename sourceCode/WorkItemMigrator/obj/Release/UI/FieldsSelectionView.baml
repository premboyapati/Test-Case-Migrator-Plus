   M S B A M L   `   `   `  ���� [  WTestCaseMigratorPlus, Version=1.0.0.0, Culture=neutral, PublicKeyToken=9d26ba8cb6a39082u?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator1Microsoft.VisualStudio.TestTools.WorkItemMigrator  K   EMicrosoft.VisualStudio.TestTools.WorkItemMigrator.FieldsSelectionView   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    ZWorkItemMigrator?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml       � 	Resources     FieldsListBoxItemStyle �� 
TargetType �� Property �� ValueT   NMicrosoft.VisualStudio.TestTools.WorkItemMigrator.BooleanToVisibilityConverter   BooleanToVisibilityConverter��%[   &	        &	 B     �� " ��.���  � ��$ Stretch=���  � ��$ Center=�  .���  � RowDefinitions �� $
��Auto ��� $��* �.+�� $��0�� �� $
��Auto ��� $
��Auto ��� $
��Auto �.+��� $��0��$/��)Sample mht/word file for field selection:��� $��1��$��0,5,0,0q� � ColumnDefinitions �� $��* ��� $
��Auto �.+��� $��0��$��5,0,0,0q����� *
SourcePath+�� -   $��1��$��3,0,0,0q�$
��_...��� $��2��$��Right=�$��0,10,0,0q� �� $
��Auto ��� $
��Auto �.+��� $��16��$��16��$��0��$��Center=�$��0,0,3,0q�  Word_32#���  �� $��1���� -   4  .System.Windows.Automation.AutomationProperties   AutomationId$' !SelectDataSource.PreviewMHTButton��   Name$ PreviewMHTButton��%!SelectDataSource.PreviewMHTButton�� �� 	IsEnabled �� *IsPreviewEnabled+.�	Preview���� " ��.���  � ��  HyperLinkBrush# �  	 �� Triggers	 P� 
 P� Property
 �@��	IsEnabled P� Value .  . ���  � ��
 �����.�� $��1�� �� $
��Auto ��� $��* ��� $
��Auto �.+��� $��Field names:��$��0��$��0,5,0,0q�$
��Bold"���  �� SelectionMode$ Single=�$��Disabled=�#��[  $��0,2,0,0q�$��1������ *Fields+���� .R��� $��Stretch=�
�������$	��0.5q� J� SnapsToDevicePixels . |�.  ._�� $��25�� �� $��* ��� $
��Auto �.+��� $��0��$��CharacterEllipsis=�$��3,0,0,0q�$��Center=� � ToolTip �� *	FieldName+���� *	FieldName+��  ��Name$ DeleteFieldNameButton��-   $��Delete��$��1��y��� *	CanDelete+ �� 	Converter# [ � $��2��{��� *IsPreviewEnabled+ �� $��* ��� $
��Auto �.+��� -    ��Name$ NewFieldNameInputBox��$��0���� -   $ AddNewFieldNameButton��AddNewFieldNameButton��$ AddNewFieldNameButton��$ AddNewFieldNameButton��$��1��$��3,0,0,0q�.���� $��_Add field name��.������ " ��	 P� 
 � {� .  . ���  �@��
Foreground
 �����.��