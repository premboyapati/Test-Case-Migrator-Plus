   M S B A M L   `   `   `  ����[  WTestCaseMigratorPlus, Version=1.0.0.0, Culture=neutral, PublicKeyToken=9d26ba8cb6a39082u?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator1Microsoft.VisualStudio.TestTools.WorkItemMigrator  5      K   EMicrosoft.VisualStudio.TestTools.WorkItemMigrator.FieldsSelectionView   5      \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      ZWorkItemMigrator?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5        � 	Resources  5         FieldsListBoxItemStyle �� 
TargetType �� Property �� ValueT   NMicrosoft.VisualStudio.TestTools.WorkItemMigrator.BooleanToVisibilityConverter   BooleanToVisibilityConverter��%�   &	        &	 �     �� 5   
   " ��5      .�5	      ��  � ��$ Stretch=�5
      �� 5       � ��$ Center=�5      5        5   
   68   5      .��5      �  � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $��* �6   .+�5   
   � $��0��6    5      �� 5      $
��Auto �6    �� 5      $
��Auto �6    �� 5      $
��Auto �6    .+�5      �� $��0��6   $/��)Sample mht/word file for field selection:��5       � 5"      $��1��6   $��0,5,0,0q�5#       � ColumnDefinitions 5$      �� 5%      $��* �6'   �� 5&      $
��Auto �6'   .+�5(      �� $��0��6   $��5,0,0,0q�5)      ��5*      �� *
SourcePath+�� 5,      -   5        $��1��5,      $��3,0,0,0q�5-      $
��_...��5.      5/      50      � 51      $��2��6   $��Right=�52      $��0,10,0,0q�53       54      �� 55      $
��Auto �6'   �� 56      $
��Auto �6'   .+�58      �� $��16��6   $��16��59      $��0��5:      $��Center=�5;      $��0,0,3,0q�5<        Word_32#���  5=      �� 5>      $��1��6   �� 5?      -   5        4  .System.Windows.Automation.AutomationProperties   AutomationId$' !SelectDataSource.PreviewMHTButton��5A          Name$ PreviewMHTButton��5B       %!SelectDataSource.PreviewMHTButton�� �� 	IsEnabled 5@       �� *IsPreviewEnabled+.�5C   J   	Preview��5E      �� 5F      " ��6$   .�5G   "   ��  � ��  HyperLinkBrush# �  5H   )   	 �� Triggers	 5I   "   P� 5J   &   
 P� Property
 �@��	IsEnabled P� Value .  . �5L   *   ��  � ��
 �����5M   1   5N   '   5P      .�5Q   +   5R      5S      5T      5V      � 5X   
   $��1��6    5Y      �� 5Z      $
��Auto �6    �� 5[      $��* �6    �� 5\      $
��Auto �6    .+�5`      �� $��Field names:��6   $��0��5a      $��0,5,0,0q�5b      $
��Bold"�5c      �� 5e       �� SelectionMode$ Single=�5f      $��Disabled=�5g      #��[  5h      $��0,2,0,0q�5i      $��1��5j      ��5e      �� *Fields+��5k      �� 5l      .R�5m      �� $��Stretch=�6!   
�������$	��0.5q�5o       J� SnapsToDevicePixels . |�.  ._�5r      � $��25��6#    5s   "   �� 5t   &   $��* �67   �� 5u   &   $
��Auto �67   .+�5x   "   �� $��0��6,   $��CharacterEllipsis=�5y   ,   $��3,0,0,0q�5z   ,   $��Center=�5|   ,    � ToolTip 5{   ,   �� *	FieldName+��5}   ,   �� *	FieldName+�� 5~   "    ��Name$ DeleteFieldNameButton��5�   )   -   5        $��Delete��5~   )   $��1��5   )   y�5�   )   �� *	CanDelete+ �� 	Converter# [ 5�   )   5�      5�      5�      5�      � 5�      $��2��6   {�5�      �� *IsPreviewEnabled+ 5�      �� 5�      $��* �6'   �� 5�      $
��Auto �6'   .+�5�      �� -   5         ��Name$ NewFieldNameInputBox��5�      $��0��5�      5�      �� 5�      -   5        $ AddNewFieldNameButton��5�      AddNewFieldNameButton��$ AddNewFieldNameButton��5�      $ AddNewFieldNameButton��5�      $��1��5�      $��3,0,0,0q�5�      .��5�      �� $��_Add field name��6!   .��68   ��5�      �� 5�      " ��6$   	 5�   "   P� 5�   &   
 � {� .  . �5�   *   ��  �@��
Foreground
 �����5�   5   5�   '   5�      .��5�   ,   5�      5�      5�      5�      5�      5�      