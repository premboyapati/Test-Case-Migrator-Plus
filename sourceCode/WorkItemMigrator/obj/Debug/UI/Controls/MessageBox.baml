   M S B A M L   `   `   `  ����[  WTestCaseMigratorPlus, Version=1.0.0.0, Culture=neutral, PublicKeyToken=9d26ba8cb6a39082u?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator1Microsoft.VisualStudio.TestTools.WorkItemMigrator  5   	   B   <Microsoft.VisualStudio.TestTools.WorkItemMigrator.MessageBox   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   \WorkItemMigrator?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator     5   	     � 
ResizeMode$  NoResize=�5   	    � Title$ Test Case Migrator Plus��5   	    � WindowStartupLocation$ CenterOwner=�5   	    � SizeToContent$ WidthAndHeight=�5	   	    � ShowInTaskbar .  $	��350��5   	   $	��600��5   	   $	��400��5   	    � 	Resources 5      S   MMicrosoft.VisualStudio.TestTools.WorkItemMigrator.StringToVisibilityConverter   StringToVisibilityConverter  MessageIcon �� 
TargetType �� Property �� Value  ErrorIcon_32	 �� TriggersG  AMicrosoft.VisualStudio.TestTools.WorkItemMigrator.MessageCategory
   Warning �� Value �� Binding  WarningIcon_32   Information  InfoIcon_32  MessageDialogTitleStyle �� BasedOn  MessageDialogTextStyle��%�  &	        &	      &	 �    &	 �      5   
   67   �� 5   
   " ��5      .�5      ��  � ��$ 32��5      �� 5       � ��$ 32��5      �� 5       � ��$ 5,10,5,0q�5      �� 5       � ��$	 Top=�5      �� 5       � ��# �  5      	 5      �� 5      # Z
 5       5      �� *
Category+.a�5       ��  � ��# �  6/   5!      �� 5"      # Z 5#       5"      �� *
Category+.a�5$      ��  � ��# �  6/   5%      5'      �� 5*   
   " ��5+       5,      Z� .�5.      ��  � ��$ 0q�5/      �� 50       � |� . 51      �� 52       � ��$ 16(�53      �� 54       � ��
 ���� 55      �� 56       � ��
 �   �57      �� 58       � ��$
 Left=�59      �� 5:       � ��$ 5,10,5,5q�5;      �� 5<       �@��TextWrapping$
 Wrap=�5=      �� 5>       � ��$ Stretch=�5?      �� 5@       � ��$
 Auto��5A      5C      �� 5E   
   " ��5F       5G      Z� .�5H      ��  � ��$ 0q�5I      �� 5J       � |� . 5K      �� 5L       � ��$ 12(�5M      �� 5N       � ��
 ���� 5O      �� 5P       � ��
 �   �5Q      �� 5R       � ��$ 5,5,5,10q�5S      �� 5T       �@��TextWrapping$
 Wrap=�5U      �� 5V       � ��$
 Auto��5W      5X      5Y      .��5[      �  � RowDefinitions 5\   
   �� 5]      $��* �6   �� 5^      $
��Auto �6   .+�5a   
   � $��0��6   $��Stretch=�5b      $��Stretch=�5c       � ColumnDefinitions 5d      �� 5e      $��45 �6#   �� 5f       5i      �� 5j      $
��Auto �6    �� 5k      $
��Auto �6    �� 5l      $
��Auto �6    .+�5o      �� -   5         ��Name$ MessagePanelIcon��5p      #���  5o      |�.  $��0��5r      $��3��5s      �� 5t      -   5         ��Name$ ProblemTextBlock��5t      4  .System.Windows.Automation.AutomationProperties   AutomationId$! MessageBox.ProblemTextBlock��5u      $��1��5v      $��0��5w      $	��350��5x       �� 
IsReadOnly . #���  5z      ��5{      �� *Title+y�5|      ��  �� Path$ ProblemStatement� �� 	Converter# [  �� 5~      -   5        $ CauseTextBlock��5~      $ MessageBox.CauseTextBlock��5      $��1��5�      $��1��5�       . ��.  " �� VerticalScrollBarVisibility$
 Auto=�5�      #���  5�      ��5�      �� *LikelyCause+y�5�      �� $ ProblemCause�# [  .��6s   5�      �� 5�      -   5        $ SolutionTextBlock��5�      $" MessageBox.SolutionTextBlock��5�      $��1��5�      $��2��5�       . ��.  #���  5�      ��5�      �� *PotentialSolution+y�5�      �� $ ProblemSolution�# [  .��6v   5�      5�      � 5�   
   -   5         �Name$ ButtonBorder��5�      $��1��5�      
�������$��40��5�      .+�5�      �� -   5         ��Name$ ButtonPanel��5�      $ MessageBox.ButtonPanel��5�      $��Right=�5�      $��3,5,3,5q�5�       �� Orientation$ 
Horizontal=�5�      .+�5�      �� -   5         ��Name$ FirstButton��5�      MessageBox_FirstButton��$ MessageBox.FirstButton��5�        MessageBox_Button#���  5�      ��5�      �� *FirstButtonLabel+y�5�      �� *FirstButtonLabel+# [  �� 5�      -   5        $ SecondButton��5�       UserHostControl_SecondButton��$ MessageBox.SecondButton��5�      #���  5�      ��5�      �� *SecondButtonLabel+y�5�      �� *SecondButtonLabel+# [  �� 5�      -	   5        $ CloseButton��5�      UserHostControl_CloseButton��$ MessageBox.CloseButton��5�      #���  5�       �� IsCancel . ��5�      �� $ CancelButtonLabel�5�      5�      5�      5�      