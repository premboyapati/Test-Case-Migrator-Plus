   M S B A M L   `   `   `  ����[  WTestCaseMigratorPlus, Version=1.0.0.0, Culture=neutral, PublicKeyToken=9d26ba8cb6a39082u?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator1Microsoft.VisualStudio.TestTools.WorkItemMigrator  5   	   L   FMicrosoft.VisualStudio.TestTools.WorkItemMigrator.WorkItemImportWizard   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   \WorkItemMigrator?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator     5   	   -   5           WizardBackgroundColor#���   5   	     � 
ResizeMode$  CanMinimize=�5   	   $	��600��5   	   $	��750��5   	    � Title 5   	   �� *Title+ � 	Resources 5      S  MMicrosoft.VisualStudio.TestTools.WorkItemMigrator.WizardPartToVisualConverter  WizardPartToVisualConverterS   MMicrosoft.VisualStudio.TestTools.WorkItemMigrator.StringToVisibilityConverter  StringToVisibilityConverterT   NMicrosoft.VisualStudio.TestTools.WorkItemMigrator.BooleanToVisibilityConverter   BooleanToVisibilityConverter[   UMicrosoft.VisualStudio.TestTools.WorkItemMigrator.InverseBooleanToVisibilityConverter ' #InverseBooleanToVisibilityConverterO  IMicrosoft.VisualStudio.TestTools.WorkItemMigrator.InverseBooleanConverter  InverseBooleanConverter  NavigationButtonStyle �� 
TargetType �� Property �� Value �� CornerRadius ) %NavigationPanelControlBackgroundColor  WizardBorderBrush �� Padding 	 WizardBorderGradientBrush � ToolTip	 �� Triggers
 �� Binding �� Value �� Triggers �� 
TargetName 
 WizardPageBackgroundColor ^� 
Conditions �� Binding �� RelativeSource �� Path �� Value % !WizardNavigationPanelItemTemplate ��Name��%{  &	       &	      &	 &     &	 9     &	 L     &	 _     &	 �    7    5   
   67     5   
   67     5   
   68     5   
   6?     5   
   63   �� 5   
   " ��5      .�5      ��  � ��$ 7q�6(   �� 5       � ��$ Stretch=�64   �� 5       � �� 5      �� 5      .R�5      �� $��Border��6!   $ 0��5    !   $��0q�5!   !   $��0,0,1,0q�5"   !   #���  5#   !   #���  5$   !   ._�5%      �� $��MouseOverBorder��5'   %   
������ $��0,0.6,0,0.6q�5&   %   ._�5(   "   �� $��InnerBorder��5*   )   $��0,0.1,0,0.1q�5(   )   $ 3q�5)   )   #��� 	 5+   )   ._�5,   &   �� $��13(�5-   0   $��CharacterEllipsis=�5.   0   $��8,4,2,2q�50   0   ��5,   0   �� *Header+ 5/   0   �� *Header+��51   *   �� 52   .   " ��64   .�53   2   ��  � ��$ Normal"�6O   �� 54   2    � ��
 �   �6O   	 55   2   �� 56   6   
 6B   �� *	CanShow+	 False57   B   .a�58   :   ��  � �� �	#FFC0C0C06W   59   7   5;   /   5=   '   5>   #   5?      5@       5A      �� 5B      $
 True��5C   )   
 5B   *   �� *IsActiveWizardPart+.a�5D   "   �� $ Border��6)   # � 
 5F   )    � ���� 5G   "   $ Border��6)    � ��$ 2,1,0,1q�5I   )   �� 5J   "   $ InnerBorder��6)    � ��$ 0q�5L   )   �� 5M   "   $ Border��6)    �@��CornerRadius$ 3,3,0,3��5O   )   �� 5P   "   $ Border��6)    � ��$ -3,0,0,0q�5R   )   5S      ^� 5U       5V   "   �� 5W   &    60   ��  �� *Self+$ IsMouseOver� True5X   0   �� 5Y   &    60   �� *IsActiveWizardPart+	 False5Z   0   .0�5\   "   �� $ MouseOverBorder��6)   # �  5^   )    � ��5_      5a      5c      5d      �� 5f   
   .R�5g      �� $ WizardNavigationButton��6   -   5        8��  {�5i      �� *	CanShow+5j      5k      5m      .��5o      � $��5q�6    � RowDefinitions 5p   
   �� 5q      $��* �6   �� 5r      $
��Auto �6   .+�5u   
   �� $��0��6   ._�5v      �  � ColumnDefinitions 5w      �� 5x      $
��Auto �6'   �� 5y      $��* �6'   .+�5|      �� $��0��6   $��2,1,0,0q�5}        WizardPageOuterBorderBrush#���  5~      $��0,4,0,0q�5      ._�5�      �  5�      �� 5�      $
��Auto �6,   �� 5�      $
��Auto �6,   �� 5�      $��* �6,   .+�5�      �� $��0��6'   #��[ 5�   +   ��5�   +   �� *WizardParts+5�   +   � 5�      $��1��6    5�      �� 5�   "   $
��Auto �60   �� 5�   "   $
��Auto �60   .+�5�      �� $��0��6%   #���  5�   )   $��0,0,1,0q�5�   )   ._�5�   "   � .+�5�   &   �  � Points; �    	= @KL �� ��[ @t*����������#���  5�   2   � 5�   &   O �   @t*�����9�2
��H u��H !��H ��� ������#���  5�   2   5�   #   5�      � 5�      $��1��6&   = �   ��H��H��-  ;X 	=  ;X@KL  D�@KL  D�#���  5�   *   5�      5�      5�      �� 5�      $��1��6   $��0,1,0,0q�5�      #���  5�      $��0,4,4,4q�5�      ._�5�      �� #���  6   $��0,1,1,1q�5�   !   #��� 
 5�   !   ._�5�      �  5�      �� 5�   "   $
��Auto �60   �� 5�   "   $
��Auto �60   �� 5�   "   $��* �60   �� 5�   "   $
��Auto �60   .+�5�      �� $��0��6)   $	��Top=�5�   -   $��7,8,10,0q�5�   -   .+�5�   "   �� $
��Bold"�5�   0   $��16(�5�   0   ��5�   ,   �� *WizardPart.Header+ �� Mode$ OneWay=�5�   0   �� 5�   "   $��10,5,0,5q�5�   0   $
��Wrap=�5�   0   ��5�   ,   �� *WizardPart.Description+y�5�   0   �� *WizardPart.Description+ �� 	Converter# [ 5�   0   5�      �� 5�      $��1��6(   $	��Top=�5�   ,   $��7,7,7,0q�5�   ,   �� 5�      $��2��6/   $��	11,5,11,5q�5�   3   ��5�   3   �� *
WizardPart+# [ �� 5�      $��3��6(   $��24��5�   ,   $��0,1,0,0q�5�   ,     StatusBorderColor#���  5�   ,     StatusBackgroundColor#���  5�   ,   .@�5�   "   � $��Center=�5�   +   y�5�   '   �� *WizardPart.Warning+# [  5�   &   �� 5�   *   $
��Auto �6;   �� 5�   *   $��* �6;   .+�5�   &   �� $��16��6,   $��16��5�   0   $��0��5�   0   $��Center=�5�   0   $��3,0,3,0q�5�   0     
Warning_16#���  5�   0   �� 5�   &   $
��Wrap=�5�   4   $��Center=�5�   4   $��1��5�   4   ��5�   0   �� *WizardPart.Warning+$ TwoWay=�5�   4   5�   #   5�      5�      5�      5�      5�      5�      � 5�   
   $��1��6   $��Bottom=�5�      $��7,14,7,7q�5�       5�      �� 5�      $
��Auto �6#   �� 5�      $��* �6#   .+�5�      �� -   5        $��
HelpButton��5�      $��0��5�      
HelpButton��  NavigationButton#���  5�      $
��Help��5�      5�      �� 5�      $��1��6   $��Right=�5�       �� Orientation$ 
Horizontal=�5�      .+�5�      �� -   5        $��
BackButton��5�      
BackButton��$��< _Back��5�      #���  5�      {�5�      �� *WizardPart.CanBack+y�5�      �� *IsSummaryPage+# [ 5�      �� 5�      -   5        $��
NextButton��5�      
NextButton��$��_Next >��5�      #���  5      y�5     �� *IsConfirmVisible+# [ {�5     �� *WizardPart.CanNext+5     �� 5     -   5        $��ConfirmButton��5     ConfirmButton��#���  5     $��_Confirm >>��5     y�5	     �� *IsConfirmVisible+# [ {�5
     �� *WizardInfo.CanConfirm+5     �� 5     -   5        $��SaveandExitButton��5     SaveandExitButton��#���  5     $��_Save and exit��5     {�5     �� *	CanSave+y�5     �� *IsConfirmPage+# [ 5     �� 5     -   5        $��SaveAndMigrateButton��5     SaveAndMigrateButton��#���  5     $��Save and _migrate��5     y�5     �� *IsConfirmPage+# [ 5     �� 5     -	   5        $��StartAgainButton��5     StartAgainButton��#���  5     $��S_tart again��5     {�5      �� *CanStartAgain+y�5!     �� *IsSummaryPage+# [ 5"     �� 5$     -
   5        $��CloseButton��5$     CloseButton��#���  5&     $��_Close��5'     y�5(     �� *IsSummaryPage+# [ 5)     �� 5+     -   5        $��CancelButton��5+     CancelButton��#���  5-     $��_Cancel��5/     y�5.     �� *IsSummaryPage+# [ 50     51     52     53     54     