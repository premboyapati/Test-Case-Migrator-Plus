   M S B A M L   `   `   `  ���� [  WTestCaseMigratorPlus, Version=1.0.0.0, Culture=neutral, PublicKeyToken=9d26ba8cb6a39082u?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator1Microsoft.VisualStudio.TestTools.WorkItemMigrator  L   FMicrosoft.VisualStudio.TestTools.WorkItemMigrator.WorkItemImportWizard   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     \WorkItemMigrator?clr-namespace:Microsoft.VisualStudio.TestTools.WorkItemMigrator     -      WizardBackgroundColor#���     � 
ResizeMode$  CanMinimize=�$	��600��$	��750�� � Title �� *Title+ � 	Resources S  MMicrosoft.VisualStudio.TestTools.WorkItemMigrator.WizardPartToVisualConverter  WizardPartToVisualConverterS   MMicrosoft.VisualStudio.TestTools.WorkItemMigrator.StringToVisibilityConverter  StringToVisibilityConverterT   NMicrosoft.VisualStudio.TestTools.WorkItemMigrator.BooleanToVisibilityConverter   BooleanToVisibilityConverter[   UMicrosoft.VisualStudio.TestTools.WorkItemMigrator.InverseBooleanToVisibilityConverter ' #InverseBooleanToVisibilityConverterO  IMicrosoft.VisualStudio.TestTools.WorkItemMigrator.InverseBooleanConverter  InverseBooleanConverter  NavigationButtonStyle �� 
TargetType �� Property �� Value �� CornerRadius ) %NavigationPanelControlBackgroundColor  WizardBorderBrush �� Padding 	 WizardBorderGradientBrush � ToolTip	 �� Triggers
 �� Binding �� Value �� Triggers �� 
TargetName 
 WizardPageBackgroundColor ^� 
Conditions �� Binding �� RelativeSource �� Path �� Value % !WizardNavigationPanelItemTemplate ��Name��%F  &	       &	      &	 
     &	      &	      &	      &	 �    7            �� " ��.���  � ��$ 7q���  � ��$ Stretch=���  � �� �� .R��� $��Border��$ 0��$��0q�$��0,0,1,0q�#���  #���  ._��� $��MouseOverBorder��
������ $��0,0.6,0,0.6q�._��� $��InnerBorder��$��0,0.1,0,0.1q�$ 3q�#��� 	 ._��� $��13(�$��CharacterEllipsis=�$��8,4,2,2q����� *Header+ �� *Header+���� " ��.���  � ��$ Normal"���  � ��
 �   �	 �� 
 �� *	CanShow+	 False.a���  � �� �	#FFC0C0C0 �� $
 True��
 �� *IsActiveWizardPart+.a��� $ Border��# � 
  � ���� $ Border�� � ��$ 2,1,0,1q��� $ InnerBorder�� � ��$ 0q��� $ Border�� �@��CornerRadius$ 3,3,0,3���� $ Border�� � ��$ -3,0,0,0q�^�  ��  ��  �� *Self+$ IsMouseOver� True��  �� *IsActiveWizardPart+	 False.0��� $ MouseOverBorder��# �   � ���� .R��� $ WizardNavigationButton��-   8��  {��� *	CanShow+.��� $��5q� � RowDefinitions �� $��* ��� $
��Auto �.+��� $��0��._��  � ColumnDefinitions �� $
��Auto ��� $��* �.+��� $��0��$��2,1,0,0q�  WizardPageOuterBorderBrush#���  $��0,4,0,0q�._��  �� $
��Auto ��� $
��Auto ��� $��* �.+��� $��0��#��[ ���� *WizardParts+� $��1�� �� $
��Auto ��� $
��Auto �.+��� $��0��#���  $��0,0,1,0q�._�� .+��  � Points; �    	= @KL �� ��[ @t*����������#���  � O �   @t*�����9�2
��H u��H !��H ��� ������#���  � $��1��= �   ��H��H��-  ;X 	=  ;X@KL  D�@KL  D�#���  �� $��1��$��0,1,0,0q�#���  $��0,4,4,4q�._��� #���  $��0,1,1,1q�#��� 
 ._��  �� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto �.+��� $��0��$	��Top=�$��7,8,10,0q�.+��� $
��Bold"�$��16(����� *WizardPart.Header+ �� Mode$ OneWay=��� $��10,5,0,5q�$
��Wrap=����� *WizardPart.Description+y��� *WizardPart.Description+ �� 	Converter# [ �� $��1��$	��Top=�$��7,7,7,0q��� $��2��$��	11,5,11,5q����� *
WizardPart+# [ �� $��3��$��24��$��0,1,0,0q�  StatusBorderColor#���    StatusBackgroundColor#���  .@�� $��Center=�y��� *WizardPart.Warning+# [  �� $
��Auto ��� $��* �.+��� $��16��$��16��$��0��$��Center=�$��3,0,3,0q�  
Warning_16#���  �� $
��Wrap=�$��Center=�$��1������ *WizardPart.Warning+$ TwoWay=�� $��1��$��Bottom=�$��7,14,7,7q� �� $
��Auto ��� $��* �.+��� -   $��
HelpButton��$��0��
HelpButton��  NavigationButton#���  $
��Help���� $��1��$��Right=� �� Orientation$ 
Horizontal=�.+��� -   $��
BackButton��
BackButton��$��< _Back��#���  {��� *WizardPart.CanBack+y��� *IsSummaryPage+# [ �� -   $��
NextButton��
NextButton��$��_Next >��#���  y��� *IsConfirmVisible+# [ {��� *WizardPart.CanNext+�� -   $��ConfirmButton��ConfirmButton��#���  $��_Confirm >>��y��� *IsConfirmVisible+# [ {��� *WizardInfo.CanConfirm+�� -   $��SaveandExitButton��SaveandExitButton��#���  $��_Save and exit��{��� *	CanSave+y��� *IsConfirmPage+# [ �� -   $��SaveAndMigrateButton��SaveAndMigrateButton��#���  $��Save and _migrate��y��� *IsConfirmPage+# [ �� -	   $��StartAgainButton��StartAgainButton��#���  $��S_tart again��{��� *CanStartAgain+y��� *IsSummaryPage+# [ �� -
   $��CloseButton��CloseButton��#���  $��_Close��y��� *IsSummaryPage+# [ �� -   $��CancelButton��CancelButton��#���  $��_Cancel��y��� *IsSummaryPage+# [ 