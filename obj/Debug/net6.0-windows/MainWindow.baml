   M S B A M L   `   `   `  ����G  Cwraplauncher, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null+clr-namespace:WrapLauncherWrapLauncher  5          WrapLauncher.MainWindow   5      W SPresentationCore, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5      R NSystem.Xaml, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      ,localclr-namespace:WrapLauncher     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      -   5          � Title$  スタート��5      $	��800��5	      $	��450��5
      
������� � 
ResizeMode$ CanResizeWithGrip=�5       � 	Resources 5       �� 
TargetType �� Property �� Value �� 
TargetType �� Padding �� CornerRadius	 �� Triggers
 P� Property P� Value   GroupTitleStyle��%�  '��       &	      �� 5   
   " ��6   .�5      ��  � ��$ 5q�6'   �� 5       � �� 5      �� 5      " ��6&   .R�5      �� $	 5,3q�5      #��z��5      
���   �$��1q�5      $ 3��5      ._�5       �� $��Center=�6/   $��Center=�6L   5!      5"      5$      	 5&      P� 5(      
 �@J�IsMouseOver$
 true��61   . �5*      ��  � �� �#bee6fd63   5+      P� 5-      
 � ��$
 true��6/   . �5/      ��  � �� �#88d3fd63   50      52      �� 55   
   " ��6(   .�56      ��  � ��$ 5,10,5,0q�6'   �� 57       � ��$
 Bold"�6+   �� 58       � ��
 ��  �6+   59      5:       � ContextMenu 5<      �� 5>   
   -   5         ��Name$ WindowContextMenu��5>      .@�5?      m� -   5        $��設定再読み込み(_R)��5?   0   m� 5@      -   5        $+��%ランチャーの場所を開く(_O)��5@   4   �� 5B      m� 5D      -   5         m�Name$ MinimizedMenuItem��5E      $=��7アプリを起動したらランチャー最小化(_M)��5F       m� IsCheckable$
 True��5G       m� 	IsChecked$ False��5H       � ToolTip$[ U※Ctrlキーを押しながらアプリを起動した場合は最小化しません��5I      �� 5K      -   5         ��Name$ 	ScaleMenu��5K       �� Orientation$ 
Horizontal=�6,   .+�5L      �� $��0,0,10,0q�6   $��表示倍率��6.   	� 5M      -   5         	�Name$ DefaultScaleItem��5N      $	��5,0q�5O      $	��1.0��5Q       d� 	IsChecked$
 TrueY�5R      	� 5S      -   5        $	��5,0q�5T      $	��1.5��5V      	� 5W      -	   5        $	��5,0q�5X      $	��2.0��5Z      	� 5[      -
   5        $	��5,0q�5\      $	��3.0��5^      5_      �� 5a      m� 5c      -   5        $ InfoMenuItem��5d      $��
情報(_I)��5f      �� 5h      m� 5j      -   5        $��ヘルプ(_H)��5j   .   5k      .��5n      �� $��0,0,0,16q�6   .��5p   
   �� -   5        $ MainContainer��5p      5q      5r      