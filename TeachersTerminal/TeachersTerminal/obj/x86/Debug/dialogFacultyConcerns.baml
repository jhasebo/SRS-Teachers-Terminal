   M S B A M L   `   `   `  ����U  QMahApps.Metro, Version=0.11.0.0, Culture=neutral, PublicKeyToken=f4fb5a3c4d1e5b4fV;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.MetroMahApps.Metro.Controls  5       TeachersTerminal,   &TeachersTerminal.dialogFacultyConcerns   5      HControls;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro  5      \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      @d2http://schemas.microsoft.com/expression/blend/2008     5   3   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     6p   -   5        $��dialogFacultyConcerns��5   Q    � 
ResizeMode$  NoResize=�5      (   "MahApps.Metro.Controls.MetroWindow   ShowMinButton .     ShowMaxRestoreButton .   � SizeToContent$ WidthAndHeight=�5   �    � Title$ Faculty Concerns��6�    � Icon$7 1/TeachersTerminal;component/Images/books-icon.png��6  .��5      � $	��265��6   $	��536��6    � RowDefinitions 5	   
   �� 5
      $	��66* �6   �� 5      $
��188* �6   .+�5   
   �� -   5        $��
tbConcerns��5   M   $	��199��6   $
��Left=�6   $��	12,12,0,0q�6:   $	��Top=�6_   $	��505��6w   *   $MahApps.Metro.Controls.TextboxHelper   	Watermark$" Up to 200 Characters only...��6�   $��2��6�    �� TextWrapping$
 Wrap=�6�   �� 5   
   -   5        $��	btnSubmit��5   ^   $��Submit��6   $��23��6"   $
��Left=�6.   $��361,151,0,0q�6I   $	��Top=�6o   $��75��6�   $��1��6�   �� 5   
   -   5        $��	btnCancel��5   ^   $��Cancel��6   $��23��6"   $
��Left=�6.   $��442,151,0,0q�6I   $	��Top=�6o   $��75��6�   $��1��6�   �� 5   
   -   5        $��Label1��5   h   $��Character Counter:��6   $��28��6-   $
��Left=�69   $��
12,151,0,0q�6T   $	��Top=�6v   $��1��6�   �� 5   
   -   5        $��lblNumChars��5   Z   $	��200��6   $��28��6   $
��Left=�6*   $��127,151,0,0q�6E   $	��Top=�6m   $��1��6�   5      5      