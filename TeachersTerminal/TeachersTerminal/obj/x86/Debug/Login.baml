   M S B A M L   `   `   `  ����U  QMahApps.Metro, Version=0.11.0.0, Culture=neutral, PublicKeyToken=f4fb5a3c4d1e5b4fV;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.MetroMahApps.Metro.Controls  5       TeachersTerminal   TeachersTerminal.Login   5      HControls;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro  5      \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      @d2http://schemas.microsoft.com/expression/blend/2008     5   p   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     6�   -   5        $��Login��5   R     � 
ResizeMode$  NoResize=�5      (   "MahApps.Metro.Controls.MetroWindow   ShowMinButton .     ShowMaxRestoreButton .   � Title$ Login��5       � UseLayoutRounding .   � WindowStartupLocation$ CenterScreen=�6-   $	��188��6�   $	��293��6   � Icon$@ :/TeachersTerminal;component/Images/Bookmarks-Red-Metro.ico��6     WindowCommands 5	      +   %MahApps.Metro.Controls.WindowCommands  5
   
   5      .��5      0   *MahApps.Metro.Controls.MetroContentControl  .��5      � $	��268��6   $	��149��6    � ColumnDefinitions 5      �� 5      $
��255* �6#   �� 5      $��3* �6#   .+�5      �� -   5        $��tbEmployeeNum��5   5   $��23��6   $��	25,34,0,0q�6"   $	��Top=�6J   $�� ��6b   
���   �*   $MahApps.Metro.Controls.TextboxHelper	   	Watermark$	 Employee Number��6}   $
��Left=�6�   $��140.977��6�   $��0��6�   G� 5   
   -   5        $��
tbPassword��5   P   $��23��6   $
��Left=�6"   $��	26,68,0,0q�6=   $	��Top=�6b   $	��140��6z   
���   �$	 Password��6�   $��1��6�   �� 5      -   5        $��btnLogin��5   `   $��Login��6   $��37��6%   $
��Left=�61   $��
191,41,0,0q�6L   $	��Top=�6p   $��64��6�   $��2��6�   �� 5   
   -   5        $��	lblForgot��5   f   $��Forgot Password?��6   $��28��6+   $
��Left=�67   $��
26,100,0,0q�6R   $	��Top=�6w   $��104.993��6�   
����  �|�.  6�   5      5      5      