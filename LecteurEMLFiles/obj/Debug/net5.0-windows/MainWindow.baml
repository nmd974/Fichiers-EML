   M S B A M L   `   `   `  ����J  FLecteurEMLFiles, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:LecteurEMLFilesLecteurEMLFiles  5   	        LecteurEMLFiles.MainWindow   5   	   W SPresentationCore, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   /localclr-namespace:LecteurEMLFiles     5   	     � Title$  
MainWindow��5   	   $	��550��6   $	��800��6)   .��5	      �� .+�5
   
   � $��40, 0q�6   $��50��6   $	��700��6*    � ColumnDefinitions 5      �� 5      $
��700* �6#   62   .+�5      �� $(��"Sélectionnez le fichier à lire :��6   $��Bottom=�6?   5      � 5   
   $��50��6   $	��700��6    5      �� 5      $
��600* �6#   62   �� 5      $
��100* �6#   62   .+�5      �� -   5         ��Name$ 	File_list��5   '   $��0��6   6'   �� 5      -   5         ��Name$ LoadEml��5   4   $��1��6   $
��Lire��6%   6E   5      � 5   
   $��0,10,0,0q�6   $	��350��6!   $	��700��6.    � RowDefinitions 5      �� 5      $
��350* �6    60   .+�5      &   System.Windows.Controls.DataGrid  -   5        $��Headers��5"         AutoGenerateColumns$ False��5         CanUserAddRows$ False��5      $	��350��5       $
��Left=�5!      $	��Top=�5#      $	��700��5$      $��0��5%         Columns 5&      0  *System.Windows.Controls.DataGridTextColumn  5'      ,  &System.Windows.Controls.DataGridColumn   MinWidth$	 250Q�6D   1  +System.Windows.Controls.DataGridBoundColumn	   Binding	 6)   �� 
 �� Path$
 A�   ElementStyle 5(      �� 5)       �� 
TargetType" ��6$   .�5*   "   ��  �� Property � �� �� Value 6I   �� $
 ColorA��� 5+   "    � �� 6I   �� $
 ForegroundA�5,      5.        5/      $	 450Q�6D   	 6)   �� $
 B� 50      �� 51      " ��6$   .�52   "   ��  � �� 6I   �� $
 ColorB��� 53   "    � �� 6I   �� $
 ForegroundB�54      56      58      5:      5;      5=      