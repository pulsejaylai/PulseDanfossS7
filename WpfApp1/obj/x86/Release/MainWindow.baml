   M S B A M L   `   `   `  ���� B  >WpfApp1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null!clr-namespace:WpfApp1WpfApp1      WpfApp1.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     'localclr-namespace:WpfApp1     -     � Title$  DanfossTest��$	��600��$
��1500�� � 
ResizeMode$ CanMinimize=�.��� $
��Left=�$	��590��$��
58,0,0,-18q�$	��Top=�$
��1363�� J� RenderTransformOrigin$ 0.532,0.589&� � RowDefinitions �� $��5* ��� $��4* �.+�o� -   $��	menustrip��$
��Left=�$��48��$��	10,22,0,0q�$	��Top=�$	��412��$��0q�
�������
���   �.@�m� $��文件��$��11(�$��45��
�������.@�m� -    m�Name$ open1��$��打开��$��11(�
�������
���   �{�. m� -   $ equset��$��设备地址��$��11(�$��80��
�������m� -   $ reset��$��电机复位��$��11(�$��90��
�������{�.  m� -   $ 
DataExpert��$��数据导出��$��11(�$��90��
�������{�.  m� -   $ 	StopMotor��$��停止电机��$��11(�$��90��
�������{�.  &   System.Windows.Controls.DataGrid  -     Name$ dataGridView1��   CanUserAddRows .  $
��Left=�$	��180��$��149,125,0,0q�$��2��$	��Top=�$	��575��$��Visible=�$��Visible=�   AutoGenerateColumns .  ����    Columns 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn	   Header$		 SEQ��
   Width5  /System.Windows.Controls.DataGridLengthConverter$
 35 1  +System.Windows.Controls.DataGridBoundColumn   Binding �� *seq+  $
	 Item��$	
 150  �� *item+4  .System.Windows.Controls.DataGridTemplateColumn  $
 40    HeaderTemplate �� .R��� $	��NC1��   CellTemplate �� .R��� $��Center=�$��Center=����� *nc1+���� *	nc1FColor+  $
 40  �� .R��� $	��NO2�� �� .R��� $��Center=�$��Center=����� *no2+���� *	no2FColor+  $
 40  �� .R��� $	��NC3�� �� .R��� $��Center=�$��Center=����� *nc3+���� *	nc3FColor+  $
 40  �� .R��� $	��NO4�� �� .R��� $��Center=�$��Center=����� *no4+���� *	no4FColor+  $
 40  �� .R��� $	��LED�� �� .R��� $��Center=�$��Center=����� *led+���� *	ledFColor+  $
 40  �� .R��� $
��DCR1�� �� .R��� $��Center=�$��Center=����� *dcr1+���� *
dcr1FColor+  $
 40  �� .R��� $
��DCR2�� �� .R��� $��Center=�$��Center=����� *dcr2+���� *
dcr2FColor+  $
 40  �� .R��� $��CanData�� �� .R��� $��Center=�$��Center=����� *	candata+���� *candataFColor+  $
 50  �� .R��� $��Result�� �� .R��� $��Center=�$��Center=����� *result+���� *resultFColor+   RowStyle ��  �� 
TargetType)  #System.Windows.Controls.DataGridRow"   �� Triggers P�  P� Property �@J�IsMouseOver P� Value . . ���  �� Property � �� �� Value
 �����P�  �@ 
IsSelected . . ���  � ��
 �������  � ��
 ��  �   WpfApp1.ButtonEx  -	     Name$
 test��	   WpfApp1.ButtonEx 	  Click$ ButtonEx_Click_1��$
��Test��$
��Left=�$��50��$��
83,152,0,0q�$��1��$	��Top=�$	��100��   Icon$ D:/ico/button3.ico��
���  ��   
ButtonType$ IconText=�  -
   $ pause3��$ Pause_Click��$��Pause��$
��Left=�$��50��$��251,152,0,0q�$��1��$	��Top=�$	��100��$ D:/ico/button1.ico��
���  ��$ IconText=�  -   $ si��$ Si_Click��$��Single��$
��Left=�$��50��$��427,152,0,0q�$��1��$	��Top=�$	��100��$ D:/ico/button2.ico��
���  ��$ IconText=��� -    ��Name$ 	CheckBox1��$��CheckBox��$
��Left=�$��50��$��727,131.333,0,0q�$��1��$	��Top=�$	��125���� -    ��Name$ ErrBox��$
��Left=�$	��144��$��
805,55,0,0q�$	��Top=�$	��180���� -    ��Name$ MODEL��$��Model��$
��Left=�$��50��$��
390,20,0,0q�$	��Top=�$	��500��
����  �$��35(�$��Normal$�$��DemiBold"�$��./#仿宋_GB2312)��� -   $ Statue��$��Ready��$
��Left=�$��50��$��863,255,0,0q�$	��Top=�$	��500��
����  �$��35(�$��Normal$�$��DemiBold"�$��./#仿宋_GB2312)��� -   $��snbox2��$
��Left=�$��26��$��
260,77,0,0q� �� TextWrapping$
 Wrap=�$�� ��$	��Top=�$	��399�� �� 
IsReadOnly .   -   $
 STOP��$ 
STOP_Click��$
��Stop��$
��Left=�$��50��$��562,152,0,0q�$��1��$	��Top=�$	��100��$ D:/ico/button4.ico��
���  ��$ IconText=��� -   $��list1��$
��Left=�$	��144��$��160,0,0,122q�$��Bottom=�$	��491��$ -0.277,0.629&�$��2���� -   $
 Time��$
��time��$
��Left=�$��50��$��
863,15,0,0q�$	��Top=�$	��500��
����  �$��25(�$��Normal$�$��DemiBold"�$��./#仿宋_GB2312)�$��1��