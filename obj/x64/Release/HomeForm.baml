   M S B A M L   `   `   `  ���� C  ?SCAdvert, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null#clr-namespace:SCAdvertSCAdvert  3clr-namespace:SCAdvert.ClassesSCAdvert.Classes      SCAdvert.HomeForm   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     +scAdvertclr-namespace:SCAdvert     1filterclr-namespace:SCAdvert.Classes     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     -   $��FormHome��  � Title$  SCAdvert��$��566.5��$	��868��$��566.5��$	��868�� � WindowStartupLocation$ CenterScreen=� � WindowState$ 	Maximized=� � 	Resources ��%    .��� .+��� -   $��BtnEnter��$��Вход��$��	0,10,20,0q�$	��Top=�
����������	#FF16ADE6$��Right=�$��75��&   System.Windows.Controls.DataGrid  -   $��DataGridSql��$��10, 50,20,51q�   AutoGenerateColumns .     CanUserResizeColumns .  $   HorizontalScrollBarVisibility$
 Auto=�   CanUserAddRows .  
�������   RowBackground
 �����	   AlternatingRowBackground
	 �����
   Columns
 0  *System.Windows.Controls.DataGridTextColumn    1  +System.Windows.Controls.DataGridBoundColumn   Binding �� *MediaTypeName+,  &System.Windows.Controls.DataGridColumn   Header � -    �Name$ GridMediaType�� � ColumnDefinitions �� �� .+��� $��0��$	��100��$��	MediaType�� �� TextAlignment$ Center=�$��Center=��� -    ��Name$ CboMediaType��$��1��
�������   �� *	Column1+ � -   $ GridYear�� �� �� .+��� $��0��$��50��$
��Year��$ Center=�$��Center=��� -   $ CboYear��$��1��
�������   �� *Month+ �  �� �� .+��� $��0��$��50��$��Month��$ Center=�$��Center=��� -   $��CboMonth��$��1��
�������   �� *
SectorName+ �  �� �� .+��� $��0��$	��100��$��
SectorName��$ Center=�$��Center=�$
��Left=�$	��0,3q�$��16���� -	   $��	CboSector��$��1��
�������$
��Left=�   �� *CategoryName+ �  �� �� .+��� $��0��$	��100��$��CategoryName��$ Center=�$��Center=�$
��Left=�$	��0,3q��� �� -
   $��CboCategory��$��1��
�������   �� *	ClassName+ �  �� �� .+��� $��0��$	��100��$��	ClassName��$ Center=�$��Center=��� -   $��CboClass��$��1��
�������   �� *
Producer+ �  �� �� .+��� $��0��$	��100��$��Producer��$ Center=�$��Center=��� -   $��CboProducer��$��1��
�������   �� *Brand+ �  �� �� .+��� $��0��$	��100��$��Brand��$ Center=�$��Center=��� -   $��CboBrand��$��1��
�������   �� *	Product+ �  �� �� .+��� $��0��$	��100��$��Product��$ Center=�$��Center=��� -   $��
CboProduct��$��1��
�������   �� *Market+ �  �� �� .+��� $��0��$	��100��$��Market��$ Center=�$��Center=��� -   $��	CboMarket��$��20��$��1��
�������   �� *Distributor+ �  �� �� .+��� $��0��$	��100��$��Distributor��$ Center=�$��Center=��� -   $��CboDistributor��$��1��
�������   �� *AdType+ �  �� �� .+��� $��0��$	��100��$��AdType��$ Center=�$��Center=��� -   $��	CboAdType��$��1��
�������   �� *
AdFormat+ �  �� �� .+��� $��0��$	��100��$��AdFormat��$ Center=�$��Center=��� -   $��CboAdFormat��$��1��
��������� -   $��ExportToExcel��$+��%Экспортировать в EXCEL��$��	0,0,20,16q����	#FF16ADE6
�������$��22��$��Bottom=�$��Right=�$	��166���� -   $��btnFirst��$��<<��$
��Left=�$��	10,0,0,16q�$��45��$��22��$��Bottom=����	#FF16ADE6
��������� -   $��btnPrevious��$��<��$
��Left=�$��	60,0,0,16q�$��Bottom=�$��75��$��22�����	#FF16ADE6
��������� -   $��btnNext��$��>��$
��Left=�$��
335,0,0,16q�$��Bottom=�$��75��$��22�����	#FF16ADE6
��������� -   $��btnLast��$��>>��$
��Left=�$��
415,0,0,16q�$��Bottom=�$��45��$��22�����	#FF16ADE6
��������� -   $��	txtPaging��$
��Left=�$��23��{�.   �� TextAlignment$ Center=�$��
140,0,0,15q� �� TextWrapping$
 Wrap=�$�� ��$��Bottom=�$	��190���� -   $��btnClearFilter��$%��Сбросить фильтры��$
��Left=�$��	10,10,0,0q�$	��Top=�$	��125��$��22�����	#FF16ADE6
�������