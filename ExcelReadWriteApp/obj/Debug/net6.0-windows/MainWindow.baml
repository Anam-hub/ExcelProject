   M S B A M L   `   `   `  ????L  HExcelReadWriteApp, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null5clr-namespace:ExcelReadWriteAppExcelReadWriteApp  5   	   "    ExcelReadWriteApp.MainWindow   5   	   W SPresentationCore, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   1localclr-namespace:ExcelReadWriteApp     5   	     ? Title$  Connect to Excel Sheet??5   	   
???????$	??450??5
   	   $	??800??5   	   .??5      ? .+?5      ?? $
??Left=?6   $??
120,30,0,0q?5       ?? TextWrapping$
 Wrap=?5      $	??Top=?5      $	??424??5      $??40??5      $??16(?5      ??5      ?? *
filepath+?? 5      $??Browse??6   $
??Left=?5      $??
584,30,0,0q?5      $	??Top=?5      $??40??5      $	??142??5      $??18(?5      ???	#FFA3D2EF??5      ?? *	BrowseCmd+?? 5      $??Import??6   $
??Left=?5       $??584,104,0,0q?5!      $	??Top=?5"      $??37??5#      $	??142??5$      ???	#FFA3D2EF$??18(?5&      ??5'      ?? *	ImportCmd+&   System.Windows.Controls.DataGrid  5(      $??43,89,256,55q?5)      ???	#FFB8EEB5??5(      ?? *dataTable.DefaultView+5*      ?? 5+      $??
Add Record??6   $
??Left=?5,      $??584,165,0,0q?5-      $	??Top=?5.      $??38??5/      $	??142??50      $??18(?51      $|?False??53       J? RenderTransformOrigin$ 0.521,1.314&?54      ???	#FF90D6EA??52      ?? *AddCmd+55      ?? 56      $??Export??6   $
??Left=?57      $??584,234,0,0q?58      $	??Top=?59      $??36??5:      $	??142??5;      ???	#FF6F7676???	#FF8BC5DF$??18(?5?      $ 0.483,-0.222&?5@      ??5<      ?? *	ExportCmd+5@      ?? 5A      $??Search??6   $
??Left=?5B      $??	43,34,0,0q?5C      $	??Top=?5D      $??16(?5E      5H      5I      