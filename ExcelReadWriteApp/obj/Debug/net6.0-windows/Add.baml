   M S B A M L   `   `   `  ????L  HExcelReadWriteApp, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null5clr-namespace:ExcelReadWriteAppExcelReadWriteApp  5   	       ExcelReadWriteApp.Add   5   	   W SPresentationCore, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   1localclr-namespace:ExcelReadWriteApp     5   	     ? Title$	  Add??5   	   $	??450??5	   	   $	??450??5
   	   ???	#FFDCE9F5 ? 	Resources 5       ?? 
TargetType ?? Triggers P? Property P? Value ?? Property ?? Value ?? Self	 ?? RelativeSource
 ?? Path??%?   '??       ?? 5      " ??6    5      P? 5   
    ?@=?HasError$
 True??5      . ?5      ??  ?@?ToolTip 5      ?? #	 Z $&
  (Validation.Errors)/ErrorContent?5      5      5      .??5      ? .+?5      ?? $??Employee ID:??6   $
??Left=?5      $??	71,45,0,0q?5      $	??Top=?5      $??14(?5      ?? 5      $??Name:??6   $
??Left=?5      $??
71,106,0,0q?5       $	??Top=?5!      $??83??5"      $??14(?5#      ?? 5$      $??Salary:??6   $
??Left=?5%      $??
71,165,0,0q?5&      $	??Top=?5'      $??14(?5(      $??83??5)      ?? 5*      $??Department:??6   $
??Left=?5+      $??
67,218,0,0q?5,      $	??Top=?5-      $??14(?5.      ?? 5/      $
??Left=?6   $??
216,45,0,0q?50       ?? TextWrapping$
 Wrap=?51      $	??Top=?52      $	??209??53      $??14(?54      $??32??55      ??56      ?? *Employee.EmployeeId+?? 57      $
??Left=?6   $??216,106,0,0q?58      $
 Wrap=?59      $	??Top=?5:      $	??209??5;      $??14(?5<      $??32??5=      ??5>      ?? *Employee.EmployeeName+ ?? UpdateSourceTrigger$ PropertyChanged=? ?? ValidatesOnDataErrors$
 True???? 5A      $
??Left=?6   $??216,165,0,0q?5B      $
 Wrap=?5C      $	??Top=?5D      $	??209??5E      $??29??5F      ??5G      ?? *Employee.Salary+?? 5H      $
??Left=?6   $??216,218,0,0q?5I      $
 Wrap=?5J      $	??Top=?5K      $	??209??5L      $??29??5M      ??5N      ?? *Employee.Department+?? 5O      $	??Add??6   $??Center=?5P      $??	0,294,0,0q?5Q      $	??Top=?5R      $??28??5S      $??77??5T      ???	#FF92A5B1??5U      ?? *	SaveCmd+5V      ?? 5W      $
??Left=?6   $??116,353,0,0q?5X      $
??Wrap=?5Y      $	??Top=?5Z      $??25??5[      $	??225??5\      ???	#FF71DE22$??Heavy"?5_      $??16(?5`      ??5]      ?? *msg+5`      5b      5c      