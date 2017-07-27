# ExcelDna.COM.QuickStart

*Author [@Daniel NGUYEN](https://www.linkedin.com/in/nguyendaniel)*
July 27th 2017
EXCEL DNA COM Sample : Generation of an XLL file to be used in VBA

1. Check post build event of project "DataApiAddin": it will generate tlb files, then build a DataApiAddin-packed.xll.

2. Open Excel, load the DataApiAddin-packed.xll. 
2. Or, Open Excel, Tools > References > Browse in order to add DataApiAddin-packed.xll
Not the tlb but the .xll. DataApiAddin-packed.xll can be added because of the following Pack="true" in DataApiAddin.dna file:

```
<ExternalLibrary Path="DataApiAddin.dll" ComServer="true" Pack="true" />
```

tlb could be added but it is not the purpose.

3. Alt+F11 
Add a module and the following code:


```
Public Sub Test()
    Dim api As New COMDataManager
    Dim res As String    
    res = api.GetHello("2017")
    MsgBox (res)
End Sub
```


' These work only after the COM Server has been registered
' either by calling ExcelDna.ComInterop.ComServer.DllRegisterServer() 
' (like from a macro or AutoOpen routine in the Excel-DNA addin)
' or by calling "regsvr32 DataApiAddin-packed.xll" from a command prompt (once, before opening Excel).

```
Check ExcelAddin.cs class : use of AutoOpen routine
```

4. Press F5 to check the result

5. Additionnal: Check sample codes for details https://github.com/Excel-DNA/ExcelDna
```
ExcelDna-master\ExcelDna-master\Distribution\Samples\ComServer 
```
