'argument: <ExcelFile> <MacroName>
'ex: "c:\xxx\yyy\zzz.xlsm" Macro1

Dim obj
Set obj=WScript.CreateObject("Excel.Application")

obj.Visible=False
obj.Workbooks.Open WScript.Arguments(0)
obj.Application.Run WScript.Arguments(1)