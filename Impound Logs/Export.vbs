ScriptLocation = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
Clone = ScriptLocation & "\Impound Logs.xlsx"
Output = ScriptLocation & "\Impound Logs.pdf"

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile "\\Server_Folder_Path\2020 Impound Log.xlsx", Clone, True

Const PaperLegal = 5
Const Landscape  = 2

Dim App
Dim XLSX

Set App = createObject("Excel.Application")
App.Visible = true
App.UserControl = true

Set XLSX = App.Workbooks.Open(Clone)

XLSX.ActiveSheet.PageSetup.Orientation = Landscape
XLSX.ActiveSheet.PageSetup.CenterHorizontally = True
XLSX.ActiveSheet.ExportAsFixedFormat 0, Output
XLSX.Close True
Set XLSX = Nothing

App.Quit
Set App = Nothing

Set FSO = Nothing
