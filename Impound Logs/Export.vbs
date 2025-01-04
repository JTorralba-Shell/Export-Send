ScriptLocation = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
Clone = ScriptLocation & "\Impound Logs.xlsx"
Output = ScriptLocation & "\Impound Logs.pdf"

Dim currentDate
currentDate = Date
WScript.Echo "currentDate = " & currentDate

Dim currentYear
currentYear = Year(currentDate)

Dim goLiveDate
goLiveDate = cdate("Jan 2 " & currentYear)
WScript.Echo "goLiveDate = " & goLiveDate

If currentDate < goLiveDate Then
    YYYY = currentYear - 1
    WScript.Echo "currentDate < goLiveDate"
Else
    YYYY = currentYear
    WScript.Echo "currentDate >= goLiveDate"
End If

Source = "\\EPTEPCNAS\EPCCAD\Impound Logs\" & YYYY & " Impound Log.xlsx"
WScript.Echo "Source = " & Source

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile Source, Clone, True

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
