ScriptLocation = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
Clone = ScriptLocation & "\Missing Persons.mdb"
Output = ScriptLocation & "\Missing Persons.pdf"


Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile "\\Server_Folder_Path\Missing Persons\Database - Missing Persons.mdb", Clone, True

Dim App

Set App = createObject("Access.Application")

App.OpenCurrentDataBase Clone, False,"yourmama"

App.DoCmd.OutputTo 3,"detectives report","PDF Format (*.pdf)", Output

App.Quit
Set App = Nothing
