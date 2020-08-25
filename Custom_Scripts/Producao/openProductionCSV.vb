Sub openProductionCSV()

Dim MyFolder,strFuncName
Dim objWshShell

strFuncName = "openProductionCSV"
MyFolder = "D:\ArquivosCSV\"

On Error Resume Next

Set objWshShell = CreateObject("Wscript.Shell")

objWshShell.Run MyFolder

If Err.Number<>0 Then
	ShowSystemAlarm "Erro " & Err.Number & strFuncName & ":" & Err.Description & ", " & "Erro ao Executar Comando objWshShell.Run"
	Err.Clear
End If

MyFolder = MyFolder &"\TMP.csv"


objWshShell.Exec "C:\Program Files (x86)\CSV Viewer\CSVViewer.exe " & MyFolder


Set objWshShell=Nothing

showLog "C:\Program Files (x86)\CSV Viewer\CSVViewer.exe " & MyFolder

If Err.Number<>0 Then 
	ShowSystemAlarm "Erro ao Executar Comando objWshShell.Exec " & "C:\Program Files (x86)\CSV Viewer\CSVViewer.exe " & MyFolder & ", " & strFuncName
	Err.Clear
End If

End Sub