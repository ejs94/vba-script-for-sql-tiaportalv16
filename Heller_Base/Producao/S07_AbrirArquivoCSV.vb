Sub S07_AbrirArquivoCSV()
'Tip:
' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:

'ABRIR ARQUIVO

Dim txtFileToOpen ,strFuncName
Dim objWshShell

strFuncName = "S07_AbrirArquivoCSV"
txtFileToOpen = "D:\Arquivos"

On Error Resume Next


Set objWshShell = CreateObject("Wscript.Shell")

objWshShell.Run txtFileToOpen


If Err.Number<>0 Then
	STD_Erro "Erro ao Executar Comando objWshShell.Run " & txtFileToOpen, strFuncName
	Err.Clear
End If

txtFileToOpen = txtFileToOpen &"\"& SmartTags("MSG_FILENAME")


objWshShell.Exec "C:\Program Files (x86)\OpenOffice 4\program\soffice.exe " & txtFileToOpen


Set objWshShell=Nothing

STD_Log "C:\Program Files (x86)\OpenOffice 4\program\soffice.exe " & txtFileToOpen

If Err.Number<>0 Then  STD_Erro "Erro ao Executar Comando objWshShell.Exec " & "C:\Program Files (x86)\Microsoft Office\Office12\XLVIEW.EXE " & txtFileToOpen, strFuncName
Err.Clear
	



End Sub