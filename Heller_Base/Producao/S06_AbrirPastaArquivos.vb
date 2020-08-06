Sub S06_AbrirPastaArquivos()
'Tip:
' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:

'ABRIR PASTA DE ARQUIVOS

Dim SH, txtFolderToOpen ,strFuncName
Set SH = CreateObject("WScript.Shell")


strFuncName = "S06_AbrirPastaArquivos"
txtFolderToOpen = "D:\Arquivos"


On Error Resume Next

If Err.Number<>0 Then  STD_Erro "Erro ao Abrir a Pasta " & txtFolderToOpen, strFuncName
Err.Clear
	

SH.Run txtFolderToOpen
Set SH = Nothing


End Sub