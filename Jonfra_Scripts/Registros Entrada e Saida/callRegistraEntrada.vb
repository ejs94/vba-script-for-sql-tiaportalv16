Sub callRegistraEntrada(ByRef PLCTag)
'Alterar propriedade da PLCTag para monitorar Ciclicamente
Dim strFuncName

strFuncName = "callRegistraEntrada"

On Error Resume Next

'Just work if the PLCTag is Boolean
If PLCTag = True Then
    showLog strFuncName & ": Chamou a Sub RegistraEntrada"
	Call RegistraEntrada()
    showLog strFuncName & ": Saiu da Sub RegistraEntrada"
    Exit Sub
End If

showLog strFuncName & " : Requisição : " & PLCTag

'Error routine - Fehlerroutine
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & " Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Exit Sub
End If


End Sub