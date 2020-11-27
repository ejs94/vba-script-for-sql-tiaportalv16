Sub CarregarModelosSQLServer()
'Para Funcionar, precisa alterar configuração nas propriedades da TAG, para ciclico.
Dim strFuncName

strFuncName = "CarregarModelosSQLServer"

On Error Resume Next

'Just work if the PLCTag is Boolean
If SmartTags("DB110_IHM_IPC.Req_CarregarModelosSQLServer") = True Then
    showLog strFuncName & ": Entrou na Func"
	'Call updateTipoCargaPLC()
    SmartTags("DB110_IHM_IPC.Completo_CarregarModelosSQLServer") = True
    showLog strFuncName & ": Terminou a Func"
Else
    showLog strFuncName & ": Nem rodou a Func"
    Exit Sub
End If

'Error routine - Fehlerroutine
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & " Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Exit Sub
End If


End Sub