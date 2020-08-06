'Precisa de uma Tag: "Debug" para HMI
'Input: mensagem que é motivo do erro.
'Retorna: mensagem no sistema de alarmes.
Sub STD_Log(ByRef pMsg)
'Loga em System Messages, caso debug esteja ativo
If SmartTags("Debug") = True Then
	ShowSystemAlarm pMsg
End If

End Sub