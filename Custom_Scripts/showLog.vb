'Precisa de uma Tag: "showLog" para HMI
'Input: mensagem que é motivo do erro.
'Retorna: mensagem no sistema de alarmes.
Sub showLog(ByRef showLogMessage)
'Loga em System Messages, caso Debug esteja ativo
If SmartTags("Debug") = True Then
	ShowSystemAlarm showLogMessage
End If

End Sub