'Precisa de uma Tag: "showLog" para HMI
'Input: mensagem que é motivo do erro.
'Retorna: mensagem no sistema de alarmes.
Sub showLog(ByRef pshowLogMessage)
'Loga em System Messages, caso Debug esteja ativo
If SmartTags("Debug") = True Then
	ShowSystemAlarm pshowLogMessage
End If

End Sub