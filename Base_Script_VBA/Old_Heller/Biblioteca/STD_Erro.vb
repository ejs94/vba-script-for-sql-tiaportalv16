'Input: Mensagem do erro, Função ou subrotinha que está ocorrendo o erro.
'Retorna: No sistema de alarme irá printar o erro.
Sub STD_Erro(ByRef pMsg, ByRef pFuncName)
'Loga em System Messages
	ShowSystemAlarm "Erro " & Err.Number & "(" & pFuncName & "): " & Err.Description & ", " & pMsg
	Err.Clear
End Sub