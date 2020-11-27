Function querySQL(ByRef pConnection, ByRef pSQL_Table)
'Query ao DB manipulada pelo VBA do Tia Portal
' OrderBy deve ser uma String 'ASC' or 'DESC'
Dim rst, strFuncName

'Essas Tags precisam ser criadas na IHM e associadas aos diplays de input

On Error Resume Next

strFuncName = "querySQL"

'Se o Debug estiver ativado
showLog "Query: " & pSQL_Table

'EXECUTA COMANDO SQL
Set rst = pConnection.Execute(pSQL_Table)
	
'TRATA ERRO
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": conn.Execute: " & pSQL_Table
	Err.Clear
	'ENCERRA
	pConnection.close
	showLog strFuncName & ": Conexão com o MSSQL fechada"
	rst = Nothing
End If

'Retorna Resultset da pesquisa
Set querySQL = rst

End Function