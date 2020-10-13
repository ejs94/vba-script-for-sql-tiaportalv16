Function queryAllModels(ByRef pConnection)
'////////////////////////////////////////////////////////////////
' Query ao DB manipulada pelo VBA do Tia Portal
' essa sub é chamada pela tela da IHM.
' 
'
' Created: 12-10-2020
' Version: v0.7
' Author:  EJS 
'////////////////////////////////////////////////////////////////

Dim rst, SQL_Table, strAscDesc, strFuncName

'Essas Tags precisam ser criadas na IHM e associadas aos diplays de input

On Error Resume Next

strFuncName = "queryAllModels"

SQL_Table = "USE hmiDB; " &_
		"SELECT Modelo_id, ModeloString, NomeModelo, DiametroCamisa From ModelosBlocos "

'Ordena para padrão decrescente
SQL_Table = SQL_Table & ";"

'Se o Debug estiver ativado
showLog "Select: " & SQL_Table

'EXECUTA COMANDO SQL
Set rst = pConnection.Execute(SQL_Table)

'TRATA ERRO
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": conn.Execute: " & SQL_Table
	Err.Clear
	'ENCERRA
	pConnection.close
	showLog strFuncName & ": Conexão com o MSSQL fechada"
	rst = Nothing
End If

showLog "Retornando ResultSet"

'Retorna Resultset da pesquisa
Set queryAllModels = rst

End Function