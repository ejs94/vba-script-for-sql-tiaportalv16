Sub DesabilitarRegra(ByRef idRegra)
'DESABILITA REGRA DA TABELA

'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strErro, idTagEstagio, descricao, DTEvento, STR_DTEvento

On Error Resume Next

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	ShowSystemAlarm "Erro DesabilitarRegra, conn.Open: #" & Err.Number & " " & Err.Description
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If


'EDITA ENABLED DO ID ESPECIFICO
'UPDATE [dbo].[tb_prod_manut_plan]
'   Set [enabled] = x
' WHERE [id_manut_plan] = y
    
    
SQL_Table = "UPDATE [dbo].[tb_pr_manut_plan] " & _
			"SET [enabled] = 0 " & _
			"WHERE [id_manut_plan] = " & idRegra
			  
STD_Log SQL_Table

Err.Clear
'EXECUTA COMANDO
Set rst = conn.Execute(SQL_Table)

'TRATA ERROS
If Err.Number <> 0 Then
	ShowSystemAlarm "Erro DesabilitarRegra, conn.Execute:  #" & Err.Number & " " & Err.Description  & " " & SQL_Table
	Err.Clear
	'Close data source
	conn.close
	Set conn = Nothing
	Set rst = Nothing 
	Exit Sub
End If

'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing

End Sub