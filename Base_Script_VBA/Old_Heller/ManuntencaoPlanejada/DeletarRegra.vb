Sub DeletarRegra(ByRef idRegra)
'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strErro

On Error Resume Next

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	ShowSystemAlarm "Erro DeletarRegra, conn.Open: #" & Err.Number & " " & Err.Description
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If


'Tags

'DELETAR ITEM DA TABELA

'DELETE FROM [dbo].[tb_prod_manut_plan]
'      WHERE <Search Conditions,,>

'INSERT INTO [dbo].[tb_prod_manut_plan]
'           ([id_tag_param]
'           ,[id_tag_alarm]
'           ,[absolute_limit]
'           ,[info_alarm])
'     VALUES
'           (<id_tag_param, Int,>
'           ,<id_tag_alarm, Int,>
'           ,<absolute_limit, Int,>
'           ,<descricao, nvarchar(150),>)
    
    
SQL_Table = "DELETE FROM [dbo].[tb_pr_manut_plan] " & _
			"WHERE [id_manut_plan] = " & idRegra
			  
    ShowSystemAlarm SQL_Table
    
'EXECUTA COMANDO
Set rst = conn.Execute(SQL_Table)

'TRATA ERROS
If Err.Number <> 0 Then
	ShowSystemAlarm "Erro DeletarRegra, conn.Execute:  #" & Err.Number & " " & Err.Description  & " " & SQL_Table
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

PreencheRegras()
End Sub