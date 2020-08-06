Sub AtualizarRegra(ByRef idRegra)
'PREENCHE OS CAMPOS PRO USUARIO RE-INSERIR, AO MESMO TEMPO DELETA ITEM DA TABELA

'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strErro, idTagEstagio, Descricao, DTEvento, STR_DTEvento, c

On Error Resume Next

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "Conn.Open Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" , "Atualiza Regra"
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If


'SELECIONA VALORES DO ID ESPECIFICO

'SQL_Table = "SELECT " & _
'			"[id_tag_param], " & _
'			"[datetime_limit], " & _
'			"[id_tag_alarm], " & _
'			"[info_alarm], " & _
'			"[absolute_limit], " & _
'			"[info_alarm] " & _
'			"FROM [dbo].[tb_pr_manut_plan] " & _
'			"WHERE [id_manut_plan] = " & idRegra
	
	SQL_Table = "SELECT " & _
			"[id_tag_param], " & _
			"[datetime_limit], " & _
			"[id_tag_alarm], " & _
			"[info_alarm], " & _
			"[absolute_limit], " & _
			"[info_alarm], " & _
			"[TG].tag_description " & _
			"FROM [dbo].[tb_pr_manut_plan] AS MP " & _
			"LEFT JOIN [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] as TG On MP.id_tag_param = TG.id_tag " & _
			"WHERE [id_manut_plan] = " & idRegra
			  
STD_Log SQL_Table

'EXECUTA COMANDO
Set rst = conn.Execute(SQL_Table)

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "Conn.Execute " & SQL_Table, "Atualiza Regra Linha 43"
	Err.Clear
	'Close data source
	conn.close
	Set conn = Nothing
	Set rst = Nothing 
	Exit Sub
End If

'Tags Eventos
If Not IsNull(rst.Fields(1).Value) Then 
	STD_Log "Atualizar evento, Fields(1)=" & rst.Fields(1).Value &  ", eventEstagioIndex = Fields(2)=" & rst.Fields(2).Value & _
	", Fields(3)=" & rst.Fields(3).Value
	SmartTags("eventEstagioIndex") = rst.Fields(2).Value
	SmartTags("eventYear") = Year(rst.Fields(1).Value)
	SmartTags("eventMonth") = Month(rst.Fields(1).Value)
	SmartTags("eventDay") = Day(rst.Fields(1).Value)
	SmartTags("eventDescription") = rst.Fields(5).Value

Else
	
	'Tags Ciclo
	STD_Log "Atualizar ciclo"

	SmartTags("cicloEstagioIndex") = rst.Fields(2).Value
	SmartTags("cicloTarget") = rst.Fields(4).Value
	SmartTags("cicloDescription") = rst.Fields(5).Value
	Descricao = rst.Fields(6).Value
	PreencheCBTags


	For c=1 To 10
		If SmartTags("strTag"& c ) = Descricao Then
			SmartTags("cicloTagIndex") = c
			Exit For
		End If
	Next
	
End If




        
SQL_Table = "DELETE FROM [dbo].[tb_pr_manut_plan] " & _
			"WHERE [id_manut_plan] = " & idRegra
			  
 
 Err.Clear  
'EXECUTA COMANDO
Set rst = conn.Execute(SQL_Table)

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "Conn.Execute " & SQL_Table, "Atualiza Regra Linha 100"
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