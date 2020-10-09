Sub InserirRegraEvento()
'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strErro, idTagEstagio, descricao, DTEvento, STR_DTEvento

On Error Resume Next
Err.Clear
STD_Log("Inserindo Regra Evento")

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "conn.Open Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP","InserirRegraEvento"
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If


'Tags
idTagEstagio = SmartTags("eventEstagioIndex")
DTEvento = CDate(SmartTags("eventYear") & "-" & SmartTags("eventMonth") & "-" & SmartTags("eventDay"))
STR_DTEvento = STD_DateISO(DTEvento)
descricao = SmartTags("eventDescription")

'VALORES A SEREM INSERIDOS
'INSERT INTO [dbo].[tb_prod_manut_plan]
'           ([datetime_limit]
'           ,[id_tag_alarm]
'           ,[info_alarm]
'			,[enabled])
'     VALUES
'           (<id_tag_param, Int,>
'           ,<id_tag_alarm, Int,>
'           ,<absolute_limit, Int,>
'           ,<descricao, nvarchar(150),>, 1)
    
    
SQL_Table = "INSERT INTO [dbo].[tb_pr_manut_plan] " & _
			"([datetime_limit], " & _
			"[id_tag_alarm], " & _
			"[info_alarm], " & _
			"[enabled]) " & _
			"VALUES (" & _
			"'" & STR_DTEvento & "', " & _
			"" & idTagEstagio & ", " & _
			"'" & descricao & "', " & _
			"1)"
			  
STD_Log(SQL_Table)
    
'EXECUTA COMANDO
Set rst = conn.Execute(SQL_Table)

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "Erro conn.Execute: #" & SQL_Table,"InserirRegraEvento"
	Err.Clear
	'Close data source
	conn.close
	Set conn = Nothing
	Set rst = Nothing 
	Exit Sub
Else
	STD_Log "Regra Evento Inserida Com Sucesso ( Estagio " & idTagEstagio & ", target " & ", Descr " & descricao & ")"
End If


'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing

'limpa campos
SmartTags("eventEstagioIndex") = 0
SmartTags("eventYear") = 2018
SmartTags("eventMonth") = 1
SmartTags("eventDay") = 1
SmartTags("eventDescription") = ""

PreencheRegras()
 
End Sub