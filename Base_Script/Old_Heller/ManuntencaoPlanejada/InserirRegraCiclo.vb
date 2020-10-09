Sub InserirRegraCiclo()
'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strErro, idTagEstagio, idTagParam, target, descricao, DTInicio, STR_DTInicio, Barcode, nCicloTagIndex, NextVal, MaxVal

On Error Resume Next

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "conn.Open Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP","InserirRegraCiclo"
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If


'Tags
idTagEstagio = SmartTags("cicloEstagioIndex")
nCicloTagIndex = SmartTags("cicloTagIndex")
idTagParam = SmartTags("IdTag" & nCicloTagIndex )
target = SmartTags("cicloTarget")
descricao = SmartTags("cicloDescription")

STD_Log "cicloEstagioIndex = "& idTagEstagio & ", SmartTags(cicloTagIndex) = " & SmartTags("cicloTagIndex") & _
", SmartTags(IdTag & SmartTags(cicloTagIndex)) = " & idTagParam & _
", SmartTags(cicloTarget) = " & target & ", SmartTags(cicloDescription) = " & descricao

'VALORES A SEREM INSERIDOS
    

SQL_Table = "INSERT INTO [dbo].[tb_pr_manut_plan] " & _
			"([id_tag_param] " & _
			",[id_tag_alarm] " & _
			",[absolute_limit] " & _
			",[info_alarm] " & _
			",[enabled]) " & _
			"VALUES (" & _
			"" & idTagParam & ", " & _
			"" & idTagEstagio & ", " & _
			"" & target & ", " & _
			"'" & descricao & "', " & _ 
			"1) "

'SQL_Table = "SELECT MAX (id_manut_plan) FROM [dbo].[tb_pr_manut_plan]"
''EXECUTA COMANDO
'Set rst = conn.Execute(SQL_Table)
'
'NextVal=rst.fields(0) + 1
'STD_Log "NextVal=" & NextVal
'If Not IsNumeric(NextVal) Then NextVal=1

'(NEXT VALUE FOR [dbo].[producao_id_sequence]
'SQL_Table = "INSERT INTO [dbo].[tb_pr_manut_plan] " & _
'			"([id_manut_plan] " & _
'			",[id_tag_param] " & _
'			",[id_tag_alarm] " & _
'			",[absolute_limit] " & _
'			",[info_alarm] " & _
'			",[enabled]) " & _
'			"VALUES (" & _
'			NextVal & ", " & _
'			"" & idTagParam & ", " & _
'			"" & idTagEstagio & ", " & _
'			"" & target & ", " & _
'			"'" & descricao & "', " & _ 
'			"1) "
'			  
STD_Log SQL_Table
    
'EXECUTA COMANDO
Set rst = conn.Execute(SQL_Table)

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "Erro conn.Execute: #" & SQL_Table,"InserirRegraCiclo"
	Err.Clear
	'Close data source
	conn.close
	Set conn = Nothing
	Set rst = Nothing 
	Exit Sub
Else
	STD_Log "Regra Inserida Com Sucesso (ID " & idTagParam & ", Estagio " & idTagEstagio & ", Target " & target & ", Descr " & descricao & ")"
End If



'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing

'limpa campos
SmartTags("cicloEstagioIndex") = 0
SmartTags("cicloTagIndex") = 0
SmartTags("cicloTarget") = 0
SmartTags("cicloDescription") = ""

PreencheRegras()
 
End Sub