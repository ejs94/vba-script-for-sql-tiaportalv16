Sub PreencheCBTags()
'Preenche tags com informações cadastradas no banco
	
'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strAscDesc, i, j, INT_NrEstagio

On Error Resume Next

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "conn.Open Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP", "PreencheCBTags"
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If

'Nr estagio selecionado:
INT_NrEstagio = SmartTags("cicloEstagioIndex")

'SELECIONA SOMENTE TAGS QUE NAO SAO TRIGGERS DE ALARME E ORDENA DE ACORDO COM PARAMETRO

'SELECT tags.[id_tag]
'      ,[tag_name]
'      ,[tag_description]
'
'  FROM [dbo].[tb_ana_tags] tags
'  WHERE Not EXISTS (
'  Select alarm.id_tag FROM [dbo].[tb_ana_manut_plan_alarm] alarm WHERE alarm.id_tag = tags.id_tag)

SQL_Table = "SELECT " & _
			"[id_tag], " & _
			"[tag_name], " & _ 
			"[tag_description] " & _ 
			"FROM [dbo].[tb_ana_tags] " & _ 
			"WHERE[tb_ana_tags].[mostrar_cb] = 1 AND [nr_estagio] = " & INT_NrEstagio
			
'EXECUTA COMANDO SQL
Set rst = conn.Execute(SQL_Table)

'TRATA ERRO
If Err.Number <> 0 Then
	STD_Erro "Erro conn.Execute: #"  & SQL_Table, "PreencheCBTags"
	Err.Clear
	
	'ENCERRA
	conn.close
	Set conn = Nothing
	Set rst = Nothing 
	Exit Sub
End If
	
If Not (rst.EOF And rst.BOF) Then 
	'RETORNOU COM DADOS VÁLIDOS, PREENCHE TAGS:
	
	rst.MoveFirst 'PRIMEIRO DADO RECEBIDO 
	
	'ZERA ITERADOR
	j=0
	
	'VERIFICA QUANTIDADE DE ELEMENTOS NA TABELA
	Do
		j=j+1
		rst.MoveNext
	Loop Until rst.EOF
	
	rst.MoveFirst 'VOLTA AO PRIMEIRO DADO RECEBIDO 
	
	For i=1 To 30
		'Completa tabela de tags
		STD_Log "PreencheCBTags - SmartTags(strTag" & i & ")=" & rst.Fields(2).Value & ", " & "SmartTags(idTag" & i & ")=" & rst.Fields(0).Value
		
		If rst.EOF Then
			SmartTags("strTag" & i) = ""
			SmartTags("IdTag" & i) = 0
			Exit For
		Else
			SmartTags("strTag" & i) = rst.Fields(2).Value 'Field 2 = Description
			SmartTags("IdTag" & i) = rst.Fields(0).Value
			rst.MoveNext
		End If
	Next
	
	rst.close 
	
Else
	STD_Log "DADOS DE TAGS RETORNARAM VAZIOS!"
End If


'Close data source
conn.close



Set rst = Nothing
Set conn = Nothing



'Atualiza CB Tags 



SmartTags("cicloTagIndex")=1

'Call AtualizaTagContagem


End Sub