Sub PreencheRegras()

'Preenche tags com informações cadastradas no banco
	
'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strAscDesc, i, j

On Error Resume Next

STD_Log("Preenchendo Regras")

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "conn.Open Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP","PreencheRegras"
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If

'SELECIONA DADOS E ORDENA DE ACORDO COM PARAMETRO

'Select [id_manut_plan], [info_alarm], [enabled],
'CONCAT([tags_alarm].[tag_description], ' - ', CASE WHEN datetime_limit IS NULL THEN
'CONCAT('Quando ', tags_param.tag_description, ' chegar em: ', absolute_limit)
'Else CONCAT('Em ', CAST(datetime_limit AS DATE)) END)
'FROM [SNP_PI1846_7oEixo_Banco].[dbo].[tb_pr_manut_plan]
'Left Join [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] tags_param On id_tag_param = tags_param.id_tag
'Join  [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] tags1_ALARM On id_tag_ALARM = tags1_ALARM.id_tag
'INNER Join [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] tags_alarm On id_tag_ALARM - Case WHEN [tags1_ALARM].[id_tag] > 7 Then 1 Else 0 End = tags_alarm.id_tag


SQL_Table = "SELECT [id_manut_plan], [info_alarm], [enabled], " & _
			"CONCAT([tags_alarm].[tag_description], ' - ', CASE WHEN datetime_limit IS NULL THEN " & _
			"CONCAT('Quando ', tags_param.tag_description, ' chegar em: ', absolute_limit) " & _
			"ELSE CONCAT('Em ', CAST(datetime_limit AS DATE)) END) " & _
			"FROM [SNP_PI1846_7oEixo_Banco].[dbo].[tb_pr_manut_plan] " & _
			"LEFT JOIN [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] tags_param ON id_tag_param = tags_param.id_tag " & _
			"JOIN  [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] tags1_ALARM On id_tag_ALARM = tags1_ALARM.id_tag " & _
			"INNER JOIN [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] tags_alarm On id_tag_ALARM - CASE WHEN [tags1_ALARM].[id_tag] > 7 THEN 1 ELSE 0 END = tags_alarm.id_tag "
			
STD_Log(SQL_Table)
'EXECUTA COMANDO SQL
Set rst = conn.Execute(SQL_Table)




'TRATA ERRO
If Err.Number <> 0 Then
	STD_Erro "Erro conn.Execute: #" & SQL_Table,"PreencheRegras"
	Err.Clear
	
	'ENCERRA
	conn.close
	Set conn = Nothing
	Set rst = Nothing 
	Exit Sub
	
End If
	

	
'If Not (rst.EOF And rst.BOF) Then 
	'RETORNOU COM DADOS VÁLIDOS, PREENCHE TAGS:
	
	rst.MoveFirst 'PRIMEIRO DADO RECEBIDO 
	
	'ZERA ITERADOR
	j=0
	
	'VERIFICA QUANTIDADE DE ELEMENTOS NA TABELA
	Do
		j=j+1
		STD_Log "Posição "&j  & " - Valores: " & rst.Fields(0).Value &", "&rst.Fields(1).Value &", "&rst.Fields(2).Value &", "&rst.Fields(3).Value & vbNewLine
		rst.MoveNext
	Loop Until rst.EOF
	
	rst.MoveFirst 'VOLTA AO PRIMEIRO DADO RECEBIDO 
	
	'VERIFICA SHIFT DE SETAS PRA CIMA / PRA BAIXO
	If SmartTags("nTab")>=j-4 Then
		SmartTags("nTab")=j-4
	End If
	If SmartTags("nTab")<j-3 Then
		For i=1 To SmartTags("nTab")
			rst.MoveNext
		Next
	End If
	If SmartTags("nTab")<0 Then
		SmartTags("nTab")=0
	End If
	
	
	
	For i=1 To 4	
		'Completa tabela de tags
		If rst.EOF Then
			SmartTags("ID_ManutPlan_" & i) = ""
			SmartTags("Descricao2_ManutPlan_" & i) = ""
			SmartTags("Descricao_ManutPlan_" & i) = ""
			SmartTags("Cor_ManutPlan_" & i) = 0
		Else
			SmartTags("ID_ManutPlan_" & i) = rst.Fields(0).Value
			SmartTags("Descricao2_ManutPlan_" & i) = "''" & rst.Fields(1).Value & "''"
			SmartTags("Descricao_ManutPlan_" & i) = "''" & rst.Fields(3).Value & "''"
			SmartTags("Cor_ManutPlan_" & i) = rst.Fields(2).Value
			rst.MoveNext
		End If
	Next
	
	rst.close 
'Else
	'ShowSystemAlarm "DADOS RETORNARAM VAZIOS!"
'End If

'Close data source
conn.close

Set rst = Nothing
Set conn = Nothing


SmartTags("eventDay")=Day(Date)
SmartTags("eventMonth")=Month(Date)
SmartTags("eventYear")=Year(Date)
SmartTags("DataPlanejada")=Date





End Sub