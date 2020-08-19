Sub showProduction()
'////////////////////////////////////////////////////////////////
' Essa função atualiza todos os On-Display do tela para mostrar 
' os valores do SQL Server
' 
'Ordenacao padrao = ID Descrescente
' Created: 10-08-2020
' Version: v1
' Author:  EJS 
'////////////////////////////////////////////////////////////////

'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strAscDesc, i, j, strFuncName


On Error Resume Next

strFuncName = "showProduction"

'ABRIR CONEXAO
If Not connect_MSSQL(conn,"hmiDB") Then	
	Exit Sub
End If


'PESQUISA BANCO DE DADOS
showLog "Chamando Select"
Set rst = queryProduction(conn)
	

'BOF Indicates that the current record position is before the first record in a Recordset object. - Tabela está vazia
'EOF Indicates that the current record position is after the last record in a Recordset object. - Tabela chegou ao final das linhas

If Not (rst.EOF And rst.BOF) Then 
	'RETORNOU COM DADOS VÁLIDOS, PREENCHE TAGS:
	showLog "Retornou Dados Válidos"
	
	rst.MoveFirst 'PRIMEIRO DADO RECEBIDO 

	'ZERA ITERADOR
	j=0
	
	'VERIFICA QUANTIDADE DE ELEMENTOS NA TABELA
	Do
		j=j+1
		rst.MoveNext
	Loop Until rst.EOF
	
	rst.MoveFirst 'VOLTA AO PRIMEIRO DADO RECEBIDO 
	
	'VERIFICA SHIFT DE SETAS PRA CIMA / PRA BAIXO
	If SmartTags("nTab")>=j-13 Then
		SmartTags("nTab")=j-13
	End If
	If SmartTags("nTab")<j-12 Then
		For i=1 To SmartTags("nTab")
			rst.MoveNext
		Next
	End If
	If SmartTags("nTab")<0 Then
		SmartTags("nTab")=0
	End If
	showLog "Valores de i: " & i & " e j: " & j
	'TODO : Alteras as Smartags para que fiquem conforme as tags configuradas para a tela.
	For i=1 To 13	
		'Completa tabela de tags
		If rst.EOF Then
			SmartTags("ID_Value_" & i) = 0
			SmartTags("PN_Value_" & i) = ""
			SmartTags("Modelo_Value_" & i) = ""
			SmartTags("NomeModelo_Value_" & i) = ""
			SmartTags("BB155_Value_" & i) = ""
			SmartTags("BB165_Value_" & i) = ""
			SmartTags("BB175_Value_" & i) = ""
			SmartTags("BB185_Value_" & i) = ""
			SmartTags("Inspecao_Value_" & i) = ""
			SmartTags("DT_Inicio_Value_" & i) = ""
			SmartTags("DT_Fim_Value_" & i) = ""
		Else
			'Caso algum valor seja NULL isso irá evitar a replicação do valor do FOR_i Anterior
			If IsNull(rst.Fields(0)) Then SmartTags("ID_Value_" & i) = "NULL"
			If IsNull(rst.Fields(1)) Then SmartTags("PN_Value_" & i) = "NULL"
			If IsNull(rst.Fields(2)) Then SmartTags("Modelo_Value_" & i) = "NULL"
			If IsNull(rst.Fields(3)) Then SmartTags("NomeModelo_Value_" & I) = "NULL"
			If IsNull(rst.Fields(4)) Then SmartTags("BB155_Value_" & i) = "NULL"
			If IsNull(rst.Fields(5)) Then SmartTags("BB165_Value_" & i) = "NULL"
			If IsNull(rst.Fields(6)) Then SmartTags("BB175_Value_" & i) = "NULL"
			If IsNull(rst.Fields(7)) Then SmartTags("BB185_Value_" & i) = "NULL"
			If IsNull(rst.Fields(8)) Then SmartTags("Inspecao_Value_" & i) = "NULL"
			If IsNull(rst.Fields(9)) Then SmartTags("DT_Inicio_Value_" & i) = "NULL"
			If IsNull(rst.Fields(10))Then SmartTags("DT_Fim_Value_" & i) = "NULL"
			'Condição para escrever em toda tela
			SmartTags("ID_Value_" & i) = rst.Fields(0).Value
			SmartTags("PN_Value_" & i) = rst.Fields(1).Value
			SmartTags("Modelo_Value_" & i) = rst.Fields(2).Value
			SmartTags("NomeModelo_Value_" & i) = rst.Fields(3).Value
			SmartTags("BB155_Value_" & i) = rst.Fields(4).Value
			SmartTags("BB165_Value_" & i) = rst.Fields(5).Value 
			SmartTags("BB175_Value_" & i) = rst.Fields(6).Value
			SmartTags("BB185_Value_" & i) = rst.Fields(7).Value
			SmartTags("Inspecao_Value_" & i) = rst.Fields(8).Value
			'Como não achei uma melhor forma de converter em VBA DataTime para String, esse "quick-fix" é necessário			
			If Minute(rst.Fields(9).Value) < 10 Then
				SmartTags("DT_Inicio_Value_" & i) = Day(rst.Fields(9).Value) & "/" & Month(rst.Fields(9).Value) & " - " & Hour(rst.Fields(9).Value) & ":" & "0" & Minute(rst.Fields(9).Value)
			Else
				SmartTags("DT_Inicio_Value_" & i) = Day(rst.Fields(9).Value) & "/" & Month(rst.Fields(9).Value) & " - " & Hour(rst.Fields(9).Value) & ":" & Minute(rst.Fields(9).Value)
			End If
			If Minute(rst.Fields(9).Value) < 10 Then
				SmartTags("DT_Fim_Value_" & i) = Day(rst.Fields(10).Value) & "/" & Month(rst.Fields(10).Value) & " - " & Hour(rst.Fields(10).Value) & ":" & "0" & Minute(rst.Fields(10).Value)
			Else
				SmartTags("DT_Fim_Value_" & i) = Day(rst.Fields(10).Value) & "/" & Month(rst.Fields(10).Value) & " - " & Hour(rst.Fields(10).Value) & ":" & Minute(rst.Fields(10).Value)
			End If

			rst.MoveNext
		End If
	Next
	
	rst.close 
	
Else
	showLog "DADOS RETORNARAM VAZIOS!"
	
	For i=1 To 13	
		'Apaga tabela de tags
			SmartTags("ID_Value_" & i) = 0
			SmartTags("PN_Value_" & i) = ""
			SmartTags("Modelo_Value_" & i) = ""
			SmartTags("NomeModelo_Value_" & i) = ""
			SmartTags("BB155_Value_" & i) = ""
			SmartTags("BB165_Value_" & i) = ""
			SmartTags("BB175_Value_" & i) = ""
			SmartTags("BB185_Value_" & i) = ""
			SmartTags("Inspecao_Value_" & i) = ""
			SmartTags("DT_Inicio_Value_" & i) = ""
			SmartTags("DT_Fim_Value_" & i) = ""
	Next
End If

'Close data source - Fecha a conexão com o SQL Server
conn.close

Set rst = Nothing
Set conn = Nothing

End Sub