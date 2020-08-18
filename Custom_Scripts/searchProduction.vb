Sub searchProduction(ByRef pColuna_Ordem, ByRef pInverteOrdem, ByRef pSearchPN, ByRef pFiltroDataInicial, ByRef pFiltroDataFinal)
'////////////////////////////////////////////////////////////////
' Seleciona dados e Ordena de Acordo com o filtro setado na tela da IHM
' Ordenacao padrao = ID Descrescente
' Created: 10-08-2020
' Version: v1
' Author:  EJS 
'////////////////////////////////////////////////////////////////

'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strAscDesc, i, j, strFuncName, Tipo_Ordem

pColuna_Ordem = SmartTags("Coluna_Ordem") 'Coluna que será baseada a ordem
Tipo_Ordem = SmartTags("Tipo_Ordem") 'Crescente ou Descrescente

On Error Resume Next

strFuncName = "searchProduction"

'ABRIR CONEXAO
If Not connect_MSSQL(conn,"hmiDB") Then	
	Exit Sub
End If

'Verifica Ordem e Inversão

'Inverte Ordem
'If ((pInverteOrdem = 1) And (pOrdem = SmartTags("nOrdem"))) Then
'	SmartTags("nAscDesc") = Not(SmartTags("nAscDesc"))
'	showLog "Inverteu Ordem. Asc=" & CStr(SmartTags("nAscDesc"))
'End If

'Ordenar
'If pOrdem = "" Then
'	SmartTags("nOrdem") = "Producao_id"
'	SmartTags("nAscDesc") = False
'Else
'	SmartTags("nOrdem") = pOrdem
'End If



'PESQUISA BANCO DE DADOS
showLog "Chamando Select"
Set rst = queryProduction(Coluna_Ordem, Tipo_Ordem, pSearchPN, pFiltroDataInicial, pFiltroDataFinal, conn)
	

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
	If SmartTags("nTab")<j-11 Then
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
		Elses
			SmartTags("ID_Value_" & i) = rst.Fields(0).Value
			SmartTags("PN_Value_" & i) = rst.Fields(1).Value
			SmartTags("Modelo_Value_" & i) = rst.Fields(2).Value
			SmartTags("NomeModelo_Value_" & i) = rst.Fields(3).Value
			SmartTags("BB155_Value_" & i) = rst.Fields(4).Value
			SmartTags("BB165_Value_" & i) = rst.Fields(5).Value 
			SmartTags("BB175_Value_" & i) = rst.Fields(6).Value
			SmartTags("BB185_Value_" & i) = rst.Fields(7).Value
			SmartTags("Inspecao_Value_" & i) = rst.Fields(8).Value
			SmartTags("DT_Inicio_Value_" & i) = Hour(rst.Fields(9).Value) & ":" & Minute(rst.Fields(9).Value) & ":" & Second(rst.Fields(9).Value)
			SmartTags("DT_Fim_Value_" & i) = Hour(rst.Fields(10).Value) & ":" & Minute(rst.Fields(10).Value) & ":" & Second(rst.Fields(10).Value)
			rst.MoveNext
		End If
	Next
	
	rst.close 
	
Else
	showLog "DADOS RETORNARAM VAZIOS!"
	
	For i=1 To 13	
		'Apaga tabela de tags
			SmartTags("ID_Value_" & i) = ""
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