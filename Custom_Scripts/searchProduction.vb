Sub searchProduction(ByRef pOrdem, ByRef pInverteOrdem, ByRef pFiltroPN, ByRef pFiltroDataInicial, ByRef pFiltroDataFinal)
'////////////////////////////////////////////////////////////////
' Seleciona dados e Ordena de Acordo com o filtro setado na tela da IHM
' Ordenacao padrao = ID DESC
' Created: 10-08-2020
' Version: v1
' Author:  EJS 
'////////////////////////////////////////////////////////////////

'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strAscDesc, i, j, strFuncName
On Error Resume Next

strFuncName = "searchProduction"

'ABRIR CONEXAO
If Not connect_MSSQL(conn) Then	
	Exit Sub
End If

'Verifica Ordem e Inversão

'Inverte Ordem
If ((pInverteOrdem = 1) And (pOrdem = SmartTags("nOrdem"))) Then
	SmartTags("nAscDesc") = Not(SmartTags("nAscDesc"))
	showLog "Inverteu Ordem. Asc=" & CStr(SmartTags("nAscDesc"))
End If

'Ordenar
If pOrdem = "" Then
	SmartTags("nOrdem") = "Producao_id"
	SmartTags("nAscDesc") = False
Else
	SmartTags("nOrdem") = pOrdem
End If



'PESQUISA BANCO DE DADOS
showLog "Chamando Select"
Set rst = searchSQL(SmartTags("nOrdem"), SmartTags("nAscDesc"), pFiltroPN, pFiltroDataInicial, pFiltroDataFinal, conn)
	

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
	If SmartTags("nTab")>=j-12 Then
		SmartTags("nTab")=j-12
	End If
	If SmartTags("nTab")<j-11 Then
		For i=1 To SmartTags("nTab")
			rst.MoveNext
		Next
	End If
	If SmartTags("nTab")<0 Then
		SmartTags("nTab")=0
	End If
	
	'TODO : Alteras as Smartags para que fiquem conforme as tags configuradas para a tela.
	For i=1 To 12	
		'Completa tabela de tags
		If rst.EOF Then
			SmartTags("Value_ID_" & i) = ""
			SmartTags("Value_Data_" & i) = ""
			SmartTags("Value_Barcode_" & i) = ""
			SmartTags("Value_Status_" & i) = ""
			SmartTags("Value_DTInicio_" & i) = ""
			SmartTags("Value_DTFim_" & i) = ""
			SmartTags("Value_Modelo_" & i) = ""
		Else
			SmartTags("Value_ID_" & i) = rst.Fields(0).Value
			SmartTags("Value_Data_" & i) = rst.Fields(1).Value
			SmartTags("Value_Barcode_" & i) = rst.Fields(2).Value
			SmartTags("Value_Status_" & i) = rst.Fields(7).Value
			SmartTags("Value_DTInicio_" & i) = rst.Fields(3).Value
			If IsNull(rst.Fields(4)) Then SmartTags("Value_DTFim_" & i)=""
			SmartTags("Value_DTFim_" & i) = rst.Fields(4).Value
			SmartTags("Value_Modelo_" & i) = rst.Fields(8).Value
			rst.MoveNext
		End If
	Next
	
	rst.close 
	
Else
	showLog "DADOS RETORNARAM VAZIOS!"
	
	For i=1 To 12	
		'Apaga tabela de tags
			SmartTags("Value_ID_" & i) = ""
			SmartTags("Value_Data_" & i) = ""
			SmartTags("Value_Barcode_" & i) = ""
			SmartTags("Value_Status_" & i) = ""
			SmartTags("Value_DTInicio_" & i) = ""
			SmartTags("Value_DTFim_" & i) = ""
			SmartTags("Value_Modelo_" & i) = ""
	Next
End If

'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing

End Sub