Sub showallModels()
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

strFuncName = "showallModels"

'ABRIR CONEXAO
If Not connect_MSSQL(conn,"hmiDB") Then	
	Exit Sub
End If


'PESQUISA BANCO DE DADOS
showLog "Chamando Select"
Set rst = queryAllModels(conn)
	

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
	If SmartTags("model_nTab")>=j-13 Then
		SmartTags("model_nTab")=j-13
	End If

	If SmartTags("model_nTab")<j-12 Then
		For i=1 To SmartTags("model_nTab")
			rst.MoveNext
		Next
	End If

	If SmartTags("model_nTab")<0 Then
		SmartTags("model_nTab")=0
	End If
	
	showLog "Valores de i: " & i & " e j: " & j
	'TODO : Alteras as Smartags para que fiquem conforme as tags configuradas para a tela.
	For i=1 To 13	
		'Completa tabela de tags
		If rst.EOF Then
			SmartTags("ID_Model_Value_" & i) = 0
			SmartTags("ModelString_Value_" & i) = ""
			SmartTags("ModelNameString_Value_" & i) = ""
		Else
			'Caso algum valor seja NULL isso irá evitar a replicação do valor i Anterior
			If IsNull(rst.Fields(0)) Then SmartTags("ID_Model_Value_" & i) = ""
			If IsNull(rst.Fields(1)) Then SmartTags("ModelString_Value_" & i) = ""
			If IsNull(rst.Fields(2)) Then SmartTags("ModelNameString_Value_" & i) = ""
			'Condição para escrever em toda tela
			SmartTags("ID_Model_Value_" & i) = rst.Fields(0).Value
			SmartTags("ModelString_Value_" & i) = rst.Fields(1).Value
			SmartTags("ModelNameString_Value_" & i) = rst.Fields(2).Value

			rst.MoveNext
		End If
	Next
	
	rst.close 
	
Else
	showLog "DADOS RETORNARAM VAZIOS!"
	
	For i=1 To 13	
		'Apaga tabela de tags
			SmartTags("ID_Model_Value_" & i) = 0
			SmartTags("ModelString_Value_" & i) = ""
			SmartTags("ModelNameString_Value_" & i) = ""
	Next
End If

'Close data source - Fecha a conexão com o SQL Server
conn.close

Set rst = Nothing
Set conn = Nothing

End Sub