Sub showAllManPlan()

'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, i, j, strFuncName

On Error Resume Next

'Definição de Variavel
strFuncName = "showAllManPlan"

'Dá para usar o Fomart() do SQL Server e já mandar a String trabalhada
SQL_Table = " USE hmiDB;" &_
            "SELECT manPlan_id AS ID" &_
                ",equip AS 'Máq./Equip'" &_
                ",tipoManunt AS 'Tipo de Man.'" &_
                ",descri AS 'Descrição'" &_
	            ",priorid AS 'Priorid.'" &_
                ",FORMAT(dia_manunt, 'dd/MM/yyyy') AS 'Dia Reserv.'" &_
                ",FORMAT(hr_planej, 'hh\:mm') AS 'Hrs. Planej.'" &_
	            ",resposavel AS 'Responsavel'" &_
            "FROM manPlanejada"

SQL_Table = SQL_Table & " WHERE ativo=1 " & " ;"

'ABRIR CONEXAO
showLog "Conectanto ao Banco de Dados"
If Not connect_MSSQL(conn,"hmiDB") Then	
	Exit Sub
End If

'PESQUISA BANCO DE DADOS
showLog "Chamando a Query"
Set rst = querySQL(conn,SQL_Table)
	
'BOF Indicates that the current record position is before the first record in a Recordset object. - Tabela está vazia
'EOF Indicates that the current record position is after the last record in a Recordset object. - Tabela chegou ao final das linhas

If Not (rst.EOF And rst.BOF) Then 
	'RETORNOU COM DADOS VÁLIDOS, PREENCHE TAGS:
	showLog "Retornou Dados!"
	
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
	If SmartTags("manut_Tab")>=j-6 Then
		SmartTags("manut_Tab")=j-6
	End If

	If SmartTags("manut_Tab")<j-5 Then
		For i=1 To SmartTags("manut_Tab")
			rst.MoveNext
		Next
	End If

	If SmartTags("manut_Tab")<0 Then
		SmartTags("manut_Tab")=0
	End If
	
	showLog "Valores de i: " & i & " e j: " & j
	'TODO : Alteras as Smartags para que fiquem conforme as tags configuradas para a tela.
	For i=1 To 6	
		'Completa tabela de tags
		If rst.EOF Then
			SmartTags("ID_Manuntecao_" & i) = 0
			SmartTags("Maquina_Field_" & i) = ""
			SmartTags("TipoMaquina_Field_" & i) = ""
			SmartTags("Descricao_Field_" & i) = ""
			SmartTags("Prioridade_Field_" & i) = ""
			SmartTags("Dia_Reservado_" & i) = ""
			SmartTags("Hrs_Planejada_" & i) = ""
            SmartTags("Responsavel_" & i) = ""
		Else
			'Caso algum valor seja NULL isso irá evitar a replicação do valor i Anterior
			If IsNull(rst.Fields(0)) Then SmartTags("ID_Manuntecao_" & i) = ""
			If IsNull(rst.Fields(1)) Then SmartTags("Maquina_Field_" & i) = ""
			If IsNull(rst.Fields(2)) Then SmartTags("TipoMaquina_Field_" & i) = ""
			If IsNull(rst.Fields(3)) Then SmartTags("Descricao_Field_" & i) = ""
			If IsNull(rst.Fields(4)) Then SmartTags("Prioridade_Field_" & i) = ""
			If IsNull(rst.Fields(5)) Then SmartTags("Responsavel_" & i) = ""
			If IsNull(rst.Fields(6)) Then SmartTags("Dia_Reservado_" & i) = ""
			If IsNull(rst.Fields(7))Then SmartTags("Hrs_Planejada_" & i) = ""
			' Chaveamento
			
			'Condição para escrever em toda tela
			SmartTags("ID_Manuntecao_" & i) = rst.Fields(0).Value
			SmartTags("Maquina_Field_" & i) = rst.Fields(1).Value
			SmartTags("TipoMaquina_Field_" & i) = rst.Fields(2).Value
			SmartTags("Descricao_Field_" & i) = rst.Fields(3).Value
			SmartTags("Prioridade_Field_" & i) = rst.Fields(4).Value
			'Como não achei uma melhor forma de converter em VBA DataTime para String, esse "quick-fix" é necessário			
			SmartTags("Dia_Reservado_" & i) = Day(rst.Fields(5).Value) & "/" & Month(rst.Fields(5).Value) & "/" & Year(rst.Fields(5).Value)
			SmartTags("Hrs_Planejada_" & i) = rst.Fields(6).Value
            SmartTags("Responsavel_" & i) = rst.Fields(7).Value
			rst.MoveNext
		End If
	Next
	
	rst.close 
	
Else
	showLog "DADOS RETORNARAM VAZIOS!"
	
	For i=1 To 6	
		'Apaga tabela de tags
			SmartTags("ID_Manuntecao_" & i) = 0
			SmartTags("Maquina_Field_" & i) = ""
			SmartTags("TipoMaquina_Field_" & i) = ""
			SmartTags("Descricao_Field_" & i) = ""
			SmartTags("Prioridade_Field_" & i) = ""			
			SmartTags("Dia_Reservado_" & i) = ""
			SmartTags("Hrs_Planejada_" & i) = ""
            SmartTags("Responsavel_" & i) = ""
	Next
End If

'Close data source - Fecha a conexão com o SQL Server
conn.close

Set rst = Nothing
Set conn = Nothing


End Sub