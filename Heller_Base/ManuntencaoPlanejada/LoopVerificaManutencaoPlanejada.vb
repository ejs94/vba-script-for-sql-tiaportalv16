Sub LoopVerificaManutencaoPlanejada()
'Verifica cadastro de manutenção e dispara alarmes nos casos verdadeiros
	
'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strAscDesc, i, j, countT, countF,strFuncName, TxtLog, Valor, VTarget, Data
Dim IdRegra

strFuncName = "Loop Verifica Manutenção Planejada"

On Error Resume Next

'STD_Log("Iniciando Verificações de Manutenção Planejada")

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "Erro Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" , strFuncName
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If


'********************************************
'SELECIONA EVENTOS COM DATA DE HOJE PARA TRÁS
'********************************************

STD_Log "Linha 32 ID_MSG = " & SmartTags("ID_MSG")
SQL_Table = "SELECT [id_manut_plan], " & _
			"[datetime_limit], " & _
			"[tags_alarm].[tag_name], " & _
			"[info_alarm], " & _
			"[tags_alarm].[id_tag_texto_alarme], " & _
			"[tags_descricao].[tag_name] " & _
			"FROM [SNP_PI1846_7oEixo_Banco].[dbo].[tb_pr_manut_plan] " & _
			"INNER JOIN [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] tags_alarm ON id_tag_alarm = tags_alarm.id_tag " & _
  			"INNER JOIN [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] tags_descricao ON [tags_alarm].[id_tag_texto_alarme] = [tags_descricao].[id_tag] " & _
			"WHERE [datetime_limit] < GETDATE() AND [enabled] = 1"


Err.Clear			
'EXECUTA COMANDO SQL
STD_Log("Verificando Eventos...")
STD_Log(SQL_Table)
Set rst = conn.Execute(SQL_Table)

'TRATA ERRO
If Err.Number <> 0 Then
	STD_Erro "Erro Linha 53 conn.Execute: " & SQL_Table, strFuncName
	Err.Clear
	
	'ENCERRA
	conn.close
	Set conn = Nothing
	Set rst = Nothing 
	Exit Sub
End If
	
STD_Log "Field 2 = " & rst.Fields(2).Value &", Valor Field 2 = " & SmartTags(rst.Fields(2).Value)

If Not (rst.EOF) Then 
	'RETORNOU COM DADOS VÁLIDOS, SETA TAGS:
	
	rst.MoveFirst 'PRIMEIRO DADO RECEBIDO 
	'rst.MoveNext Gama

	While Not (rst.EOF) 
		If SmartTags(rst.Fields(2).Value) = False Then
			STD_Log "Encontrado Evento Manutenção Planejada: ID" & rst.Fields(0).Value & ", " & rst.Fields(2).Value & ": ''" & rst.Fields(3).Value & "''"
			
			'Escreve a descrição
			SmartTags(rst.Fields(5).Value) = rst.Fields(3).Value
			
			'Seta Alarme
			SmartTags(rst.Fields(2).Value) = True
			
			
			'UPDATE REGRA: DESABILITA NO BANCO
			STD_Log("Desabilitando Regra ID" & rst.Fields(0).Value)
			If SmartTags("ID_MSG")=0 Then 'ID_Msg - Verifica se já não existe uma mensagem na Tela
				DesabilitarRegra(rst.Fields(0).Value)
			End If
			
			'VERIFICA A EXISTÊCIA DE REGRAS E SINALIZA MENSAGEM DE FALHA (Pop-Up MSG)
			SmartTags("ID_MSG")= rst.Fields(0).Value
			SmartTags("Falhas_ManutPl")=True
			STD_Log "Linha 91 ID_MSG = " & SmartTags("ID_MSG")
			Call PopupMSG(rst.Fields(0).Value)
			
				
			
		End If
		
		rst.MoveNext
	Wend
	
	rst.close 
Else
	STD_Log("Não foram encontrados eventos de Manutenção Planejada")
End If


'********************************************
'SELECIONA CICLOS E VALIDA TARGETS
'********************************************

STD_Log "Linha 111 ID_MSG = " & SmartTags("ID_MSG")

'Verifica se N existem mensagens pendentes
If SmartTags("ID_MSG")<>0 Then Exit Sub

SQL_Table = "SELECT [id_manut_plan], " & _
			"[tags_param].[tag_name], " & _
			"[tags_alarm].[tag_name], " & _
			"[absolute_limit], " & _
			"[info_alarm], " & _
			"[tags_alarm].[id_tag_texto_alarme], " & _
			"[tags_descricao].[tag_name] " & _
			"FROM [SNP_PI1846_7oEixo_Banco].[dbo].[tb_pr_manut_plan] " & _
			"INNER JOIN [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] tags_param On id_tag_param = tags_param.id_tag " & _
			"INNER JOIN [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] tags_alarm On id_tag_alarm = tags_alarm.id_tag " & _
			"INNER JOIN [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] tags_descricao ON [tags_alarm].[id_tag_texto_alarme] = [tags_descricao].[id_tag] " & _
			"WHERE [enabled] = 1 "
			
'			SQL_Table = "SELECT * FROM [SNP_PI1846_7oEixo_Banco].[dbo].[tb_pr_manut_plan] WHERE [enabled] = 1 "
			
STD_Log("Verificando Ciclos...")
STD_Log(SQL_Table)

Err.Clear
'EXECUTA COMANDO SQL
Set rst = conn.Execute(SQL_Table)

'TRATA ERRO
If Err.Number <> 0 Then
	STD_Erro "Erro Linha 136 conn.Execute: " & SQL_Table, strFuncName
	Err.Clear
	
	'ENCERRA 	
	conn.close
	Set conn = Nothing
	Set rst = Nothing 
	Exit Sub
End If



TxtLog=""	

		 
If Not (rst.EOF And rst.BOF) Then 
	'RETORNOU COM DADOS VÁLIDOS, VERIFICA TAGS TARGETS:
	
	rst.MoveFirst 'PRIMEIRO DADO RECEBIDO 
	

	
	While Not (rst.EOF)
		'Valor atual > target?

		
		'************ Verifica se Contagem Atingiu Tatget ***************
		If SmartTags(rst.Fields(1).Value) >= rst.Fields(3).Value And rst.Fields(3).Value > 0 And SmartTags("ID_MSG")=0 Then 'ID_Msg - Verifica se já não existe uma mensagem na Tela
			STD_Log "Linha 164 Verificando atual/target (Field 1 > Field 3)"
			STD_Log "Field 1 = " & rst.Fields(1).Value & ", Valor = " & SmartTags(rst.Fields(1).Value)
			STD_Log "Field 2 = " & rst.Fields(2).Value & ", Valor = " & SmartTags(rst.Fields(2).Value)
			STD_Log "Field 3 = " & rst.Fields(3).Value	
			STD_Log "Encontrado Ciclo Manutenção Planejada: ID (" & rst.Fields(0).Value & ")"
			IdRegra = rst.Fields(0).Value
			
			STD_Log "Linha 176 Alarme não está ativado? " & rst.Fields(2).Value & "=" & SmartTags(rst.Fields(2).Value)
			'Alarme não está ativado?
			If SmartTags(rst.Fields(2).Value) = False Then
			
				STD_Log  "Set bit SmartTags(" & rst.Fields(2).Value & ")"
				SmartTags(rst.Fields(2).Value) = True
				
				'UPDATE REGRA: DESABILITA NO BANCO
				STD_Log "Desabilitando Regra ID" & IdRegra
				DesabilitarRegra(IdRegra)
				
				'Escreve a descrição
				SmartTags(rst.Fields(6).Value) = rst.Fields(4).Value
				
				
				'VERIFICA A EXISTÊCIA DE REGRAS E SINALIZA MENSAGEM DE FALHA (Pop-Up MSG)
				SmartTags("ID_MSG")= IdRegra
				SmartTags("Falhas_ManutPl")=True
				STD_Log "Linha 187 ID_MSG = " & IdRegra
				Call PopupMSG(IdRegra)
						
			End If
		End If
		rst.MoveNext
	Wend
	rst.close 

Else
	STD_Log "Não foram encontrados Ciclos de Manutenção Planejada"
End If






'Close data source
conn.close

Set rst = Nothing
Set conn = Nothing



End Sub