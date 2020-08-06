Sub PopupMSG(ByRef ID_Msg)
'VERIFICA SE EXISTEM OCORRÊNCIAS DE MANUTENÇÃO PLANEJADA

'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, Falhas_P1, Falhas_P2, Falhas_P3, Falhas_P4, Falhas_P5, Falhas_P6, Reset_P1, Reset_P2, Reset_P3, Reset_P4, Reset_P5, Reset_P6
Dim ID_Manut, ID_Tag_Param, Date_Limit, Contagem, Msg1, Msg2

On Error Resume Next
Err.Clear
'************** Verifica e Sinaliza Mensagem de Falhas *****************
ID_Manut = ID_Msg

STD_Log "PopupMSG Linha 13 - ID_MSG=" & ID_Manut

If ID_Manut > 0 Then
	
	STD_Log "PopupMSG Linha 17 - ID_Manut=" & ID_Manut  
	
	Set conn = CreateObject("ADODB.Connection")
	Set rst = CreateObject("ADODB.Recordset")
	
	'ABRIR CONEXAO
	conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database
	
	'TRATA ERROS
	If Err.Number <> 0 Then
		STD_Erro "conn.Open Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP","PopupMSG"
		Err.Clear
	End If



	SQL_Table = "SELECT * " & _
	"FROM [SNP_PI1846_7oEixo_Banco].[dbo].[tb_pr_manut_plan] " & _
	"WHERE id_manut_plan=" & ID_Manut
			
	STD_Log "PopupMSG Linha 37 - " & (SQL_Table)
	
	Err.Clear
	'EXECUTA COMANDO SQL
	Set rst = conn.Execute(SQL_Table)
	
	
	'TRATA ERRO
	If Err.Number <> 0 Then
		STD_Erro "Erro conn.Execute: #" & SQL_Table,"PopupMSG"
		Err.Clear
		
		'ENCERRA
		conn.close
		Set conn = Nothing
		Set rst = Nothing
	
	Else
			ID_Manut = rst.Fields(0).Value
			ID_Tag_Param = rst.Fields(1).Value
			Msg2 = rst.Fields(5).Value
			Date_Limit=rst.Fields(2).Value
			Contagem=rst.Fields(4).Value
			STD_Log "Pop-up Linha 60 - ID_Manut=" & ID_Manut & ", ID_Tag_Param=" & ID_Tag_Param & ", Msg2=" & Msg2  & ", Date_Limit=" & Date_Limit & _
			", Contagemt=" & Contagem

	End If
	
	If ID_Tag_Param <> "" Then
		SQL_Table = "SELECT [tag_description] " & _
		"FROM [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] " & _
		"WHERE id_tag=" & ID_Tag_Param
				
		STD_Log "PopupMSG Linha 70- " & SQL_Table
		
		Err.Clear
		'EXECUTA COMANDO SQL
		Set rst = conn.Execute(SQL_Table)
	
		'TRATA ERRO
		If Err.Number <> 0 Then
			STD_Erro "Erro conn.Execute: #" & SQL_Table,"PopupMSG"
			Err.Clear
			
			'ENCERRA
			conn.close
			Set conn = Nothing
			Set rst = Nothing
		Else
			Msg1 = rst.Fields(0).Value & " Valor Programado: " & Contagem 
		End If
	Else
		Msg1 = "Data Planejada Para Manutenção (" & Date_Limit & ")"
	End If
	
	
	If Err.Number=0 Then 	
		SmartTags("MSG_ManutPlanejada1")= Msg1
		SmartTags("MSG_ManutPlanejada2")= Msg2
		Call ShowPopupScreen("MSG_ManutPlanejada",300,350,hmiOn,hmiBottom,hmiMedium)
	End If		

End If


End Sub