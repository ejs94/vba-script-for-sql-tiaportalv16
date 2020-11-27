Sub VisualizaDadosRegra()
'VISUALIZA VALORES DOS DADOS DA TABELA ( PROGRAMADO / EFETIVO)

'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strErro, idTagEstagio, idRegra, DateLim, Limit, Id_Tag, VarTag, Data, Enabled, Descricao, c

On Error Resume Next


Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "conn.Open  Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" ,"VisualizaDadosRegra"
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If

PreencheRegras

If SmartTags("DadosManutPlan 1") Then idRegra=SmartTags("ID_ManutPlan_1")
If SmartTags("DadosManutPlan 2") Then idRegra=SmartTags("ID_ManutPlan_2")
If SmartTags("DadosManutPlan 3") Then idRegra=SmartTags("ID_ManutPlan_3")
If SmartTags("DadosManutPlan 4") Then idRegra=SmartTags("ID_ManutPlan_4")

If SmartTags("DadosManutPlan 1") Then
	SmartTags("Cor_ManutPlan_1")=2
End If

If SmartTags("DadosManutPlan 2") Then
	SmartTags("Cor_ManutPlan_2")=2
End If

If SmartTags("DadosManutPlan 3") Then
	SmartTags("Cor_ManutPlan_3")=2
End If

If SmartTags("DadosManutPlan 4") Then
	SmartTags("Cor_ManutPlan_4")=2
End If

SmartTags("DadosManutPlan 1")= False 
SmartTags("DadosManutPlan 2")= False 
SmartTags("DadosManutPlan 3")= False 
SmartTags("DadosManutPlan 4")= False   

If idRegra = "" Then
	STD_Log "idRegra Vazio"
	SmartTags("Field_2")= ""
	SmartTags("Field_1")= ""
	
	'Close data source
	conn.close
	
	Set rst = Nothing
	Set conn = Nothing
	Exit Sub
End If


'********** cicloEstagioIndex e cicloTagIndex ************* 
SQL_Table = "SELECT " & _
	"[id_tag_param], " & _
	"[datetime_limit], " & _
	"[id_tag_alarm], " & _
	"[info_alarm], " & _
	"[absolute_limit], " & _
	"[info_alarm], " & _
	"[TG].tag_description " & _
	"FROM [dbo].[tb_pr_manut_plan] AS MP " & _
	"LEFT JOIN [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags] as TG On MP.id_tag_param = TG.id_tag " & _
	"WHERE [id_manut_plan] = " & idRegra
	  
Err.Clear
STD_Log SQL_Table

'EXECUTA COMANDO
Set rst = conn.Execute(SQL_Table)


If Err.Number= 0 Then
	SmartTags("cicloEstagioIndex") = rst.Fields(2).Value
	'SmartTags("cicloTarget") = rst.Fields(4).Value
	'SmartTags("cicloDescription") = rst.Fields(5).Value
	Descricao = rst.Fields(6).Value
	PreencheCBTags


	For c=1 To 10
		If SmartTags("strTag"& c ) = Descricao Then
			SmartTags("cicloTagIndex") = c
			Exit For
		End If
	Next
End If

SmartTags("Field_2")= ""
'***************************************************************




SQL_Table = "SELECT " & _
			"[id_tag_param], " & _
			"[datetime_limit], " & _
			"[id_tag_alarm], " & _
			"[info_alarm], " & _
			"[absolute_limit], " & _
			"[enabled] " & _
			"FROM [dbo].[tb_pr_manut_plan] " & _
			"WHERE [id_manut_plan] = " & idRegra
			  
STD_Log  SQL_Table

Err.Clear

'EXECUTA COMANDO
Set rst = conn.Execute(SQL_Table)




'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "conn.Execute:  #" & SQL_Table,"VisualizaDadosRegra linha 75"
	Err.Clear
	'Close data source
	conn.close
	Set conn = Nothing
	Set rst = Nothing
	SmartTags("Field_2")=0
	SmartTags("Field_1")=0
	Exit Sub
Else
	
	Limit=rst.Fields(4).Value
	DateLim=rst.Fields(1).Value
	Id_Tag=rst.Fields(0).Value
	Enabled = rst.Fields(6).Value
	
	STD_Log "VisualizaDadosRegra Linha 102 - Limit= "& Limit & ", DateLim= "& DateLim & ", Id_Tag= " & Id_Tag & ", Enabled= " & Enabled
End If

STD_Log "VisualizaDadosRegra Linha 105 - IdRegra= " & idRegra & ", Limit= "& Limit & ", DateLim= "& DateLim & ", Id_Tag= " & Id_Tag & ", Date= " & Date

'Tags Ciclo
If idRegra <> "" Then

	If Limit <> "" Then
		SmartTags("Field_2")= Limit
	Else
		SmartTags("Field_2")= DateLim
		SmartTags("Field_1")= Date
	End If
	
	STD_Log "VisualizaDadosRegra Linha 118 - Limit= " & Limit &", Date= " & Date
	
Else
	
	STD_Log "VisualizaDadosRegra Linha 122 - IdRegra= " & idRegra & ", Limit= "& Limit & ", DateLim= "& DateLim & ", Id_Tag= " & Id_Tag  & ", Date= " & Date
	
End If


	
If Id_Tag <> "" And Id_Tag > 0 Then

	'********** Verifica Valor Atual da Variável *************************
	SQL_Table = "SELECT * From [dbo].[tb_ana_tags] Where [id_tag]="& Id_Tag
	
	STD_Log  SQL_Table & vbNewLine & "IdTag= " & Id_Tag
	Err.Clear	
	
	'EXECUTA COMANDO
	Set rst = conn.Execute(SQL_Table)
	
	
	'TRATA ERROS
	If Err.Number <> 0 Then
		STD_Erro "conn.Execute:  #" & SQL_Table,"VisualizaDadosRegra Linha 105"
		Err.Clear
		'Close data source
		conn.close
		Set conn = Nothing
		Set rst = Nothing
		SmartTags("Field_2")= 0
		SmartTags("Field_1")= 0
		Exit Sub
	Else
		
		VarTag=rst.Fields(1).Value
		SmartTags("Field_1")= SmartTags(VarTag)
		STD_Log "Vartag = " & VarTag & vbNewLine
	End If

End If
   






'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing


End Sub