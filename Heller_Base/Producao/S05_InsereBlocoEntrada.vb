Sub S05_InsereBlocoEntrada(ByRef pRequest)
'////////////////////////////////////////////////////////////////
' Insere dados de bloco chegando na mesa Estaco 03
' Created: 20171205
' Version: v0.1
' Author:  IMF 
'////////////////////////////////////////////////////////////////

'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strErro, PartNumber, Modelo, DTInicio, STR_DTInicio, Barcode, IDProduto
Dim StrDate, PosChr13, strFuncName
strFuncName = "***** S05_InsereBlocoEntrada *****"

If Not pRequest Then
	SmartTags("DB_0629_Auto_CicloAuto_Aux_RespostaIPCentrada") = False 
	Exit Sub
End If

'Primeira coisa: Derruba Tag "Envia IPC"
SmartTags("DB_0629_Auto_CicloAuto_Cmd_EnviaIPCentrada") = False 



On Error Resume Next

SmartTags("ErroScript")=False ' Teste Erro

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "conn.Open",strFuncName
	Err.Clear
	Set conn = Nothing 
	SmartTags("DB_0629_Auto_CicloAuto_Aux_RespostaIPCentrada") = True
	Exit Sub
End If


'BUSCA INFORMACOES DE RASTREABILIDADE NAS TAGS
Modelo = SmartTags("DB_0070 - RastreabilidadeBloco_EsteiraEntrada_TipoBloco")
Barcode = SmartTags("DB_0070 - RastreabilidadeBloco_EsteiraEntrada_CodigoBloco")
Barcode=Replace(Barcode,Chr(13),"")'Filtra Caracter (ENTER)

DTInicio = SmartTags("DB_0070 - RastreabilidadeBloco_EsteiraEntrada_DataHoraEntrada")    ''atualizar
STR_DTInicio = STD_DT2DateTimeISO(DTInicio)

STR_DTInicio = STD_DT2DateTimeISO(Now)


StrDate=STR_DTInicio &" - " & Time

SmartTags("TesteString")=StrDate

'VERIFICA DUPLICATA
SQL_Table = "SELECT * FROM tb_pr_producao WHERE Barcode = '"& Barcode & "' AND DT_InicioProducao = '" & STR_DTInicio & "'"
'SQL_Table = "SELECT * FROM tb_pr_producao WHERE Part_Number = '"& PartNumber & "' AND DT_InicioProducao = '" & STR_DTInicio & "'"

Set rst = conn.Execute(SQL_Table)


'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "conn.Execute "& SQL_Table,strFuncName
	Err.Clear
	'Close data source
	conn.close
	Set conn = Nothing
	Set rst = Nothing 
	SmartTags("DB_0629_Auto_CicloAuto_Aux_RespostaIPCentrada") = True 
	Exit Sub
End If

If Not (rst.EOF And rst.BOF) Then 
	'SE CONSULTA RETORNAR ALGUM ITEM, SIGNIFICA QUE O BLOCO JÁ FOI INSERIDO ANTERIORMENTE!
 	
 	STD_Erro "Bloco duplicado detectado!",strFuncName
 	
 	'INSERE TABELA DE ERROS
 	'Guarda comando que deu erro
 	strErro = SQL_Table
 	SQL_Table = "INSERT INTO [dbo].[tb_pr_erros] ([SQL],[DT]) VALUES " & _
				"('" & strErro &"',GETDATE())"
         
	'EXECUTA COMANDO
	Set rst = conn.Execute(SQL_Table)
	
	'TRATA ERROS
	If Err.Number <> 0 Then
		STD_Erro "conn.Execute " & SQL_Table,strFuncName
		Err.Clear
		'Close data source
		conn.close
		Set conn = Nothing
		Set rst = Nothing 
		
		SmartTags("DB_0629_Auto_CicloAuto_Aux_RespostaIPCentrada") = False 
		Exit Sub
	End If
	
 	SmartTags("DB_0629_Auto_CicloAuto_Cmd_EnviaIPCentrada") = False 
	SmartTags("DB_0629_Auto_CicloAuto_Aux_RespostaIPCentrada") = True
	
	rst.close 
	
Else
	'VALORES A SEREM INSERIDOS
	'INSERT INTO [dbo].[tb_pr_producao]
    '      ([Data],[Barcode],[DT_InicioProducao],[DT_FimProducao],[Status_Peca]) VALUES 
    '      ('2018-03-20'
    '      ,12345
    '      ,'2018-03-20 00:00:00.000'
    '      ,'2018-03-20 21:00:00.000'
    '      ,1)
    'ou
    '       (GETDATE()
    '       ,12345
    '       ,GETDATE()
    '       ,GETDATE()
    '       ,1)
    
    
	SQL_Table = "INSERT INTO tb_pr_producao ([ID],[Data],[Barcode],[DT_InicioProducao],[DT_FimProducao],[id_status_peca], [id_modelo]) VALUES " & _
				"(NEXT VALUE FOR [dbo].[producao_id_sequence], '" & STR_DTInicio & "', '" & Barcode &"', '" & STR_DTInicio & "', null, 3, '" & Modelo & "')"
    
    STD_Log SQL_Table
    
	'EXECUTA COMANDO
	Set rst = conn.Execute(SQL_Table)
	
	SmartTags("VariavelControle")=SmartTags("VariavelControle")+1
	
	'TRATA ERROS
	If Err.Number <> 0 Then
		STD_Erro "conn.Execute " & SQL_Table,strFuncName
		Err.Clear
		'Close data source
		conn.close
		Set conn = Nothing
		Set rst = Nothing 
		

		SmartTags("DB_0629_Auto_CicloAuto_Aux_RespostaIPCentrada") = True
		Exit Sub
	End If
	
	STD_Log "Produto Inserido com Sucesso"
	SmartTags("DB_0629_Auto_CicloAuto_Aux_RespostaIPCentrada") = True 
End If

'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing

Call S02_SelectBlocosProducao("",0,SmartTags("nFiltroPN"),Date(Now),Date(Now))

End Sub