Sub S10_AtualizaSaidaBloco(ByRef pRequest)
'////////////////////////////////////////////////////////////////
' Insere dados de bloco saindo na estacao de saida (e26)
' Created: 20171205
' Version: v0.1
' Author:  IMF 
'////////////////////////////////////////////////////////////////

'DECLARACAO DE TAGs
Dim strFuncName, conn, rst, SQL_Table, strErro, PartNumber, Modelo, DTInicio, STR_DTInicio, DTFim, STR_DTFim, Barcode, IDProduto
strFuncName = "***** S10_AtualizaSaidaBloco *****"


If SmartTags("DB_0529_Auto_CicloAuto_Cmd_EnviaIPCsaida") Then
	SmartTags("DB_0529_Auto_CicloAuto_Aux_RespostaIPCsaida") = True
	SmartTags("DB_0529_Auto_CicloAuto_Cmd_EnviaIPCsaida")=False
End If

If Not pRequest Then Exit Sub

SmartTags("E05 - AuxLeituraBloco")=False

'
On Error Resume Next
Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "conn.Open", strFuncName & " 33"
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If

Barcode = SmartTags("DB_0071 - DadosBloco_Estagio05_LEITURA")
Barcode=Replace(Barcode,Chr(13),"")'Filtra Caracter (ENTER)


'VERIFICA EXISTENCIA DO BLOCO
SQL_Table = "SELECT * FROM tb_pr_producao WHERE Barcode LIKE '" & Barcode & "' AND id_status_Peca = 3"


Set rst = conn.Execute(SQL_Table)

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "conn.Execute " & SQL_Table,strFuncName & " 50"
	Err.Clear
	'Close data source
	conn.close
	Set conn = Nothing
	Set rst = Nothing 
	Exit Sub
End If

If (rst.EOF And rst.BOF) Then 
	'SE CONSULTA NAO RETORNAR NADA, BLOCO NUNCA ENTROU NA LINHA!
 	
	STD_Erro "Produto ("& Barcode & ") nao encontrado!"& vbNewLine & SQL_Table , strFuncName
 	
 	'INSERE TABELA DE ERROS
 	'Guarda comando que deu erro
 	strErro = Replace(SQL_Table, "'", "*")
 	SQL_Table = "INSERT INTO [dbo].[tb_pr_erros] ([SQL],[DT]) VALUES " & _
				"('" & strErro &"',GETDATE())"
         
	'EXECUTA COMANDO
	Err.Clear
	Set rst = conn.Execute(SQL_Table)
	
	'TRATA ERROS
	If Err.Number <> 0 Then
		STD_Erro "conn.Execute",strFuncName & " 77"
		Err.Clear
		'Close data source
		conn.close
		Set conn = Nothing
		Set rst = Nothing 
		Exit Sub
	End If
 	
	rst.close 
Else
	
	'BUSCA O ID DO PRODUTO NO BANCO
	rst.MoveFirst
	IDProduto = rst.Fields(0).Value
		
	'VALORES A SEREM EDITADOS
	'COMO RETORNOU ALGUM ITEM, SIGNIFICA QUE O BLOCO REALMENTE ESTA AGUARDANDO FINALIZAR!
    
    STD_Log strFuncName & " 96 IDProduto=" & IDProduto
	SQL_Table = "UPDATE tb_pr_producao SET [DT_FimProducao] = GETDATE(), [id_status_Peca] = 1 WHERE ID = " & IDProduto
    
 	'EXECUTA COMANDO
   Err.Clear     
	Set rst = conn.Execute(SQL_Table)
	
	'TRATA ERROS
	If Err.Number <> 0 Then
		STD_Erro "conn.Execute",strFuncName & " 105"
		Err.Clear
		'Close data source
		conn.close
		Set conn = Nothing
		Set rst = Nothing 
		Exit Sub
	End If
End If

STD_Log  strFuncName & " Produto Saída Ok (" & Barcode & ")"

'Close data source
conn.close

Set rst = Nothing
Set conn = Nothing

End Sub