Sub Teste_InsereBlocoProducao(ByRef pRequest, ByRef pBarcode)
'////////////////////////////////////////////////////////////////
' Insere dados de bloco chegando na mesa Estaco 03
' Created: 20171205
' Version: v0.1
' Author:  IMF 
'////////////////////////////////////////////////////////////////

'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strErro, PNBloco, TipoBloco

If Not pRequest Then
	Exit Sub
End If

On Error Resume Next

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	ShowSystemAlarm "Erro S01_InsereBlocoProducao, conn.Open: #" & Err.Number & " " & Err.Description
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If


'Writes a data record into a table
'Select data record of the table - Datensatz der Tabelle auswählen
SQL_Table = "SELECT * FROM tb_pr_producao WHERE Barcode = '"& pBarcode & "'"

'Execute - Ausführen
Set rst = conn.Execute(SQL_Table)

'TRATA ERROS
If Err.Number <> 0 Then
	ShowSystemAlarm "Erro S01_InsereBlocoProducao, conn.Execute:  #" & Err.Number & " " & Err.Description  & " " & SQL_Table
	Err.Clear
	'Close data source
	conn.close
	Set conn = Nothing
	Set rst = Nothing 
	Exit Sub
End If

If Not (rst.EOF And rst.BOF) And False Then 
	'SE CONSULTA RETORNAR ALGUM ITEM, SIGNIFICA QUE O BLOCO JÁ FOI LIDO ANTERIORMENTE!
 	ShowSystemAlarm "Bloco duplicado detectado!"
 	
 	'INSERE TABELA DE ERROS
 	'Guarda comando que deu erro
 	strErro = SQL_Table
 	SQL_Table = "INSERT INTO [dbo].[tb_pr_erros] ([SQL],[DT]) VALUES " & _
				"('" & strErro &"',GETDATE())"
         
	'EXECUTA COMANDO
	Set rst = conn.Execute(SQL_Table)
	
	'TRATA ERROS
	If Err.Number <> 0 Then
		ShowSystemAlarm "Erro S01_InsereBlocoProducao, conn.Execute:  #" & Err.Number & " " & Err.Description  & " " & SQL_Table
		Err.Clear
		'Close data source
		conn.close
		Set conn = Nothing
		Set rst = Nothing 
		Exit Sub
	End If
 	
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
    
    
	SQL_Table = "INSERT INTO [dbo].[tb_pr_producao] ([Data],[Barcode],[DT_InicioProducao],[DT_FimProducao],[Status_Peca], [Part_Number], [Modelo]) VALUES " & _
				"(GETDATE(),'" & pBarcode &"',GETDATE(),null, 0, 12345678, 3)"
         
	'EXECUTA COMANDO
	Set rst = conn.Execute(SQL_Table)
	
	'TRATA ERROS
	If Err.Number <> 0 Then
		ShowSystemAlarm "Erro S01_InsereBlocoProducao, conn.Execute:  #" & Err.Number & " " & Err.Description  & " " & SQL_Table
		Err.Clear
		'Close data source
		conn.close
		Set conn = Nothing
		Set rst = Nothing 
		Exit Sub
	End If
End If

'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing

End Sub