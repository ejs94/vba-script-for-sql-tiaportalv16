eSub S04_ExportDatatoCSV(ByRef pFiltroPN, ByRef pFiltroDataInicial, ByRef pFiltroDataFinal)
'////////////////////////////////////////////////////////////////
' Exporta pesquisa para arquivo
' Ordenacao padrao = ID DESC
' Created: 20180710
' Version: v0.1
' Author:  IMF 
'////////////////////////////////////////////////////////////////

'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strAscDesc, i, j
Dim Cabecalho, fso, ObjFile, StrFileName, strFuncName, Linha, Dados, SqlDados

On Error Resume Next

strFuncName = "S04_ExportDatatoCSV"

'ABRIR CONEXAO
If Not ConnectDB(conn) Then
	Exit Sub
End If


'Export sempre na ordem de ID, ASCendente

'Busca
STD_Log "Chamando Select"
Set rst = S99_PesquisaSQL("ID", True, pFiltroPN, pFiltroDataInicial, pFiltroDataFinal, conn)
	

If Not (rst.EOF And rst.BOF) Then 
	'RETORNOU COM DADOS VÁLIDOS, PREENCHE TAGS:
	
	rst.MoveFirst 'PRIMEIRO DADO RECEBIDO 
	
	'ZERA ITERADOR
	j=0
	
	
	
	
	
	'Nome do Arquivo
	StrFileName = Right(Year(Date),2) & Right("0" & Month(Date),2) & Right("0" & Day(Date),2) &"-"&  _
	Right("0"& Hour(Time),2) & Right("0"& Minute(Time),2) & Right("0"& Second(Time),2)&".csv"

	SmartTags("MSG_FILENAME")="Arquivo_" & StrFileName
	StrFileName = "D:\Arquivos\Arquivo_" & StrFileName
 
	'Cabeçalho	
	Cabecalho = ""
	For j = 0 To (rst.Fields.Count - 1)'VERIFICA QUANTIDADE DE ELEMENTOS NA TABELA
		Cabecalho = Cabecalho & rst.Fields(j).Name & ";" 
	Next		

	Set fso = CreateObject("Scripting.FilesyStemObject")
	Set ObjFile= fso.CreateTextFile(StrFileName,True)



	If Err.Number <> 0 Then
		STD_Erro StrFileName,strFuncName
		Err.Clear
	End If



	HmiRuntime.Trace("VB-Script: Write file: " & StrFileName & vbCrLf)



	If Err.Number <> 0 Then
		STD_Erro StrFileName,strFuncName
		Err.Clear
	End If
	
	ObjFile.WriteLine Cabecalho
	HmiRuntime.Trace(Cabecalho & vbCrLf)


	rst.MoveFirst 'VOLTA AO PRIMEIRO DADO RECEBIDO 



	
	Do
		Linha=""
		For j=0 To 8
			SqlDados=rst.Fields(j).Value
			If j=1 Then SqlDados=STD_DateISO2Date(SqlDados)
			If j=3 Or j=4 Then SqlDados=STD_DateISO2DateTime(SqlDados)
			SqlDados=Replace(SqlDados,Chr(13),"") 'Retira o Caracter (ENTER) do código 
			Linha = Linha & SqlDados & ";"
		Next
		Dados = Dados & Linha & Chr(13)
	
		rst.MoveNext
	Loop Until rst.EOF

	rst.close 
	
	ObjFile.WriteLine Dados
	HmiRuntime.Trace(Dados & vbCrLf)
Else
	STD_Log  "DADOS RETORNARAM VAZIOS!"
	
End If


'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing





ObjFile.Close




End Sub