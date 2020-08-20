'////////////////////////////////////////////////////////////////
' Exporta pesquisa para arquivo
' Created: 2020-08-20
' Version: v1
' Author:  EJS 
'////////////////////////////////////////////////////////////////

'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, strAscDesc, i, j
Dim Cabecalho, fs, ObjFile, StrFileName, strFuncName, Linha, Dados, SqlDados, pDATABASE, datevar

pDATABASE = "hmiDB"
datevar = Year(Now) & "_" & Month(Now) & "_" & Day(Now) & "_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now)

On Error Resume Next

strFuncName = "exportProductionCSV"

'ABRIR CONEXAO
Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

conn.Open "DRIVER={SQL Server};" & _
	"SERVER=.\SQLEXPRESS;" & _
	"DATABASE=" & pDATABASE & ";" & _
	"UID=;PWD=;"

'Error routine - Rotina de Erro
If Err.Number <> 0 Then
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Set conn = Nothing
	Exit Sub
End If

'Export sempre na ordem de ID, ASCendente

'Busca
showLog "Chamando Select"
Set rst = queryProduction(conn,"ASC")
	

If Not (rst.EOF And rst.BOF) Then 
	'RETORNOU COM DADOS VÁLIDOS, PREENCHE TAGS:
	
	rst.MoveFirst 'PRIMEIRO DADO RECEBIDO 
	
	'ZERA ITERADOR
	j=0
	
	'Nome do Arquivo
	'SmartTags("MSG_FILENAME")="Arquivo_" & StrFileName
	StrFileName = "D:\ArquivosCSV\Arquivo_" & datevar & ".csv"
	StrFileName = "D:\ArquivosCSV\TMP.csv"
 
	'Cabeçalho	
	Cabecalho = ""
	For j = 0 To (rst.Fields.Count - 1)'VERIFICA QUANTIDADE DE ELEMENTOS NA TABELA
		Cabecalho = Cabecalho & rst.Fields(j).Name & "," 
	Next		

	Set fs = CreateObject("Scripting.FilesyStemObject")
	Set ObjFile= fs.CreateTextFile(StrFileName,True)

	'HmiRuntime.Trace("VB-Script: Write file: " & StrFileName & vbCrLf)

	ObjFile.WriteLine(Cabecalho)
	showLog "Cabecalho"

	rst.MoveFirst 'VOLTA AO PRIMEIRO DADO RECEBIDO 
	
	Do
		Linha=""
		For j=0 To (rst.Fields.Count - 1)
			SqlDados=rst.Fields(j).Value
            If j = 9 Or j = 10 Then SqlDados= "'" & rst.Fields(j).Value & "'" End If
			'SqlDados=Replace(SqlDados,Chr(13),"") 'Retira o Caracter (ENTER) do código 
			Linha = Linha & SqlDados & ","
		Next
		Dados = Dados & Linha & vbCrLf
	
		rst.MoveNext
	Loop Until rst.EOF

	rst.close
	
	ObjFile.WriteLine(Dados)
	'HmiRuntime.Trace(Dados & vbCrLf)

Else
	showLog  "DADOS RETORNARAM VAZIOS!"
	
End If


If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": Error #" & Err.Number & " " & Err.Description
	Err.Clear
	'Close data source
	conn.close
	Set conn = Nothing
	Set rst = Nothing 
	Exit Sub
End If    


'Close data source - Datenquelle schließen
conn.close
ObjFile.Close

Set rst = Nothing
Set conn = Nothing