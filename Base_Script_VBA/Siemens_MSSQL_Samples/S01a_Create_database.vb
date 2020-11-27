Sub S01a_Create_database(ByRef Database_Name)
'////////////////////////////////////////////////////////////////
' en: Creating a new database
' pt-br: Cria um novo banco de dados
' Created: 20-05-2020
' Version: v1.0
' Author:  EJS
'////////////////////////////////////////////////////////////////

'Declaration of local tags - Declarando as tags que serão utilizadas no script
Dim conn, conStrg, szDatabase, rst, SQL_Table

szDatabase = SmartTags("szDatabase")

On Error Resume Next

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

conStrg = "DRIVER={SQL Server};" & _
	"SERVER=192.168.88.129;" & _
	"DATABASE=TestDB;" & _
	"UID=sa;" & _
	"PWD=My$eCurePwd123#;"
	
'Open data source - Abre a conexão com a fonte de dados
conn.Open conStrg

'Error routine - Rotina de Erro
If Err.Number <> 0 Then
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Set conn = Nothing
	Exit Sub
End If


'Create a database - Cria o database

SQL_Table = "CREATE DATABASE " & szDatabase


'Execute - Executa o comando SQL
Set rst = conn.Execute(SQL_Table)

'Error routine - Rotina de erro
If Err.Number <> 0 Then
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	'Close data source - Fecha a conexão com a fonta de dados
	conn.close
	Set conn = Nothing
	Set rst = Nothing
	Exit Sub
End If

'Close data source - Fecha a conexão com a fonta de dados
conn.close

Set conn = Nothing
Set rst = Nothing

End Sub

