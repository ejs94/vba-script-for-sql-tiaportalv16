Sub S01b_Delete_database(ByRef Database_Name)
'////////////////////////////////////////////////////////////////
' en: The script deletes the indicated SQL database 
' pt-br: Esse script deleta o banco de dados selecionado
' Created: 20-05-2020
' Version: v1.0
' Author:  EJS
'////////////////////////////////////////////////////////////////

'Declaration of local tags - Declarando as tags que serão utilizadas no script
Dim conn, connStrg,connStrg2,rst, SQL_Table
Dim szDatabase

szDatabase=SmartTags("szDatabase")

On Error Resume Next

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'Conexão local através do uso de HMI
connStrg = "Provider=MSDASQL;DSN="&Database_Name&"" 'DSN= name of the odbc-database - DSN= Name der ODBC-Datenbank

'Conexão remota com um servidor que não esteja na HMI
connStrg2 = "DRIVER={SQL Server};" & _
	"SERVER=192.168.88.129;" & _
	"DATABASE=TestDB;UID=sa;PWD=My$eCurePwd123#;"

'Open data source - Abre a conexão com a fonte de dados
conn.Open connStrg2 

'Error routine - Rotina de Erro
If Err.Number <> 0 Then
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Set conn = Nothing
	Exit Sub
End If


'Delete a database - Deleta o banco de dados
SQL_Table = "DROP DATABASE " & szDatabase

'Execute - Executa o comando SQL
Set rst = conn.Execute(SQL_Table)

'Error routine - Rotina de Erro
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