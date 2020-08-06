Sub S02_Create_new_table(ByRef Database_Name)
'////////////////////////////////////////////////////////////////
' en: Creating a new table in the SQL database 
' pt-br: Criando uma nova tabela em um banco de dados
' Created: 11-05-2020
' Version: v1.0
' Author:  EJS 
'////////////////////////////////////////////////////////////////

'Declaration of local tags - Deklaration von lokalen Variablen
Dim conn, connStrg, connStrg2, rst, SQL_Table
Dim szDatabase, szTableName, szName_1, szName_2, szName_3

szDatabase = SmartTags("szDatabase")
szTableName = SmartTags("szTableName")
szName_1 = SmartTags("szName_1")
szName_2 = SmartTags("szName_2")
szName_3 = SmartTags("szName_3") 

On Error Resume Next

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

connStrg = "Provider=MSDASQL;" & _
	"Initial Catalog=" & szDatabase & ";" & _
	"DSN="&Database_Name&"" 
'DSN= Name of the ODBC database - DSN= Name der ODBC-Datenbank

connStrg2 = "DRIVER={SQL Server};" & _
	"SERVER=192.168.88.129;" & _
	"DATABASE=" & szDatabase & ";" & _
    "UID=sa;PWD=My$eCurePwd123#;"

'Open data source - Datenquelle öffnen
conn.Open connStrg2

'Error routine - Fehlerroutine
If Err.Number <> 0 Then
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Set conn = Nothing
	Exit Sub
End If


'Create table - Tabelle erstellen
'Definition of SQL table - Definition der SQL-Tabelle
SQL_Table = "CREATE TABLE "& szTableName & " (Nr SMALLINT, " _
            & szName_1 & " CHAR(30), " & szName_2 & " SMALLINT, " _
            & szName_3 & " SMALLINT)"

'Execute - Ausführen
Set rst = conn.Execute(SQL_Table)

'Error routine - Fehlerroutine 
If Err.Number <> 0 Then
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	'Close data source - Datenquelle schließen
	conn.close
	Set conn = Nothing
	Set rst = Nothing
	Exit Sub
End If

'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing

End Sub