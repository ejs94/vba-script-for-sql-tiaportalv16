Sub S08_Copy_table(ByRef Database_Name)
'////////////////////////////////////////////////////////////////
' en: Copy table containing the data records 
' pt-br: Copia a tabela contendo os dados recordados
' Created: 11-05-2020
' Version: v1.0
' Author:  EJS
'////////////////////////////////////////////////////////////////

'Declaration of local tags - Deklarartion von lokalen Variablen
Dim conn, rst, SQL_Table
Dim szDatabase, szTableName, szNewTableName

szDatabase = SmartTags("szDatabase")
szTableName = SmartTags("szTableName")
szNewTableName = SmartTags("szNewTableName")


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


'Copying the SQL table - Kopieren der SQL Tabelle
SQL_Table = " Select * Into " & szNewTableName & " From " & szTableName

'Execute - Ausführen
conn.execute(SQL_Table)

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