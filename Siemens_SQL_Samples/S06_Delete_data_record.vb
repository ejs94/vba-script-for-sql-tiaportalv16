Sub S06_Delete_data_record(ByRef Database_Name)
'////////////////////////////////////////////////////////////////
' en: Delete a data record in the SQL database 
' pt-br: Delete um registro de dados na tabela do Banco de Dados
' Created: 11-05-2020
' Version: v1.0
' Author:  EJS
'////////////////////////////////////////////////////////////////

'Declaration of local tags - Deklaration von lokalen Variablen
Dim conn, connStrg, connStrg2, rst, SQL_Table
Dim szDatabase, szTableName, nDat_No

szDatabase = SmartTags("szDatabase")
szTableName = SmartTags("szTableName")

nDat_No = SmartTags("nDat_No")


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


'Delete data record - Datensatz aus Tabelle löschen
SQL_Table = "DELETE FROM " & szTableName &" WHERE Nr = " & nDat_No  

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

ShowSystemAlarm "Data record was deleted"

'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing

End Sub