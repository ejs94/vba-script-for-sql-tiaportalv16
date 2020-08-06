Sub S03_Write_data_record_into_a_table(ByRef Database_Name)
'////////////////////////////////////////////////////////////////
' en: The script writes the indicated data record into a table
' pt-br: Esse script escreve um registro de dados passado uma tabela
' Created: 11-05-2020
' Version: v1.0
' Author:  EJS 
'////////////////////////////////////////////////////////////////

'Declaration of local tags - Deklaration von lokalen Variablen
Dim conn, connStrg, connStrg2, rst, SQL_Table
Dim szDatabase, szTableName, nDat_No, szName_1, szName_2, szName_3

szDatabase = SmartTags("szDatabase")
szTableName = SmartTags("szTableName")
nDat_No = SmartTags("nDat_No")
nValue_1 = SmartTags("nValue_1")
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


'Writes a data record into a table
'Select data record of the table - Datensatz der Tabelle auswählen
SQL_Table = "SELECT * FROM " & szTableName & " WHERE Nr = " & nDat_No  '* = all data

'Execute - Ausführen
Set rst = conn.Execute(SQL_Table)

'Error routine - Fehler Routine
If Err.Number <> 0 Then
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	'Close data source - Datenquelle schließen
	conn.close
	Set conn = Nothing
	Set rst = Nothing 
	Exit Sub
End If

If Not (rst.EOF And rst.BOF) Then 
	'Compare if "End of File" or "Begin of File" exists, if not the pointer will be reset to the first entry
	'Vergleich ob "End of File" oder "Begin of File" ist, wenn nicht wird der Zeiger auf den Ersten Eintrag zurueckgesetzt
 	ShowSystemAlarm "Dat No. exists already!"
	rst.close 
Else
	'Definition of data record - Definition des Datensatzes
	SQL_Table = "INSERT INTO "& szTableName & " VALUES ('" & nDat_No & _
	            "' , '" & nValue_1 & "' , '" & nValue_2 & _
	            "' , '" & nValue_3 & "')"
	'Insert the data reacord of the table - Datensatz in die Tabelle hinzufügen
	Set rst = conn.Execute(SQL_Table)
End If

'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing

End Sub