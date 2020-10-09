Sub S05_Edit_data_record(ByRef Database_Name)
'////////////////////////////////////////////////////////////////
' en: The script updates the indicated data record
' pt-br: Script para atualizar todos os dados indicados
' Created: 08-05-2020
' Version: v1.0
' Author:  EJS
'////////////////////////////////////////////////////////////////

'Declaration of local tags - Deklaration von lokalen Varaiblen
Dim conn, connStrg, connStrg2, rst, SQL_Table
Dim szDatabase, szTableName, nDat_No
Dim szName_1, szName_2, szName_3, nValue_1, nValue_2, nValue_3

szDatabase = SmartTags("szDatabase")
szTableName = SmartTags("szTableName")

nDat_No = SmartTags("nDat_No")

Set szName_1 = SmartTags("szName_1")
Set szName_2 = SmartTags("szName_2")
Set szName_3 = SmartTags("szName_3")
Set nValue_1 = SmartTags("nValue_1")
Set nValue_2 = SmartTags("nValue_2")
Set nValue_3 = SmartTags("nValue_3")


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


'Reads the data records of the SQL table
'Select data record of the table - Datensatz der Tabelle auswählen
SQL_Table = "SELECT * FROM " & szTableName & " WHERE Nr = " & nDat_No  ' * = all data

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
		
	rst.MoveFirst 'reset to 1st entry - auf 1. Eintrag zuruecksetzen 
	 
	szName_1.value = rst.Fields(1).Name
	szName_2.value = rst.Fields(2).Name
	szName_3.value = rst.Fields(3).Name 
	
	rst.close 
Else
	ShowSystemAlarm "Dat_No. is not available" 'dispaly at ALARM View, if no entry
End If

'Definition of data record - Definição de dados gravados
SQL_Table = "UPDATE "& szTableName & " Set " & szName_1 & " = '" & nValue_1 & "' , " & szName_2 & " = " & nValue_2 & " , " & szName_3 & " = " & nValue_3 & " WHERE Nr = " & nDat_No 

'Update data record
Set rst = conn.Execute(SQL_Table)

'Error routine
If Err.Number <> 0 Then
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	'Close data source
	conn.close
	Set conn = Nothing
	Set rst = Nothing
	Exit Sub
End If

'Close datab source
conn.close

Set rst = Nothing
Set conn = Nothing

End Sub