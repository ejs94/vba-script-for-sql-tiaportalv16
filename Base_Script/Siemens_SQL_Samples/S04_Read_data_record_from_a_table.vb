Sub S04_Read_data_record_from_a_table(ByRef Database_Name)
'////////////////////////////////////////////////////////////////
' en: The script reads the indicated data record
' pt-br: Esse script le um registro de dado indicado
' Created: 08-05-2020
' Version: v1.0
' Author:  EJS
'////////////////////////////////////////////////////////////////

'Declaration of local tags - Deklaration von lokalem Variablen
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


'Reads the data records of the SQL table
'Select data record of the table
SQL_Table = "SELECT * FROM " & szTableName & " WHERE Nr = " & nDat_No  '* = Alle Daten ' * = all data

'Execute
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
	
If Not (rst.EOF And rst.BOF) Then 
	'Compare if "End of File" or "Begin of File" exists, if not the pointer will be reset to the first entry
	
	rst.MoveFirst 'reset to 1st entry
	

	SmartTags("nDatNoTable") = rst.Fields(0).Value
	SmartTags("nValue_1") = rst.Fields(1).Value
	SmartTags("nValue_2") = rst.Fields(2).Value
	SmartTags("nValue_3") = rst.Fields(3).Value
	SmartTags("szName_1") = rst.Fields(1).Name
	SmartTags("szName_2") = rst.Fields(2).Name
	SmartTags("szName_3") = rst.Fields(3).Name
	
	rst.close 
Else
	ShowSystemAlarm "Dat_No. is not available"
End If

'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing

End Sub