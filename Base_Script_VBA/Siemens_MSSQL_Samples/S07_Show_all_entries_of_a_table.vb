Sub S07_Show_all_entries_of_a_table(ByRef Database_Name)
'////////////////////////////////////////////////////////////////
' en: Show all entries of a indicated table
' pt-br: Mostra todos os registro de uma tabela do banco de dados
' Created: 08-05-2020
' Version: v1.0
' Author:  EJS
'////////////////////////////////////////////////////////////////

'Declaration of local tags - Deklaration von lokalem Variablen
Dim conn, connStrg, connStrg2, rst, SQL_Table, i, j
'Smart Tags
Dim szDatabase, szTableName, nTab

szDatabase = SmartTags("szDatabase")
szTableName = SmartTags("szTableName")
nTab = SmartTags("nTab")

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


'Select all entries of a table - Alle Einträge der Tabelle selektieren
SQL_Table = "SELECT * FROM " & szTableName

'Execute - Ausführen
Set rst = conn.Execute(SQL_Table)

'Order the table by the first column - Tabelle nach der ersten Spalte sortieren
SQL_Table = "SELECT * FROM " & szTableName & " ORDER By " &  rst.Fields(0).Name  '* = Alle Daten ' * = all data

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
	'Vergleich ob "End of File" oder "Begin of File" ist, wenn nicht wird der Zeiger auf den Ersten Eintrag zurueckgesetzt
	'Compare if "End of File" or "Begin of File" exists, if not the pointer will be reset to the first entry
	
	rst.MoveFirst 'reset to 1st entry - auf 1. Eintrag zuruecksetzen 
	
	'Definition of local tags - Definiton von loklen Variablen
	j=0
	
	'Amount of the entries in the table - Anzahl der Tabelleneinträge
	Do
		j=j+1
		rst.MoveNext
	Loop Until rst.EOF
	
	rst.MoveFirst 'reset to 1st entry - auf 1. Eintrag zuruecksetzen Do
	
	'Selection with Arrow Buttons - Auswahl mit den Pfeil-Tasten
	If nTab >= j-6 Then
		SmartTag("nTab") = j-6
	End If
	If nTab < j-5 Then
		For i=1 To nTab
			rst.MoveNext
		Next
	End If
	If nTab <0 Then
		SmartTag("nTab") =0
	End If
	
	'Name of the columns - Name der Spalten
	SmartTags("szName_1") = rst.Fields(1).Name
	SmartTags("szName_2") = rst.Fields(2).Name
	SmartTags("szName_3") = rst.Fields(3).Name 
	
	For i=1 To 6	
		'Entries of the table - Einträge in die Tabelle
		If rst.EOF Then
			SmartTags("Value_" & i & "_0") = 0
			SmartTags("Value_" & i & "_1") = 0
			SmartTags("Value_" & i & "_2") = 0
			SmartTags("Value_" & i & "_3") = 0
		Else
			SmartTags("Value_" & i & "_0") = rst.Fields(0).Value
			SmartTags("Value_" & i & "_1") = rst.Fields(1).Value
			SmartTags("Value_" & i & "_2") = rst.Fields(2).Value
			SmartTags("Value_" & i & "_3") = rst.Fields(3).Value 
			rst.MoveNext
		End If
	Next
	
	rst.close 
Else
	ShowSystemAlarm "No entries are available."
End If

'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing

End Sub