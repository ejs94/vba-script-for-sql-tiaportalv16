Dim conn, connStrg, connStrg2, rst, SQL_Table

On Error Resume Next

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'For Local Connection
connStrg = "DRIVER={SQL Server};" & _
	"SERVER=.\SQLEXPRESS;" & _
	"DATABASE=" & sqlDATABASE & ";" & _
    "UID=;PWD=;"

'For Remote Connection
connStrg2 = "DRIVER={SQL Server};" & _
	"SERVER=192.168.0.11;" & _
	"DATABASE=" & sqlDATABASE & ";" & _
    "UID=user;PWD=password;"

'Open data source - Datenquelle öffnen
conn.Open connStrg

'Error routine - Fehlerroutine
If Err.Number <> 0 Then
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If

SQL_Table = "SELECT GETDATE();"

'Execute - Ausführen
Set rst = conn.Execute(SQL_Table)

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
	
    SmartTags("escrevenoPLC") = rst.Fields(0).Value
	
	rst.close 
Else
	ShowSystemAlarm "Dat_No. is not available"
End If

'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing

End Sub