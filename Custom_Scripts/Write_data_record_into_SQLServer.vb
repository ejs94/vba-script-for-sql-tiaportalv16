Sub Write_data_record_into_SQLServer()
'Tip:
' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:

'Declaration of local tags - Deklaration von lokalen Variablen
Dim conn, connStrg, connStrg2, rst, SQL_Table
Dim sqlDATABASE, DataTable, Nome,Idade, Altura, Nascimento

sqlDATABASE = "hmiDB"
DataTable = "Dados"
Nome = SmartTags("DB_Geral_Dados_Nome")
Idade = SmartTags("DB_Geral_Dados_Idade")
Altura = SmartTags("DB_Geral_Dados_Altura")
Nascimento = SmartTags("DB_Geral_Dados_Nascimento")

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

'Writes a data record into a table
'Select data record of the table - Datensatz der Tabelle auswählen
SQL_Table = "INSERT INTO "& DataTable & " VALUES ('" & Nome & _
            "' , '" & Idade & "' , '" & Altura & _
            "' , '" & Nascimento & "')"

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

'Close data source - Datenquelle schließen
conn.close

Set rst = Nothing
Set conn = Nothing

'////////////////////////////////////////////////////////////////
' en: The script writes the indicated data record into a Database
' pt-br: Esse script escreve um registro de dados passado a um banco de dados
' Created: 11-05-2020
' Version: v1.
' Author:  EJS 
'////////////////////////////////////////////////////////////////

End Sub