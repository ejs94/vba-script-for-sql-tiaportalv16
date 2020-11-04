Sub RegistraEntrada()
'////////////////////////////////////////////////////////////////
' Created: 12-10-2020
' Version: v0.7
' Author:  EJS 
'////////////////////////////////////////////////////////////////

'Rotina irá escrever no Banco de Dados após confirmação das opções dispobnibilizadas.
Dim strFuncName, SQL_Table, conn, rst
Dim pDATABASE
Dim SerialString, ModeloString , DataSerial

strFuncName = "RegistraEntrada"

On Error Resume Next

'Pegar os valores direto da tag de rastreabilidade
SerialString = SmartTags("DB102_RastreabilidadeBloco.E01_EsteiraEntrada.QRCode.SerialString")
ModeloString = SmartTags("DB102_RastreabilidadeBloco.E01_EsteiraEntrada.QRCode.ModeloString")
DataSerial = SmartTags("DB102_RastreabilidadeBloco.E01_EsteiraEntrada.QRCode.DataString")

'ABRIR CONEXAO
Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'Opção para:
'Conexão local (usando a IHM)
conn.Open "DRIVER={SQL Server};" & _
	"SERVER=.\SQLEXPRESS;" & _
	"DATABASE=" & pDATABASE & ";" & _
	"UID=;PWD=;"

'Error routine - Rotina de Erro
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Set conn = Nothing
	Exit Sub
End If

'Caso a ID seja válida então poderá ocorrer a alteranção no Banco de Dados
If SerialString <> "" And ModeloString <> "" Then
    SQL_Table = "USE hmiDB; " &_
" INSERT INTO RegEntradaBlocos " &_
"    (PNSerialString,Modelo_id,DataString,dt_Entrada)" &_
" Values " &_
"    ( '" & SerialString & "', " &_
"        (SELECT Modelo_id" &_
"        FROM ModelosBlocos" &_
"        WHERE ModeloString='" & ModeloString & "')," &_
"        '" & DataSerial & "'," &_
"        GETDATE()); "

'Se o Debug estiver ativado
'showLog  strFuncName & " Select: " & SQL_Table
'EXECUTA COMANDO SQL
    Set rst = conn.Execute(SQL_Table)
    showLog strFuncName & "Dados Atualizados"
    showLog "SQL Table: " & SQL_Table

End If

'TRATA ERRO
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": conn.Execute: " & SQL_Table
	Err.Clear
	'ENCERRA
	showLog strFuncName & ": Conexão com o MSSQL fechada"
End If

'Fecha todas conexões
rst.close
conn.close
Set rst = Nothing
Set conn = Nothing


End Sub