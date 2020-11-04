Sub RegistraSaida()
'////////////////////////////////////////////////////////////////
' Created: 12-10-2020
' Version: v0.7
' Author:  EJS 
'////////////////////////////////////////////////////////////////

'Rotina irá escrever no Banco de Dados após confirmação das opções dispobnibilizadas.
Dim strFuncName, SQL_Table, conn, rst
Dim pDATABASE
Dim SerialString, ModeloString , DataSerial
Dim MCH250, MCH350, G704, G516, Medir

strFuncName = "RegistraSaida"

On Error Resume Next

'Pegar os valores direto da tag de rastreabilidade
SerialString = SmartTags("DB102_RastreabilidadeBloco.E06_EsteiraSaida.QRCode.SerialString")
ModeloString = SmartTags("DB102_RastreabilidadeBloco.E06_EsteiraSaida.QRCode.ModeloString")
DataSerial = SmartTags("DB102_RastreabilidadeBloco.E06_EsteiraSaida.QRCode.DataString")
'Ultimo Part Status atualizados pelas operações nas máquinas

Select Case SmartTags("DB102_RastreabilidadeBloco.E06_EsteiraSaida.PartStatusOP.MCH250")
    Case 0
        MCH250 = "Lib. Operacao"
    Case 1
        MCH250 = "Trabalha"
    Case 2
        MCH250 = "Aprovada"
    Case 3
        MCH250 = "Refugo"
    Case 4
        MCH250 = "Medicao"
End Select

Select Case SmartTags("DB102_RastreabilidadeBloco.E06_EsteiraSaida.PartStatusOP.MCH350")
    Case 0
        MCH350 = "Lib. Operacao"
    Case 1
        MCH350 = "Trabalha"
    Case 2
        MCH350 = "Aprovada"
    Case 3
        MCH350 = "Refugo"
    Case 4
        MCH350 = "Medicao"
End Select

Select Case SmartTags("DB102_RastreabilidadeBloco.E06_EsteiraSaida.PartStatusOP.G704")
    Case 0
        G704 = "Lib. Operacao"
    Case 1
        G704 = "Trabalha"
    Case 2
        G704 = "Aprovada"
    Case 3
        G704 = "Refugo"
    Case 4
        G704 = "Medicao"
End Select

Select Case SmartTags("DB102_RastreabilidadeBloco.E06_EsteiraSaida.PartStatusOP.G516")
    Case 0
        G516 = "Lib. Operacao"
    Case 1
        G516 = "Trabalha"
    Case 2
        G516 = "Aprovada"
    Case 3
        G516 = "Refugo"
    Case 4
        G516 = "Medicao"
End Select

Select Case SmartTags("DB102_RastreabilidadeBloco.E06_EsteiraSaida.Operação.Medir")
    Case True
        Medir = "Sim"
    Case False
        Medir = "Nao"
End Select

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
        "INSERT INTO RegSaidaBlocos " &_
        " (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida) " &_
        " Values " &_
        " ((SELECT TOP 1 Bloco_id " &_
        " FROM RegEntradaBlocos " &_
        " WHERE PNSerialString='" & SerialString & "' " &_
        " ORDER BY Bloco_id DESC), " &_
        " '" & MCH250 & "', '" & MCH350 & "', '" & G704 & "', " & G516 & ", '" & Medir & "', GETDATE());"

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
	conn.close
	showLog strFuncName & ": Conexão com o MSSQL fechada"
	rst = Nothing
End If

'Fecha todas conexões
rst.close
conn.close
Set rst = Nothing
Set conn = Nothing




End Sub