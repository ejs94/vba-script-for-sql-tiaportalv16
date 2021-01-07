Sub EditProduction()
'Rotina irá escrever no Banco de Dados após alterações nas opções dispobnibilizadas.
Dim strFuncName,ProductionID, SQL_Table, conn, rst
Dim pDATABASE, Reg_Edit_Table
Dim OP10 , OP20, OP30, OP40, Inpec

pDATABASE = "hmiDB"
strFuncName = "selectEditPart"
SmartTags("Ultimo_WWID") = "Edit Production"

On Error Resume Next

ProductionID = SmartTags("Edit_ID_Value")

'ABRIR CONEXAO
Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
'Para conexão local (usando a IHM)
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


'Chaveamento para conversar com o text file do TIA PORTAL
Select Case SmartTags("Edit_BB165_Field")
    Case 0
        OP10 = "Lib. Operacao"
    Case 1
        OP10 = "Trabalha"
    Case 2
        OP10 = "Aprovada P1"
    Case 3
        OP10 = "Refugo P1"
    Case 4
        OP10 = "Medicao P1"
    Case 5
        OP10 = "Aprovada P2"
    Case 6
        OP10 = "Refugo P2"
    Case 7
        OP10 = "Medicao P2"
End Select

Select Case SmartTags("Edit_BB165_Field")
    Case 0
        OP20 = "Lib. Operacao"
    Case 1
        OP20 = "Trabalha"
    Case 2
        OP20 = "Aprovada P1"
    Case 3
        OP20 = "Refugo P1"
    Case 4
        OP20 = "Medicao P1"
    Case 5
        OP20 = "Aprovada P2"
    Case 6
        OP20 = "Refugo P2"
    Case 7
        OP20 = "Medicao P2"
End Select

Select Case SmartTags("Edit_BB175_Field")
    Case 0
        OP30 = "Lib. Operacao"
    Case 1
        OP30 = "Trabalha"
    Case 2
        OP30 = "Aprovada P1"
    Case 3
        OP30 = "Refugo P1"
    Case 4
        OP30 = "Medicao P1"
    Case 5
        OP30 = "Aprovada P2"
    Case 6
        OP30 = "Refugo P2"
    Case 7
        OP30 = "Medicao P2"
End Select

Select Case SmartTags("Edit_BB185_Field")
    Case 0
        OP40 = "Lib. Operacao"
    Case 1
        OP40 = "Trabalha"
    Case 2
        OP40 = "Aprovada P1"
    Case 3
        OP40 = "Refugo P1"
    Case 4
        OP40 = "Medicao P1"
    Case 5
        OP40 = "Aprovada P2"
    Case 6
        OP40 = "Refugo P2"
    Case 7
        OP40 = "Medicao P2"
End Select

Select Case SmartTags("Edit_Inspecao_Field")
    Case 0
        Inpec = "Nao"
    Case 1
        Inpec = "Sim"
End Select

'Caso a ID seja válida então poderá ocorrer a alteranção no Banco de Dados
If ProductionID <> 0 Then
    SQL_Table = "USE hmiDB; " &_
                " UPDATE RegSaidaBlocos" &_
                " SET opBB155='" & OP10 & "'," &_
                " opBB165='" & OP20 & "'," &_
                " opBB175='" & OP30 & "'," &_
                " opBB185='" & OP40 & "'," &_
                " inspecao='" & Inpec & "'" &_
                " WHERE Producao_id=" & ProductionID & "; "
    
    Reg_Edit_Table =    "USE hmiDB; " &_
                        "INSERT INTO alterProducTable " &_
                        "(comando,wwid,dt_Alteracao) " &_
                        "Values('" & Replace(SQL_Table,"'","''") & "', '" & SmartTags("Ultimo_WWID") & "', " & "GETDATE()" & ");"

'Se o Debug estiver ativado
'showLog  strFuncName & " Select: " & SQL_Table
'EXECUTA COMANDO SQL
    Set rst = conn.Execute(SQL_Table)
    Set rst = conn.Execute(Reg_Edit_Table)
    showLog strFuncName & ": Dados Atualizados"
    showLog strFuncName & ": SQL Table: " & SQL_Table
    showLog strFuncName & ": Reg Table: " & Reg_Edit_Table

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