Sub selectEditProduction()
Dim strFuncName,ProductionID, SQL_Table, conn, rst
Dim pDATABASE

pDATABASE = "hmiDB"
strFuncName = "selectEditProduction"

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
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Set conn = Nothing
	Exit Sub
End If

If ProductionID <> 0 Then
    SQL_Table = " Select B.PNSerialString, M.NomeModelo, S.opBB155, S.opBB165, S.opBB175, S.opBB185, S.inspecao " &_ 
		" FROM RegEntradaBlocos AS B" &_
		" Join RegSaidaBlocos AS S On B.Bloco_id = S.Bloco_id" &_
		" Left Join ModelosBlocos AS M On B.Modelo_id = M.Modelo_id" &_
		" WHERE Producao_id= " & ProductionID & ";"

'Se o Debug estiver ativado
'showLog  strFuncName & " Select: " & SQL_Table
'EXECUTA COMANDO SQL
Set rst = conn.Execute(SQL_Table)

If Not (rst.EOF And rst.BOF) Then 
	'Compare if "End of File" or "Begin of File" exists, if not the pointer will be reset to the first entry
	
	rst.MoveFirst 'reset to 1st entry
    showLog rst.Fields(1).Value

    If IsNull(rst.Fields(0)) Then SmartTags("Edit_PN_Value") = "NULL"
	If IsNull(rst.Fields(1)) Then SmartTags("Edit_NomeModelo_Value") = "NULL"
	SmartTags("Edit_PN_Value") = rst.Fields(0).Value
	SmartTags("Edit_NomeModelo_Value") = rst.Fields(1).Value

    Select Case rst.Fields(2).Value
        Case "Aprovada"
            SmartTags("Edit_BB155_Field") = 1
        Case "Refugada"
            SmartTags("Edit_BB155_Field") = 2
        Case "Lib. Op"
            SmartTags("Edit_BB155_Field") = 3
        Case Else
            SmartTags("Edit_BB155_Field") = 0
    End Select

    Select Case rst.Fields(3).Value
        Case "Aprovada"
            SmartTags("Edit_BB165_Field") = 1
        Case "Refugada"
            SmartTags("Edit_BB165_Field") = 2
        Case "Lib. Op"
            SmartTags("Edit_BB165_Field") = 3
        Case Else
            SmartTags("Edit_BB165_Field") = 0
    End Select

    Select Case rst.Fields(4).Value
        Case "Aprovada"
            SmartTags("Edit_BB175_Field") = 1
        Case "Refugada"
            SmartTags("Edit_BB175_Field") = 2
        Case "Lib. Op"
            SmartTags("Edit_BB175_Field") = 3
        Case Else
            SmartTags("Edit_BB175_Field") = 0
    End Select

    Select Case rst.Fields(5).Value
        Case "Aprovada"
            SmartTags("Edit_BB185_Field") = 1
        Case "Refugada"
            SmartTags("Edit_BB185_Field") = 2
        Case "Lib. Op"
            SmartTags("Edit_BB185_Field") = 3
        Case Else
            SmartTags("Edit_BB185_Field") = 0
    End Select


    Select Case rst.Fields(6).Value
        Case "Sim"
            SmartTags("Edit_Inspecao_Field") = 1
        Case "Nao"
            SmartTags("Edit_Inspecao_Field") = 2
        Case Else
            SmartTags("Edit_Inspecao_Field") = 0
    End Select
	
	rst.close
Else
	showLog strFuncName & ": O dado não está disponivel e não pode ser editado."
End If
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
conn.close
Set rst = Nothing
Set conn = Nothing

End Sub