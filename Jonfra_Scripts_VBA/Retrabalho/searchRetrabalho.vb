Function searchRetrabalho()

'////////////////////////////////////////////////////////////////
' Dada uma serial, essa rotina irá checar nos registros do Banco de Dados,
' se o bloco registrou alguma entrada e quais op foram realizadas.
' 
' INPUT NECESSARIA: PNSerialString
'
' Created: 13-10-2020
' Version: v1.0
' Author:  EJS 
'////////////////////////////////////////////////////////////////



Dim strFuncName, SerialString, SQL_Table, conn, rst
Dim pDATABASE, Reg_Edit_Table
Dim Check_SerialString

strFuncName = "searchRetrabalho" 'Para facilitar debug

On Error Resume Next
'WWID para teste, porém ao acessar esse número um WWID será inserido.
showLog strFuncName & " Abriu a Função! "

Check_SerialString = SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.Serial_Busca")
SmartTags("DB110_IHM_IPC.EsteiraEntrada_PrecisaRetrabalho") = False 'Reseta para False

showLog strFuncName & ": Serial_Busca: " & Check_SerialString
showLog strFuncName & ": Estado inicial do PrecisaRetrabalho: " & SmartTags("DB110_IHM_IPC.EsteiraEntrada_PrecisaRetrabalho")

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
	Exit Function
End If

'Caso a ID seja válida então poderá ocorrer a alteranção no Banco de Dados
If Check_SerialString <> "" Then
    SQL_Table = "USE hmiDB; " &_
        " SELECT TOP 1 " &_
		" S.Producao_id, " &_
		" B.PNSerialString AS Serial, " &_
		" M.ModeloString AS Modelo, " &_
		" B.DataString AS 'Data Serial', " &_
		" S.opBB155 AS MCH250, " &_
		" S.opBB165 AS MCH350, " &_
		" S.opBB175 AS G705, " &_
		" S.opBB185 AS G516, " &_
		" B.dt_Entrada AS Entrada, " &_
		" S.dt_Saida AS Saida " &_
	" FROM RegEntradaBlocos AS B " &_
		" JOIN RegSaidaBlocos AS S " &_
		" ON B.Bloco_id = S.Bloco_id " &_
		" LEFT JOIN ModelosBlocos AS M " &_
		" ON B.Modelo_id = M.Modelo_id " &_
	" WHERE B.PNSerialString = '" & Check_SerialString & "' " &_
	" ORDER BY S.Producao_id DESC; "
		
'Se o Debug estiver ativado
'showLog  strFuncName & " Select: " & SQL_Table
'EXECUTA COMANDO SQL
    Set rst = conn.Execute(SQL_Table)
    showLog strFuncName & "Dados Atualizados"
    showLog "SQL Table: " & SQL_Table

End If

If Not (rst.EOF And rst.BOF) Then
    
    SmartTags("DB110_IHM_IPC.EsteiraEntrada_PrecisaRetrabalho") = True
    showLog strFuncName & ": PrecisaRetrabalho: " & SmartTags("DB110_IHM_IPC.EsteiraEntrada_PrecisaRetrabalho")
	rst.MoveFirst 'reset to 1st entry

    SmartTags("Retrabalho_Model") = rst.Fields(2).Value
    showLog strFuncName & ": Retrabalho_Model :" & SmartTags("Retrabalho_Model")

	SmartTags("Retrabalho_Ultima_DataHoraEntrada") = rst.Fields(8).Value
    showLog strFuncName & ": Retrabalho_Ultima_DataHoraEntrada :" & SmartTags("Retrabalho_Ultima_DataHoraEntrada")
    
	SmartTags("Retrabalho_Ultima_DataHoraSaida") = rst.Fields(9).Value
    showLog strFuncName & ": Retrabalho_Ultima_DataHoraSaida :" & SmartTags("Retrabalho_Ultima_DataHoraSaida")
    

    Select Case rst.Fields(4).Value
        Case "Lib. Operacao"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH250") = 0
        Case "Trabalha"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH250") = 1
        Case "Aprovada"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH250") = 2
        Case "Refugo"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH250") = 3
        Case "Medicao"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH250") = 4
        Case Else
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH250") = 0
    End Select
    showLog strFuncName & " : DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH250 : " & SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH250")

    Select Case rst.Fields(5).Value
        Case "Lib. Operacao"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH350") = 0
        Case "Trabalha"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH350") = 1
        Case "Aprovado"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH350") = 2
        Case "Refugo"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH350") = 3
        Case "Medicao"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH350") = 4
        Case Else
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH350") = 0
    End Select
    showLog strFuncName & " : DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH350 : " & SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.MCH350")

    Select Case rst.Fields(6).Value
        Case "Lib. Operacao"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G704") = 0
        Case "Trabalha"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G704") = 1
        Case "Aprovado"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G704") = 2
        Case "Refugo"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G704") = 3
        Case "Medicao"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G704") = 4
        Case Else
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G704") = 0
    End Select
    showLog strFuncName & " : DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G704 : " & SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G704")

    Select Case rst.Fields(7).Value
        Case "Lib. Operacao"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G516") = 0
        Case "Trabalha"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G516") = 1
        Case "Aprovado"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G516") = 2
        Case "Refugo"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G516") = 3
        Case "Medicao"
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G516") = 4
        Case Else
            SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G516") = 0
    End Select
    showLog strFuncName & " : DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G516 : " & SmartTags("DB110_IHM_IPC.EsteiraEntrada_DadosRetrabalho.G516")

	
	rst.close
    searchRetrabalho = True
    showLog strFuncName & " : searchRetrabalho : " & searchRetrabalho
Else
    SmartTags("DB110_IHM_IPC.EsteiraEntrada_PrecisaRetrabalho") = False
    showLog strFuncName & ": PrecisaRetrabalho: " & SmartTags("DB110_IHM_IPC.EsteiraEntrada_PrecisaRetrabalho")
	showLog strFuncName & ": Não existe entrada de blocos com essa serial."
    searchRetrabalho = False
    showLog strFuncName & " : searchRetrabalho : " & searchRetrabalho
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

End Function