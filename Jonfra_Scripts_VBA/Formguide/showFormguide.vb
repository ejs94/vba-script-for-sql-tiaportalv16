Sub showFormguide()
'Essa subrotina será responsavel para carregar os valores de produção na tela do formguide.
'Formguide é um papel onde os operadores preenchem a serial e quantidade de produção durantes os turnos.
'Será necessário realizar algumas queries no SQL Server

Dim strFuncName, conn, rst, pDATABASE, i, j, Inicial, Final
Dim SQL_Seriais, SQL_Conforme_NConfome, SQL_StartTime, SQL_EndTime
Dim tempConforme, tempNaoConforme, tempTotalTurno

strFuncName = "showFormguide"

pDATABASE = "hmiDB"


On Error Resume Next


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

''''''''''''''''''''''''''' STRINGS SQL ''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Query para buscar as Strings de todos os Turnos
SQL_Seriais = " USE hmiDB; " &_
                " SELECT S.dt_Saida,B.PNSerialString " &_
                " FROM RegEntradaBlocos AS B " &_
                " JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id " &_
                " LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id " &_
                " WHERE S.dt_Saida BETWEEN CAST(CONVERT(date,GETDATE()) AS varchar)+' "& SQL_StartTime &"' AND CAST(CONVERT(date,GETDATE()) AS varchar)+' "& SQL_EndTime &"'" &_
                " ORDER BY S.dt_Saida; "

'''Query para contar o número de peças que foram conforme ou não conforme durante a produção
SQL_Conforme_NConfome = " USE hmiDB;" &_
                                " SELECT" &_
                                " COUNT(CASE " &_
                                " WHEN opBB155 = 'Aprovada P1' or opBB155 = 'Aprovada P2'" &_
                                " OR opBB165 = 'Aprovada P1' or opBB165 = 'Aprovada P2'" &_
                                " OR opBB165 = 'Aprovada P1' or opBB165 = 'Aprovada P2'" &_
                                " OR opBB175 = 'Aprovada'" &_
                                " AND LTRIM(opBB185) = 'Aprovada'" &_
                                " AND NOT (opBB155 = 'Refugo P1' or opBB155 = 'Refugo P2' or opBB165 = 'Refugo P1' or opBB165 = 'Refugo P2' or opBB175 = 'Refugo' or LTRIM(opBB185) = 'Refugo')" &_
                                " THEN 1 END) As Conforme," &_
                                " COUNT(CASE WHEN opBB155 = 'Refugo P1' or opBB155 = 'Refugo P2'" &_
                                " or opBB165 = 'Refugo P1' or opBB165 = 'Refugo P2'" &_
                                " or opBB175 = 'Refugo'" &_
                                " or LTRIM(opBB185) = 'Refugo' THEN 1 END) As Nao_Conforme" &_
                                " FROM RegEntradaBlocos AS B" &_
                                " JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id" &_
                                " WHERE S.dt_Saida BETWEEN CAST(CONVERT(date,GETDATE()) AS varchar)+' "& SQL_StartTime &"' AND CAST(CONVERT(date,GETDATE()) AS varchar)+' "& SQL_EndTime &"';"


'''''''''''''''''' PRENCHE O CAMPO DE STRINGS DA IPC
''' TURNO 1 '''
Inicial = 7
Final = 8
'Para o Turno 1: 7h até 16h
j = 0
For j = 1 To 9
    SQL_StartTime = Inicial & ":00:00"
    SQL_EndTime = Final & ":00:00"
    Set rst = conn.Execute(SQL_Seriais)
    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        For i = 1 To 15
            If rst.EOF Then
                SmartTags("DB_Contador_Producao_Dados_Turno_1_SN_" & j & "{" & i & "}") = ""
		        Else 
                    SmartTags("DB_Contador_Producao_Dados_Turno_1_SN_" & j & "{" & i & "}") = rst.Fields(1).Value
                    rst.MoveNext
            End If
        Next
    End If
    Inicial = Inicial + 1
    Final = Final + 1
    rst.close
Next

'Para o Turno 1: 16h até 16h30
SQL_StartTime =  "16:00:00"
SQL_EndTime = "16:30:00"
Set rst = conn.Execute(SQL_Seriais)
If Not (rst.EOF And rst.BOF) Then
    rst.MoveFirst
    i = 0
    For i = 1 To 15
        If rst.EOF Then
            SmartTags("DB_Contador_Producao_Dados_Turno_1_SN_" & 10 &"{" & i & "}") = ""
            Else
                SmartTags("DB_Contador_Producao_Dados_Turno_1_SN_" & j & "{" & 10 & "}") = rst.Fields(1).Value
                rst.MoveNext
        End If
    Next
End If
rst.close









'''''''''''''''''''''''''''''''''''''NÚMERO PEÇAS CONFORME E NÂO CONFORME'''''''''''''''
'''Preenchendo os Campos de Cálculo de Conforme e Não Conforme Turno 1
Inicial = 7
Final = 8
'Para o Turno 1: 7h até 16h30
i = 0
For i = 1 To 10
    SQL_StartTime = Inicial & ":00:00"
    SQL_EndTime = Final & ":00:00"
    Set rst = conn.Execute(SQL_Seriais)
    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        If rst.EOF Then
            SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_OK{"& i &"}") = 0
            SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_OK{"& i &"}") = 0
            Else
                SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_LIB OP{"& i &"}") = rst.Fields(1).Value
                SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_LIB OP{"& i &"}") = rst.Fields(1).Value
        End If
    End If
    Inicial = Inicial + 1
    Final = Final + 1
    rst.close
Next


'''''''''''''''''SOMA O TOTAL DE PEÇAS''''''''''''''''''''''''
''' Soma Total do Turno 1
i = 0
tempConforme = 0
tempNaoConforme = 0
tempTotalTurno = 0
For i = 1 To 10
    tempConforme = tempConforme + SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_OK{"& i &"}")
    tempNaoConforme = tempNaoConforme + SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_LIB OP{"& i &"}")
Next
SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador Total OK") = tempConforme
SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador Total LIB OP") = tempNaoConforme
SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_Total_Turno") = SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador Total OK") + SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador Total LIB OP")
''' Soma Total do Turno 2
i = 0
tempConforme = 0
tempNaoConforme = 0
tempTotalTurno = 0
For i = 1 To 10
    tempConforme = tempConforme + SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador_OK{"& i &"}")
    tempNaoConforme = tempNaoConforme + SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador_LIB OP{"& i &"}")
Next
SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador Total OK") = tempConforme
SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador Total LIB OP") = tempNaoConforme
SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador_Total_Turno") = SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador Total OK") + SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador Total LIB OP")
''' Soma Total do Turno 3
i = 0
tempConforme = 0
tempNaoConforme = 0
tempTotalTurno = 0
For i = 1 To 6
    tempConforme = tempConforme + SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador_OK{"& i &"}")
    tempNaoConforme = tempNaoConforme + SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador_LIB OP{"& i &"}")
Next
SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador Total OK") = tempConforme
SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador Total LIB OP") = tempNaoConforme
SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador_Total_Turno") = SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador Total OK") + SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador Total LIB OP")


''''''''TRATA ERROS''''''''''''''''''''''''''''''''''
'TRATA ERRO
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": conn.Execute: " & SQL_Seriais
    ShowSystemAlarm strFuncName & ": conn.Execute: " & SQL_Conforme_NConfome
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