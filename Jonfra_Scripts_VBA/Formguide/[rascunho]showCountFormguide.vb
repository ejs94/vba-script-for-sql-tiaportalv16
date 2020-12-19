Sub showCountFormguide()
'Essa subrotina será responsavel para carregar os valores de produção na tela do formguide.
'Formguide é um papel onde os operadores preenchem a serial e quantidade de produção durantes os turnos.
'Será necessário realizar algumas queries no SQL Server

Dim strFuncName, conn, rst, pDATABASE, i, j
Dim SQL_Seriais, SQL_Conforme_NConfome, SQL_StartTime, SQL_EndTime, SQL_DayOffset
Dim IHM_Linha, IHM_Turno
Dim tempConforme, tempNaoConforme, tempTotalTurno

strFuncName = "showCountFormguide"
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




''''''''''''''''''''''''

SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_OK{1}")

SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador_LIB OP{1}")


''''''''''''''''' Resolvendo o contador Conforme
Dim tempCountConforme, tempCountNaoConforme

Function CountConforme(ByRef opBB155, ByRef opBB165, ByRef opBB175, ByRef opBB185)
	If (opBB155 = "Aprovada P1" OR opBB155 = "Aprovada P2" OR opBB165 = "Aprovada P1" OR opBB165 = "Aprovada P2" OR opBB165 = "Aprovada P1" OR opBB165 = "Aprovada P2" OR opBB175 = "Aprovada" OR opBB185 = "Aprovada") AND NOT (opBB155 = "Refugo P1" or opBB155 = "Refugo P2" or opBB165 = "Refugo P1" or opBB165 = "Refugo P2" or opBB175 = "Refugo" or opBB185 = "Refugo") Then
		CountConforme = 1
		Else
			CountConforme = 0
	End If
End Function

tempCountConforme = 0
tempCountConforme = tempCountConforme + CountConforme(rst.Fileds(2).Value,rst.Fileds(3).Value,rst.Fileds(4).Value,rst.Fileds(5).Value)

SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_OK{1}") = tempConforme


'''''''''''''''''' Resolvendo o contador de Não Conforme
 Dim 

 Function CountNaoConforme(ByRef opBB155, ByRef opBB165, ByRef opBB175, ByRef opBB185)
 	If opBB155 = "Refugo P1" or opBB155 = "Refugo P2"  or opBB165 = "Refugo P1" or opBB165 = "Refugo P2" or opBB175 = "Refugo"  or opBB185 = "Refugo"  or (opBB155 = "Trabalha" AND opBB165 = "Trabalha" AND opBB175 = "Trabalha" AND opBB185 = "Trabalha") Then
		CountNaoConforme = 1
		Else
			CountNaoConforme = 0
	End If
End Function

tempCountNaoConforme = 0
tempCountNaoConforme = tempCountNaoConforme + CountNaoConforme(rst.Fileds(2).Value,rst.Fileds(3).Value,rst.Fileds(4).Value,rst.Fileds(5).Value)

SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador_LIB OP{1}") = tempCountNaoConforme

''''''''''''''''SOMA O TOTAL DE PEÇAS''''''''''''''''''''''''
''' Soma Total do Turno 1
tempTotalConforme = 0
tempTotalNaoConforme = 0
tempTotalTotalTurno = 0
For i = 1 To 10
    tempConforme = tempConforme + SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_OK{"& i &"}")
    tempNaoConforme = tempNaoConforme + SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_LIB OP{"& i &"}")
Next
SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador Total OK") = tempConforme
SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador Total LIB OP") = tempNaoConforme
SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_Total_Turno") = SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador Total OK") + SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador Total LIB OP")
''' Soma Total do Turno 2
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