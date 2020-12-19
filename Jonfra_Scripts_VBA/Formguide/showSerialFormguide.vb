Sub showSerialFormguide()
'Essa subrotina será responsavel para carregar os valores de produção na tela do formguide.
'Formguide é um papel onde os operadores preenchem a serial e quantidade de produção durantes os turnos.
'Será necessário realizar algumas queries no SQL Server

Dim strFuncName, conn, rst, pDATABASE
Dim SQL_Seriais, SQL_StartTime, SQL_EndTime, SQL_DayOffset
Dim IHM_Linha, IHM_Turno
Dim tempTotalConforme, tempTotalNaoConforme, i, j

strFuncName = "showSerialFormguide"
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

'''''''''''''''''''''''''''''''''''''' Definindo turno ''''''''''''''''''''''''''''

If (Time >= TimeValue("7:00:00") And Time < TimeValue("16:30:00")) Then
    IHM_Turno = 1
ElseIf (Time >= TimeValue("16:30:00") And Time < TimeValue("23:59:00")) Then
    IHM_Turno = 2
ElseIf (Time >= TimeValue("00:00:00") And Time < TimeValue("01:30:00")) Then
    IHM_Turno = 2
ElseIf (Time >= TimeValue("1:30:00") And Time < TimeValue("7:00:00")) Then
    IHM_Turno = 3
End If



showLog strFuncName & " Turno Atual: " & IHM_Turno & " Time: " & Time

'''''''''''''''''' PRENCHE O CAMPO DE STRINGS DA IPC
Select Case IHM_Turno
    Case 1
        ''' TURNO 1 '''
        IHM_Turno = 1

        'Para o Turno 1: 7h até 8h
        IHM_Linha = 1
        SQL_DayOffset = 0
        SQL_StartTime = "7:00:00"
        SQL_EndTime = "8:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)


        'Para o Turno 1: 8h até 9h
        IHM_Linha = 2
        SQL_DayOffset = 0
        SQL_StartTime = "8:00:00"
        SQL_EndTime = "9:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)


        'Para o Turno 1: 9h até 10h
        IHM_Linha = 3
        SQL_DayOffset = 0
        SQL_StartTime = "9:00:00"
        SQL_EndTime = "10:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)


        'Para o Turno 1: 10h até 11h
        IHM_Linha = 4
        SQL_DayOffset = 0
        SQL_StartTime = "10:00:00"
        SQL_EndTime = "11:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)


        'Para o Turno 1: 11h até 12h
        IHM_Linha = 5
        SQL_DayOffset = 0
        SQL_StartTime = "11:00:00"
        SQL_EndTime = "12:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)


        'Para o Turno 1: 12h até 13h
        IHM_Linha = 6
        SQL_DayOffset = 0
        SQL_StartTime = "12:00:00"
        SQL_EndTime = "13:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)


        'Para o Turno 1: 13h até 14h
        IHM_Linha = 7
        SQL_DayOffset = 0
        SQL_StartTime = "13:00:00"
        SQL_EndTime = "14:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)



        'Para o Turno 1: 14h até 15h
        IHM_Linha = 8
        SQL_DayOffset = 0
        SQL_StartTime = "14:00:00"
        SQL_EndTime = "15:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)



        'Para o Turno 1: 15h até 16h
        IHM_Linha = 9
        SQL_DayOffset = 0
        SQL_StartTime = "15:00:00"
        SQL_EndTime = "16:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)


        'Para o Turno 1: 16h até 16h30
        IHM_Linha = 10
        SQL_DayOffset = 0
        SQL_StartTime = "16:00:00"
        SQL_EndTime = "16:30:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)

    Case 2

        ''' TURNO 2 '''
        IHM_Turno = 2

        'Para o Turno 2: 16h30 até 17h
        IHM_Linha = 1
        SQL_DayOffset = 0
        SQL_StartTime = "16:30:00"
        SQL_EndTime = "17:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)



        'Para o Turno 2: 17h até 18h
        IHM_Linha = 2
        SQL_DayOffset = 0
        SQL_StartTime = "17:00:00"
        SQL_EndTime = "18:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)




        'Para o Turno 2: 18h até 19h
        IHM_Linha = 3
        SQL_DayOffset = 0
        SQL_StartTime = "18:00:00"
        SQL_EndTime = "19:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)



        'Para o Turno 2: 19h até 20h
        IHM_Linha = 4
        SQL_DayOffset = 0
        SQL_StartTime = "19:00:00"
        SQL_EndTime = "20:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)


        'Para o Turno 2: 20h até 21h
        IHM_Linha = 5
        SQL_DayOffset = 0
        SQL_StartTime = "20:00:00"
        SQL_EndTime = "21:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)




        'Para o Turno 2: 21h até 22h
        IHM_Linha = 6
        SQL_DayOffset = 0
        SQL_StartTime = "21:00:00"
        SQL_EndTime = "22:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)




        'Para o Turno 2: 22h até 23h
        IHM_Linha = 7
        SQL_DayOffset = 0
        SQL_StartTime = "22:00:00"
        SQL_EndTime = "23:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)



        'Para o Turno 2: 23h até 23h59
        IHM_Linha = 8
        SQL_DayOffset = 0
        SQL_StartTime = "23:00:00"
        SQL_EndTime = "23:59:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)



        'Para o Turno 2: 23h59 até 01h
        IHM_Linha = 9
        SQL_DayOffset = -1
        SQL_StartTime = "23:59:00"
        SQL_EndTime = "01:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)



        'Para o Turno 2: 01h até 1h30
        IHM_Linha = 10
        SQL_DayOffset = 0
        SQL_StartTime = "01:00:00"
        SQL_EndTime = "01:30:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)


    Case 3
        ''' TURNO 3 '''
        IHM_Turno = 3

        'Para o Turno 3: 1h30 até 2h
        IHM_Linha = 1
        SQL_DayOffset = 0
        SQL_StartTime = "1:30:00"
        SQL_EndTime = "2:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)


        'Para o Turno 3: 2h até 3h
        IHM_Linha = 2
        SQL_DayOffset = 0
        SQL_StartTime = "2:00:00"
        SQL_EndTime = "3:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)



        'Para o Turno 3: 3h até 4h
        IHM_Linha = 3
        SQL_DayOffset = 0
        SQL_StartTime = "3:00:00"
        SQL_EndTime = "4:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)


        'Para o Turno 3: 4h até 5h
        IHM_Linha = 4
        SQL_DayOffset = 0
        SQL_StartTime = "4:00:00"
        SQL_EndTime = "5:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)



        'Para o Turno 3: 5h até 6h
        IHM_Linha = 5
        SQL_DayOffset = 0
        SQL_StartTime = "5:00:00"
        SQL_EndTime = "6:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)



        'Para o Turno 3: 6h até 7h
        IHM_Linha = 6
        SQL_DayOffset = 0
        SQL_StartTime = "6:00:00"
        SQL_EndTime = "7:00:00"

        Call fillRowFormGuide( conn,  rst,  SQL_StartTime,  SQL_EndTime,  SQL_DayOffset,  IHM_Turno,  IHM_Linha)

End Select
''''''''''''''''''''''''''''''''''' Apaga Telas se não Estiver no Horário Certo'''''''''''''''''''''''

Select Case IHM_Turno
    Case 1
        For j = 1 To 10
            For i = 1 To 15
                SmartTags("DB_Contador_Producao_Dados_Turno_" & 2 & "_SN_" & j & "{" & i & "}") = ""
            Next
        Next
        For j = 1 To 6
            For i = 1 To 15
                SmartTags("DB_Contador_Producao_Dados_Turno_" & 3 & "_SN_" & j & "{" & i & "}") = ""
            Next
        Next
    Case 2
        For j = 1 To 6
            For i = 1 To 15
                SmartTags("DB_Contador_Producao_Dados_Turno_" & 3 & "_SN_" & j & "{" & i & "}") = ""
            Next
        Next
End Select


''''''''''''''''SOMA O TOTAL DE PEÇAS''''''''''''''''''''''''
''' Soma Total do Turno 1
tempTotalConforme = 0
tempTotalNaoConforme = 0
For i = 1 To 10
    tempTotalConforme = tempTotalConforme + SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_OK{"& i &"}")
    tempTotalNaoConforme = tempTotalNaoConforme + SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_LIB OP{"& i &"}")
Next
SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador Total OK") = tempTotalConforme
SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador Total LIB OP") = tempTotalNaoConforme
SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador_Total_Turno") = SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador Total OK") + SmartTags("DB_Contador_Producao_Dados_Turno_1_Contador Total LIB OP")
''' Soma Total do Turno 2
tempTotalConforme = 0
tempTotalNaoConforme = 0
For i = 1 To 10
    tempTotalConforme = tempTotalConforme + SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador_OK{"& i &"}")
    tempTotalNaoConforme = tempTotalNaoConforme + SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador_LIB OP{"& i &"}")
Next
SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador Total OK") = tempTotalConforme
SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador Total LIB OP") = tempTotalNaoConforme
SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador_Total_Turno") = SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador Total OK") + SmartTags("DB_Contador_Producao_Dados_Turno_2_Contador Total LIB OP")
''' Soma Total do Turno 3
tempTotalConforme = 0
tempTotalNaoConforme = 0
For i = 1 To 6
    tempTotalConforme = tempTotalConforme + SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador_OK{"& i &"}")
    tempTotalNaoConforme = tempTotalNaoConforme + SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador_LIB OP{"& i &"}")
Next
SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador Total OK") = tempTotalConforme
SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador Total LIB OP") = tempTotalNaoConforme
SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador_Total_Turno") = SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador Total OK") + SmartTags("DB_Contador_Producao_Dados_Turno_3_Contador Total LIB OP")


''''''''TRATA ERROS''''''''''''''''''''''''''''''''''
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Exit Sub
End If

'Fecha todas conexões
rst.close
conn.close
Set rst = Nothing
Set conn = Nothing


End Sub