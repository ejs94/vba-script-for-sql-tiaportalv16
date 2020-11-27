Sub SetData(ByRef nDate)
'Tip: Call Calendar 
' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:

Dim c, Data, Dia, Mes, sMes, Ano, sWeekDay, iWeekDay, cDay
Dim PosDate, PosDay

On Error Resume Next
Err.Clear

If SmartTags("ExitSub") Then Exit Sub
SmartTags("ExitSub")=True

Dia=Day(nDate)
Mes=Month(nDate)
Ano=Year(nDate)

SmartTags("CalendarMes")= Mes
SmartTags("CalendarAno")= Ano-2000

PosDay= nDate

STD_Log "SetData 27 nDate "& nDate & ", Dia " & Dia & ", Mes " & Mes & ", Ano " & Ano 


iWeekDay = Weekday(nDate)

Select Case iWeekDay
	Case 1: sWeekDay="Domindo"
	Case 2: sWeekDay="Segunda-feira"
	Case 3: sWeekDay="Terça-feira"
	Case 4: sWeekDay="Quarta-feira"
	Case 5: sWeekDay="Quinta-feira"
	Case 6: sWeekDay="Sexta-feira"
	Case 7: sWeekDay="Sábado"
End Select

Select Case Mes
	Case 1: sMes="janeiro"
	Case 2: sMes="fevereiro"
	Case 3: sMes="março"
	Case 4: sMes="abril"
	Case 5: sMes="maio"
	Case 6: sMes="junho"
	Case 7: sMes="julho"
	Case 8: sMes="agosto"
	Case 9: sMes="setembro"
	Case 10: sMes="outubro"
	Case 11: sMes="novembro"
	Case 12: sMes="dezembro"	
End Select

SmartTags("MSG_DATA")=sWeekDay & ", " & Dia & " de " & sMes & " de " & Ano  

SmartTags("AtualDate") = "01/"& Mes & "/"& Ano 'Primeiro Dia Do Mês
Data=SmartTags("AtualDate")

iWeekDay = Weekday(Data) 'integer Week day


If Err.Number <> 0 Then 
	STD_Erro "Linha 44 " & Err.Description,  "SetData"
	Err.Clear
End If



cDay = iWeekDay - 1 'calendar day start
If cDay < 3 Then cDay = iWeekDay + 6
PosDate = Data - cDay

SmartTags("FirstDate")=PosDate

STD_Log "Primeiro dia Do Mes 77 =" & Data & ", iWeekDay=" & iWeekDay &  ", nDate=" & nDate &  ", Posdate=" & PosDate


For c = 1 To 42
	SmartTags("dia"& c)=Day(PosDate)
	If Month(PosDate) = Month(nDate) Then
		If  Weekday(PosDate) = 1 Then SmartTags("CalendCorDia"&c) = 4 Else SmartTags("CalendCorDia"&c) = 1
    Else
    	If  Weekday(PosDate) = 1 Then SmartTags("CalendCorDia"&c) = 3 Else SmartTags("CalendCorDia"&c) = 0
    End If
	If PosDate = PosDay Then SmartTags("CalendCorDia"&c) = 2	
	PosDate = PosDate + 1
Next

If Err.Number <> 0 Then 
	STD_Erro "Linha 92 " & Err.Description,  "SetData"
	Err.Clear
End If

SmartTags("NewDate") = nDate

SmartTags("ExitSub")=False

End Sub