'Precisa de uma tag pDT
Function STD_DT2DateTimeISO(ByRef pDT)
Dim strDateTime

strDateTime = ""
strDateTime = Year(pDT) & "-" & Month(pDT) & "-" & Day(pDT) & " " & Hour(pDT) & ":" & Minute(pDT) & ":" & Second(pDT)
'Retorna Data/Hora no formato ISO (AAAA-MM-DD HH:MM:SS)
STD_DT2DateTimeISO = strDateTime

End Function