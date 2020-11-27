Function STD_DateISO(ByRef pDT)
'Retorna Data no formato ISO (AAAA-MM-DD)
Dim strData

strData = ""
strData = Year(pDT) & "-" & Month(pDT) & "-" & Day(pDT)
STD_DateISO = strData

End Function