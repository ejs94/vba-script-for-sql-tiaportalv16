Function STD_DT2DateISO(ByRef pDT)
Dim strData

strData = ""
strData = Year(pDT) & "-" & Month(pDT) & "-" & Day(pDT)
'Retorna Data no formato ISO (AAAA-MM-DD)
STD_DT2DateISO = strData

End Function