Function STD_DateISO2Date(ByRef pDTISO)
'Retorna Data no formato ISO (AAAA-MM-DD)
Dim strData
Dim sArray

strData = ""
'sArray=Split(pDTISO,"-")
'strData=sArray(2) & "/" &sArray(1) & "/" &sArray(0)

STD_DateISO2Date = Day(pDTISO)&"/" & Month(pDTISO)&"/" & Year(pDTISO)

End Function