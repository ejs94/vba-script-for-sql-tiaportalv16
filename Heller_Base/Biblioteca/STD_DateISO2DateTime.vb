Function STD_DateISO2DateTime(ByRef pDTISO)
Dim strData
Dim sArray

strData = ""
sArray=Split(pDTISO," ",3)
strData=sArray(0)

'Retorna Data no formato (DD/MM/YYYY)
STD_DateISO2DateTime = Day(strData)&"/" & Month(strData)&"/" & Year(strData)&" "& sArray(1)&" "& sArray(2)

End Function