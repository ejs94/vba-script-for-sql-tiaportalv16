Sub setDateNow()

Dim strFuncName

strFuncName = "setDateNow"

SmartTags("pFiltroDataFinal") = Now
SmartTags("pFiltroDataInicial") = Now - Weekday(Now,2)
'showLog "CALCULO DA DATA: " & Weekday(Now,2)

On Error Resume Next

If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Exit Sub
End If



End Sub