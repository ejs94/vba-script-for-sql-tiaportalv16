Sub setDateNow()

SmartTags("pFiltroDataFinal") = Now
SmartTags("pFiltroDataInicial") = Now - 3

On Error Resume Next

If Err.Number <> 0 Then
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Exit Sub
End If

End Sub