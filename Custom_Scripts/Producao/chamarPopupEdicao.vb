Sub chamarPopupEdicao()
On Error Resume Next

'Just work if the PLCTag is Boolean
If SmartTags("Ultimo_WWID") <> "" Then
	Call ShowPopupScreen("Edit_Prod",454,167,hmiOn,hmiBottom,hmiMedium)
    	SmartTags("WWID_MSG") = "Valor Inserido"
    Else
    	SmartTags("WWID_MSG") = "Valor Inválido"
End If

'Error routine - Fehlerroutine
If Err.Number <> 0 Then
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Exit Sub
End If

End Sub