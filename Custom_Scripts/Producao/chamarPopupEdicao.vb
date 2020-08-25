Sub chamarPopupEdicao()
'Rotina para permitir editar a peça apenas após input do WWID
'Essa rotinha chama a rotina Edit_Prod
'Made by: Estevao J Santos
On Error Resume Next

'Just work if the PLCTag is Boolean
If SmartTags("Ultimo_WWID") <> "" Then
	SmartTags("WWID_MSG") = "Valor Inserido"
	Call ShowPopupScreen("Edit_Prod",454,167,hmiOn,hmiBottom,hmiMedium)	
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