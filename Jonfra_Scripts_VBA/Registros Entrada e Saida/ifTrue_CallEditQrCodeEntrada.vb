Sub ifTrue_CallEditQrCodeEntrada(ByRef PLCTag)
'Alterar propriedade da PLCTag para monitorar Ciclicamente
'Bloco_EstEnt_Input_QRCode
'Bloco_EstSaid_Input_QRCode
Dim strFuncName

strFuncName = "ifTrue_CallEditQrCodeEntrada"

On Error Resume Next

'Just work if the PLCTag is Boolean
If PLCTag = True Then
    showLog strFuncName & "PlcTag Value: " & PLCTag
	Call ShowPopupScreen("Bloco_EstEnt_Input_QRCode",454,167,hmiOn,hmiBottom,hmiMedium)
    Exit Sub
End If

showLog strFuncName & " : Requisição : " & PLCTag

'Error routine - Fehlerroutine
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & " Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Exit Sub
End If


End Sub