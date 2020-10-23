Sub checkReqRetrabalho(ByRef PLCReq, ByRef IHMResposta)
'Esse foi um fix para criar um script que permita o PLC escrever de forma autonoma no SQLServer:
' 1. Crie uma Tag Booleana na Ladder
' 2. Exporte essa Tag para a IHM
' 3. Nas configurações dessa Tag da IHM, procure a aba de Evento e adicione esse script no Value Change.
' 4. Coloque a Tag novamente no PLCTag e voi lá.
'Write the code as of this position:

Dim strFuncName

strFuncName = "checkReqRetrabalho"

On Error Resume Next

'Just work if the PLCTag is Boolean
If PLCReq = True Then
	Call searchRetrabalho()
    IHMResposta = True
    Call ShowPopupScreen("Bloco_Retrabalho_Setup",454,167,hmiOn,hmiBottom,hmiMedium)
    SmartTags("DB110_IHM_IPC.EsteiraEntrada_ChamaPopUpRetrabalho") = True
ElseIf PLCReq = False Then
    IHMResposta = False
    SmartTags("DB110_IHM_IPC.EsteiraEntrada_ChamaPopUpRetrabalho") = False
End If

'Error routine - Fehlerroutine
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & " Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Exit Sub
End If


End Sub