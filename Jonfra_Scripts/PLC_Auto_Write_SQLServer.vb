Sub PLC_Auto_Write_SQLServer(ByVal PLCTag)
'Esse foi um fix para criar um script que permita o PLC escrever de forma autonoma no SQLServer:
' 1. Crie uma Tag Booleana na Ladder
' 2. Exporte essa Tag para a IHM
' 3. Nas configurações dessa Tag da IHM, procure a aba de Evento e adicione esse script no Value Change.
' 4. Coloque a Tag novamente no PLCTag e voi lá.
'Write the code as of this position:

On Error Resume Next

'Just work if the PLCTag is Boolean
If PLCTag = True Then
	Call Write_data_record_into_SQLServer()
End If

'Error routine - Fehlerroutine
If Err.Number <> 0 Then
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Exit Sub
End If


End Sub