Sub DBUpdate(ByRef PosIDCodigo)
'Tip: Atualiza Dados do Produto no banco de Dados
' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:

Dim c

If PosIDCodigo > 0 Then
	SmartTags("IDCodigo")=SmartTags("Value_ID_" & PosIDCodigo)
	SmartTags("Value_Data")=SmartTags("Value_Data_" & PosIDCodigo)
	SmartTags("Value_Barcode")=SmartTags("Value_Barcode_" & PosIDCodigo)
	SmartTags("Value_Status")=SmartTags("Value_Status_" & PosIDCodigo)
	SmartTags("Value_Modelo")=SmartTags("Value_Modelo_" & PosIDCodigo)
	SmartTags("Value_DTInicio")=SmartTags("Value_DTInicio_" & PosIDCodigo)
	SmartTags("Value_DTFim")=SmartTags("Value_DTFim_" & PosIDCodigo)
End If

For c = 1 To 12
	If c = PosIDCodigo Then SmartTags("CorSelTable_" & c)=1 Else SmartTags("CorSelTable_" & c)=0
	If c < 10 Then
		If SmartTags("BlocoTip" & c) = SmartTags("Value_Modelo") Then SmartTags("BlocoNr")= c
		STD_Log "Status" & c & "=" & SmartTags("Status" & c) & ", Value_Status=" & SmartTags("Value_Status")	
		If SmartTags("Status" & c) = SmartTags("Value_Status") Then SmartTags("StatusNr")= c
	End If
Next



End Sub