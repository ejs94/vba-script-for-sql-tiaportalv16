Sub AtualizaValor(ByRef nData)
'Tip:
' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:
Dim nDate

nDate=SmartTags("NewDate")

If nData=1 Then 
	SmartTags("nFiltroDataInicial")= nDate
	If SmartTags("nFiltroDataInicial") > SmartTags("nFiltroDataFinal") Then SmartTags("nFiltroDataFinal")=SmartTags("nFiltroDataInicial")		
End If	
	
If nData=2 Then
	SmartTags("nFiltroDataFinal")= nDate
	If SmartTags("nFiltroDataInicial") > SmartTags("nFiltroDataFinal") Then SmartTags("nFiltroDataInicial")= SmartTags("nFiltroDataFinal")	
End If
	
If nData=3 Then
	SmartTags("eventDay")=Day(nDate)
	SmartTags("eventMonth")=Month(nDate)
	SmartTags("eventYear")=Year(nDate)
	SmartTags("DataPlanejada")=nDate
End If

End Sub