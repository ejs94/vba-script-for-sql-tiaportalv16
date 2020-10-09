Sub ChamarProgramaWindows(ByRef pComandoCaminho, ByRef pTituloJanelaAltTab, ByRef pComandoParams)

Dim shell
Set shell = CreateObject("WScript.Shell")  

If shell.AppActivate(pTituloJanelaAltTab)Then  
    shell.AppActivate pTituloJanelaAltTab 
Else 
    shell.Run pComandoCaminho
    StartProgram pComandoCaminho,pComandoParams, hmiShowMinimizedAndInactive, hmiYes
End If

Set shell = Nothing
End Sub