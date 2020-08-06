Sub ShowDesktop()
'Tip:
' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:

Dim shell
Set shell = CreateObject("WScript.Shell")  


'shell.SendKeys ("^{Esc}") 
'shell.SendKeys ("%{Tab}") 
'shell.SendKeys ("^{Esc}") 

shell.Run ("C:\Users\Public\ShowDesk.exe")
Set shell = Nothing

STD_Log Err.Number

End Sub