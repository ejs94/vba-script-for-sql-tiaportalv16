Sub ChangeUser()
'Tip:
' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:

Dim User


User=SmartTags("UserName")

STD_Log "User=" & User

If User="Snap" Then
	SmartTags("SnapUser")=True
	SmartTags("Debug")=True
Else
	SmartTags("SnapUser")=False
	SmartTags("Debug")=False
End If

End Sub