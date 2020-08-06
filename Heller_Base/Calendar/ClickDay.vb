Sub ClickDay(ByRef nDay)
'Tip:
' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:

Dim Data

Data=SmartTags("FirstDate")

STD_Log "ClickDay - FirstDate=" & Data & ", nDay=" & nDay

SmartTags("NewDate")= Data + nDay - 1

STD_Log "ClickDay - NewDate=" & SmartTags("NewDate")

Call SetData(SmartTags("NewDate"))

End Sub