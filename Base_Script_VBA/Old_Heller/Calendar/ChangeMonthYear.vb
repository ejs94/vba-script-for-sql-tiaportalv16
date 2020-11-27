Sub ChangeMonthYear(ByRef nMes, ByRef nAno)
'Tip:
' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:

Dim Dia, Mes, Ano, Data

If SmartTags("ExitSub") Then Exit Sub

Data=SmartTags("NewDate")

STD_Log "ChangeMonth 13 "& Data

Mes=nMes
Dia=Day(Data)
Ano=nAno+2000

Data = Right("0" & Dia,2) & "/" & Right("0" & Mes,2) & "/" & Ano


SmartTags("NewDate")=Data

STD_Log "ChangeMonthYear Data 26 "& Data

Call SetData(SmartTags("NewDate"))

End Sub