Sub ZerarContador()
'Tip:
' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:


Dim Pos_EstagioIndex, Pos_cicloTagIndex, TagOK, TagNOK, TagTemp 


Pos_EstagioIndex = SmartTags("cicloEstagioIndex")
Pos_cicloTagIndex = SmartTags("cicloTagIndex")

'DB_0100 - StatusMáquina_Reset_ContadorOK
'DB_0100 - StatusMáquina_Reset_ContadorNOK
'DB_0100 - StatusMáquina_Reset_TempoCiclo

Pos_EstagioIndex= "0" & Pos_EstagioIndex
Pos_EstagioIndex = Right(Pos_EstagioIndex,2)

TagOK = "DB_" & Pos_EstagioIndex & "00 - StatusMáquina_Reset_ContadorOK"
TagNOK = "DB_" & Pos_EstagioIndex & "00 - StatusMáquina_Reset_ContadorNOK"
TagTemp = "DB_"& Pos_EstagioIndex & "00 - StatusMáquina_Reset_TempoCiclo"

Select Case Pos_cicloTagIndex 
	Case 1: SmartTags("DB_" & Pos_EstagioIndex & "00 - StatusMáquina_Reset_ContadorOK")=True
	Case 2: SmartTags("DB_" & Pos_EstagioIndex & "00 - StatusMáquina_Reset_ContadorNOK")=True
	Case 3: 
		SmartTags("DB_" & Pos_EstagioIndex & "00 - StatusMáquina_Reset_ContadorOK")=True
		SmartTags("DB_" & Pos_EstagioIndex & "00 - StatusMáquina_Reset_ContadorNOK")=True
	Case 4: SmartTags("DB_" & Pos_EstagioIndex & "00 - StatusMáquina_Reset_TempoCiclo")=True
End Select

SmartTags("Field_1") = 0


End Sub