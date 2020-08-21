Function Escreve(nome,tipo,nivel,path,men)
	fnome=nome
	ftipo=tipo
	fnivel=nivel
	fpath=path
	fmen=men
	esp=""
	for i = 0 to fnivel 
		esp = "	" & esp
	next
	
	If (ftipo <> "DataFolder") Then
		 fmen = fmen & esp & fnome & vbCrLf
		 Escreve = fmen
	Else 
		fmen = fmen & esp & "*" & fnome & ":" & vbCrLf
		fnivel = fnivel + 1
		For each obj2 in Application.GetObject(fpath)		 	
		    Escreve = fmen+Escreve(fnome,ftipo,fnivel,fpath & "\" & fnome,fmen)
		Next
	End if
End Function



Sub Dicas_OnStartRunning()
nivel = 0
path = "Dados"
mensagem = "*Dados:" & vbCrLf
    
For each obj in Application.GetObject(path)
    mensagem = Escreve(obj.name,typename(obj),nivel,path+"\" & obj.name,mensagem)
next
msgbox mensagem
End Sub