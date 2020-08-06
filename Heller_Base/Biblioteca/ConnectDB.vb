Function ConnectDB(ByRef pConnection)

'Inicializa Variável
Set pConnection = CreateObject("ADODB.Connection")

'ABRIR CONEXAO
pConnection.Open GetConnectionString()

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "Erro ao Abrir Conexão.", "ConnectDB"
	
	Set pConnection = Nothing 
	ConnectDB = False
End If

'Se chegou até aqui é pq conectou ok	
STD_Log "Abriu Conexão"
ConnectDB = True
End Function