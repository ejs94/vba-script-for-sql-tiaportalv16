Function connect_MSSQL(ByRef pConnection, ByRef pDATABASE)
'Funcção para conectar o SQL Server, para isolar parte do código e permetir a reutilização

Dim strFuncName

strFuncName = "connect_MSSQL"

On Error Resume Next

'Inicializa Variável
Set pConnection = CreateObject("ADODB.Connection")

'ABRIR CONEXAO
'Para conexão local (usando a IHM)
pConnection.Open "DRIVER={SQL Server};" & _
	"SERVER=.\SQLEXPRESS;" & _
	"DATABASE=" & pDATABASE & ";" & _
	"UID=;PWD=;"

'Para conexão remota (usando Simulador Runtime e Docker)
'pConnection.Open "DRIVER={SQL Server};" & _
'	"SERVER=192.168.0.11;" & _
'	"DATABASE=" & pDATABASE & ";" & _
'   "UID=user;PWD=password;"

'TRATA ERROS
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": Erro ao Abrir Conexão."
	Err.Clear
	Set pConnection = Nothing
	connect_MSSQL = False
	Exit Function
End If

'Se chegou até aqui é pq conectou ok	
showLog strFuncName & ": Abriu Conexão"
connect_MSSQL = True
End Function