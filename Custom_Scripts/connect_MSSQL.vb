Function connect_MSSQL(ByRef pConnection)

'Inicializa Variável
Set pConnection = CreateObject("ADODB.Connection")

'ABRIR CONEXAO
'Para conexão local (usando a IHM)
conn.Open "DRIVER={SQL Server};" & _
	"SERVER=.\SQLEXPRESS;" & _
	"DATABASE=" & sqlDATABASE & ";" & _
	"UID=;PWD=;"

'Para conexão remota (usando Simulador Runtime e Docker)
'conn.Open "DRIVER={SQL Server};" & _
'	"SERVER=192.168.0.11;" & _
'	"DATABASE=" & sqlDATABASE & ";" & _
'   "UID=user;PWD=password;"


'TRATA ERROS
If Err.Number <> 0 Then
	ShowSystemAlarm "Connect_MSSQL : Erro ao Abrir Conexão."
	Err.Clear
	Set pConnection = Nothing
	connect_MSSQL = False
	Exit Function
End If

'Se chegou até aqui é pq conectou ok	
showLog "Connect_MSSQL : Abriu Conexão"
connect_MSSQL = True
End Function