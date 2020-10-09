Sub updateTipoCargaPLC()
'Esse Sub Serve para preencher um array no PLC, inserindo todos os modelos no Banco de Dados
'Alterações na IPC irão afetar os modelos armazenados pelo PLC
' Criado por: EJS
'DECLARACAO DE TAGs
Dim pDATABASE, conn, rst, SQL_TABLE_COUNT, SQL_TABLE, i, strFuncName, num_linhas, connStringTest
Dim tempID, tempValue(49) 'Não é possivel passar um array inteiro para o PLC, deve ser utilizado uma tag por posição.

On Error Resume Next

'Cria um objeto para acesso ao SQL Server
Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'Configurações para a Sub
strFuncName = "updateTipoCargaPLC"
pDATABASE = "hmiDB"

SQL_TABLE_COUNT = "SELECT COUNT(Modelo_id) AS Quantidade FROM ModelosBlocos WHERE ModeloString != '' AND Modelo_id BETWEEN 1 AND 49;"

SQL_TABLE = " USE hmiDB;" &_
            " SELECT Modelo_id AS 'Tipo_Carga'" &_
                " , ModeloString" &_
            " FROM ModelosBlocos" &_
            " WHERE ModeloString != '' AND Modelo_id BETWEEN 1 AND 49" &_
            " ORDER BY Modelo_id;"

connStringTest = "DRIVER={SQL Server};" & _
	"SERVER=192.168.0.5;" & _
	"DATABASE=" & pDATABASE & ";" & _
	"UID=sa;PWD=engenharia8.@;"

'Inicia a SubRotina

'ABRIR CONEXAO COM SQL SERVER
conn.Open "DRIVER={SQL Server};" & _
	"SERVER=.\SQLEXPRESS;" & _
	"DATABASE=" & pDATABASE & ";" & _
	"UID=;PWD=;"

If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": Erro ao Abrir Conexão."
	Err.Clear
	Set conn = Nothing
	Exit Sub
End If

'PESQUISA BANCO DE DADOS
showLog strFuncName & ": Chamando a Query Count"
Set rst = conn.Execute(SQL_TABLE_COUNT)
	
'TRATA ERRO
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": conn.Execute: " & SQL_TABLE_COUNT
	Err.Clear
	'ENCERRA
	conn.close
	showLog strFuncName & ": Conexão com o MSSQL fechada"
	rst = Nothing
End If

If Not (rst.EOF And rst.BOF) Then
    rst.MoveFirst
    num_linhas = rst.Fields(0).Value
    showLog strFuncName & ": Numero de linhas: " & num_linhas
    Else
        rst.close
        conn.close
        Set rst = Nothing
        Set conn = Nothing
        Exit Sub
End If

'PESQUISA BANCO DE DADOS
showLog strFuncName & ": Chamando a Query"
Set rst = conn.Execute(SQL_TABLE)
	
'TRATA ERRO
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": conn.Execute: " & SQL_TABLE
	Err.Clear
	'ENCERRA
	conn.close
	showLog strFuncName & ": Conexão com o MSSQL fechada"
	rst = Nothing
End If

SmartTags("Debug") = 1

showLog strFuncName & ": Limpando Dados Anteriores"
For i=1 To 49
	'Apaga toda a tabela do Array
	SmartTags("TipoCarga_Modelo"&CStr(i)) = ""
Next

If Not (rst.EOF And rst.BOF) Then
	'RETORNOU COM DADOS VÁLIDOS, PREENCHE TAGS:
	showLog strFuncName & ": Encontrou Dados Válidos"
	
	rst.MoveFirst 'PRIMEIRO DADO DA TABEL

	For i = 1 To num_linhas
		tempID = CInt(rst.Fields(0).Value)
		tempValue(tempID) = rst.Fields(1).Value
        showLog strFuncName & ": For:" & i & ": ID:" & rst.Fields(0).Value & " Value: " & rst.Fields(1).Value
    	SmartTags("TipoCarga_Modelo"&CStr(tempID)) = tempValue(tempID)
		rst.MoveNext
	Next
	rst.close 
Else
	showLog strFuncName & ": DADOS RETORNARAM VAZIOS!"
End If

If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": Erro ao atualizar os dados do PLC!"
	Err.Clear
	'ENCERRA
	conn.close
	rst = Nothing
End If

'Fecha a conexão com o SQL Server
conn.close
Set rst = Nothing
Set conn = Nothing

End Sub