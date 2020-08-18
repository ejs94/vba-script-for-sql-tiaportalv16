Function queryProduction(ByRef pConnection)
'Query ao DB manipulada pelo VBA do Tia Portal

Dim rst, SQL_Table, strAscDesc, strFuncName, beginDate, endDate, search

'Essas Tags precisam ser criadas na IHM e associadas aos diplays de input
beginDate = SmartTags("pFiltroDataInicial")
endDate = SmartTags("pFiltroDataFinal")
beginDate = Year(beginDate) & "-" & Month(beginDate) & "-" & Day(beginDate) & " 23:59"
endDate = Year(endDate) & "-" & Month(endDate) & "-" & Day(endDate) & " 23:59"
search = SmartTags("pSearchPN")


On Error Resume Next

strFuncName = "queryProduction"

showLog "Entrei na query"

SQL_Table = "USE hmiDB; " &_
		"SELECT S.Producao_id, B.PNSerialString, M.ModeloString, M.NomeModelo, S.opBB155, S.opBB165, S.opBB175, S.opBB185, S.inspecao, B.dt_Entrada, S.dt_Saida " &_
		"FROM RegEntradaBlocos AS B " &_
    	"JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id " &_
    	"LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id "

'Filtro de datas
SQL_Table = SQL_Table & "WHERE S.dt_Saida BETWEEN '" & beginDate & "' AND '" & endDate & "'  "

			
'Verifica se foi pedido filtro de PN
If search <> "" Then
	SQL_Table = SQL_Table & "AND B.PNSerialString LIKE '%" & search & "%' "
End If	

'Ordena para padrão decrescente
SQL_Table = SQL_Table & " ORDER BY S.dt_Saida DESC, S.Producao_id DESC;"

'Se o Debug estiver ativado
showLog "Select: " & SQL_Table

'EXECUTA COMANDO SQL
Set rst = pConnection.Execute(SQL_Table)

			
'TRATA ERRO
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": conn.Execute: " & SQL_Table
	Err.Clear
	'ENCERRA
	pConnection.close
	showLog strFuncName & ": Conexão com o MSSQL fechada"
	rst = Nothing
End If

showLog "Retornando ResultSet"

'Retorna Resultset da pesquisa
Set queryProduction = rst



End Function