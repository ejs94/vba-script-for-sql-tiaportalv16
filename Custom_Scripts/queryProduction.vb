Function queryProduction(ByRef pConnection)
'SELECIONA DADOS E ORDENA DE ACORDO COM PARAMETRO

' SELECT prod.*, st.descricao_status_peca
'  FROM [SNP_PI1846_7oEixo_Banco].[dbo].[tb_pr_producao] AS prod
'  INNER Join [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_status_peca] AS st On prod.Status_Peca = st.id_status_peca
'  WHERE prod.Barcode LIKE '%filtro%'
'	AND prod.Data BETWEEN 'dataINI' AND 'dataFIM'
'  ORDER BY prod.ID DESC

Dim rst, SQL_Table, strAscDesc, strFuncName


On Error Resume Next

strFuncName = "queryProduction"

showLog "Entrei na query"

SQL_Table = "USE hmiDB; " &_
		"SELECT S.Producao_id, B.PNSerialString, M.ModeloString, M.NomeModelo, S.opBB155, S.opBB165, S.opBB175, S.opBB185, S.inspecao, B.dt_Entrada, S.dt_Saida " &_
		"FROM RegEntradaBlocos AS B " &_
    	"RIGHT JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id " &_
    	"INNER JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id "

'Filtro de datas
'SQL_Table = SQL_Table & "WHERE prod.dt_Entrada BETWEEN '" & pFiltroDataInicial & "' AND '" & pFiltroDataFinal & "'  "

			
'Verifica se foi pedido filtro de PN
'If pFiltroPN <> "" And Not SmartTags("CmdUpDown")  Then
'	SQL_Table = SQL_Table & "AND prod.PNSerialString LIKE '%" & pFiltroPN & "%' "
'End If	
	
'SmartTags("CmdUpDown")=False		
			
'Ordena
'SQL_Table = SQL_Table & "ORDER BY prod." & pOrdem & " " & strAscDesc

SQL_Table = SQL_Table & " ;"

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