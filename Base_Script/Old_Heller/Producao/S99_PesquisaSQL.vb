Function S99_PesquisaSQL(ByRef pOrdem, ByRef pAscendente, ByRef pFiltroPN, ByRef pFiltroDataInicial, ByRef pFiltroDataFinal, ByRef pConnection)
'SELECIONA DADOS E ORDENA DE ACORDO COM PARAMETRO

' SELECT prod.*, st.descricao_status_peca
'  FROM [SNP_PI1846_7oEixo_Banco].[dbo].[tb_pr_producao] AS prod
'  INNER Join [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_status_peca] AS st On prod.Status_Peca = st.id_status_peca
'  WHERE prod.Barcode LIKE '%filtro%'
'	AND prod.Data BETWEEN 'dataINI' AND 'dataFIM'
'  ORDER BY prod.ID DESC

Dim rst, SQL_Table, strAscDesc, strFuncName, Teste

strFuncName = "S99_PesquisaSQL"

'Verifica se devemos pedir ordem ASCendente ou DESCendente
If pAscendente Then
	strAscDesc = "ASC"
Else
	strAscDesc = "DESC"
End If

SQL_Table = "SELECT prod.*, st.descricao_status_peca, mod.descricao_modelo " & _
			"FROM [SNP_PI1846_7oEixo_Banco].[dbo].[tb_pr_producao] AS prod " & _
			"INNER JOIN [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_status_peca] AS st ON prod.id_status_peca = st.id_status_peca " &_
			"INNER JOIN [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_modelo] AS mod ON prod.id_modelo = mod.id_modelo "

			
'Filtro de datas
SQL_Table = SQL_Table & "WHERE prod.Data BETWEEN '" & STD_DT2DateISO(pFiltroDataInicial) & "' AND '" & STD_DT2DateISO(pFiltroDataFinal) & "'  "

			
'Verifica se foi pedido filtro de PN
If pFiltroPN <> "" And Not SmartTags("CmdUpDown")  Then
	SQL_Table = SQL_Table & "AND prod.Barcode LIKE '%" & pFiltroPN & "%' "
End If	
	
SmartTags("CmdUpDown")=False		

			
'Ordena
SQL_Table = SQL_Table & "ORDER BY prod." & pOrdem & " " & strAscDesc

STD_Log "Select:" & SQL_Table

'EXECUTA COMANDO SQL
Set rst = pConnection.Execute(SQL_Table)

			
'TRATA ERRO
If Err.Number <> 0 Then
	STD_Erro "conn.Execute: " & SQL_Table, strFuncName
	
	'ENCERRA
	pConnection.close
	rst = Nothing
End If

STD_Log "Retornando ResultSet"

'Retorna Resultset da pesquisa
Set S99_PesquisaSQL = rst

End Function