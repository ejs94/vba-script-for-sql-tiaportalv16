Sub S03_ResetPeriodoBusca(ByRef pTagDataInicial, ByRef pTagDataFinal)
'Zera Data Inicial para 01/01/2018
'Zera Data Final para hoje

Dim d, SQL_Table, conn, rst, j

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")


d=#1/1/18#
d=Date
If IsDate(d) Then
  	SmartTags(pTagDataInicial) = CDate(d)
	SmartTags(pTagDataFinal) = CDate(d)
End If


'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database


'********* Preenche Combo Modelo *************
SQL_Table = "Select * FROM [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_modelo]"

STD_Log SQL_Table

Err.Clear
Set rst = conn.Execute(SQL_Table)


'TRATA ERRO
If Err.Number <> 0 Then
	STD_Erro "Erro conn.Execute: #"  & SQL_Table, "S03_ResetPeriodoBusca"
Else
	rst.MoveFirst
	Do Until rst.eof
		j=j+1
		SmartTags("BlocoTip" & j)=rst.Fields(1).Value
		rst.MoveNext
	Loop
End If


'********* Preenche Combo Status *************
SQL_Table = "Select * FROM [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_status_peca]"

STD_Log SQL_Table

Err.Clear
Set rst = conn.Execute(SQL_Table)

j=0
'TRATA ERRO
If Err.Number <> 0 Then
	STD_Erro "Erro conn.Execute: #"  & SQL_Table, "S03_ResetPeriodoBusca"
Else
	rst.MoveFirst
	Do Until rst.eof
		j=j+1
		SmartTags("Status" & j)=rst.Fields(1).Value
		rst.MoveNext
	Loop
End If



'Close data source
conn.close

Set rst = Nothing
Set conn = Nothing



End Sub