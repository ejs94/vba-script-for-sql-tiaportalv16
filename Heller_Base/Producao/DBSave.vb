'////////////////////////////////////////////////////////////////
' Atualiza Dados no Banco
' Created: 20180911
' Version: v0.1
' Author:  FCG 
'////////////////////////////////////////////////////////////////

'DECLARACAO DE TAGs
Dim conn, rst, SQL_Table, Data, BarCode, Modelo,nModelo, Status, nStatus, DTInicio, DTFim, IDProduto
Dim nFiltroPN,nFiltroDataInicial,nFiltroDataFinal, strFuncName
strFuncName = "***** CodeSave *****"


nFiltroPN=SmartTags("nFiltroPN")
nFiltroDataInicial=SmartTags("nFiltroDataInicial")
nFiltroDataFinal=SmartTags("nFiltroDataFinal")


On Error Resume Next



Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "conn.Open",strFuncName
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If


'BUSCA INFORMACOES DE RASTREABILIDADE NAS TAGS
IDProduto = SmartTags("IDCodigo")
Data = SmartTags("Value_Data")
BarCode = SmartTags("Value_Barcode")
Status = SmartTags("Value_Status")
Modelo = SmartTags("Value_Modelo")
DTInicio = SmartTags("Value_DTInicio")
DTFim = SmartTags("Value_DTFim")

Data=STD_DT2DateISO(Data)
If DTInicio="" Then DTInicio=DTFim
If DTFim="" Then DTFim=DTInicio
If DTInicio="" Then DTInicio=Date
If DTFim="" Then DTFim=Date
	
DTInicio=STD_DT2DateTimeISO(DTInicio)
DTFim=STD_DT2DateTimeISO(DTFim)


'********* Seleciona Status Peça *************
nStatus = SmartTags("StatusNr")-1


'****** Modelo *************
nModelo = SmartTags("BlocoNr")


SQL_Table = "UPDATE tb_pr_producao SET [Data] ='" & Data & "', [Barcode] ='" & BarCode  & _
"', [DT_InicioProducao] ='" & DTInicio & "', [DT_FimProducao] = '" & DTFim & "', [id_status_peca] = '" & nStatus & _
"', [id_modelo] ='" & nModelo & "' WHERE ID = " & IDProduto

STD_Log SQL_Table

Err.Clear
Set rst = conn.Execute(SQL_Table)
'
'
''TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "conn.Execute "& SQL_Table , strFuncName
	SmartTags("MSG_Titulo") = "Falha ao Salvar Dados"
	SmartTags("MSG") = "Verificar Dados"
	SmartTags("MSG_Cor") = 1
Else
	SmartTags("MSG_Titulo") = "Dados Atualizado"
	SmartTags("MSG") = "Arquivo Salvo"
	SmartTags("MSG_Cor") = 2		
	Call S02_SelectBlocosProducao("",0,nFiltroPN, nFiltroDataInicial, nFiltroDataFinal)
End If

Call ShowPopupScreen("MSG",350,290,hmiOn,hmiTop,hmiFast)


Call DBUpdate(1)


'Close data source
conn.close

Set rst = Nothing
Set conn = Nothing


