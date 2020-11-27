Sub AtualizaTagContagem()
'Tip: '********* Atualiza Tag Contadores *************
' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:


Dim NomeTag, TagIndex, SQL_Table, rst, conn

On Error Resume Next

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
conn.Open "Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP" 'DSN= Name of the ODBC database

'TRATA ERROS
If Err.Number <> 0 Then
	STD_Erro "conn.Open Provider=MSDASQL;Initial Catalog=SNP_PI1846_7oEixo_Banco;DSN=Database_1;uid=SNP;pwd=SNP", "PreencheCBTags"
	Err.Clear
	Set conn = Nothing 
	Exit Sub
End If



TagIndex = SmartTags("StrTag"& SmartTags("cicloTagIndex"))

STD_Log "AtualizaTagContagem-Linha 32 - SmartTags(cicloTagIndex)="&SmartTags("cicloTagIndex")& ", SmartTags(StrTag)=" & SmartTags("StrTag"& SmartTags("cicloTagIndex"))

SQL_Table = "SELECT [tag_name] FROM [SNP_PI1846_7oEixo_Banco].[dbo].[tb_ana_tags]" &_
" where [tag_description] = '"& TagIndex & "'"

STD_Log "AtualizaTagContagem Linha 37 - " & SQL_Table

Err.Clear
'EXECUTA COMANDO SQL
Set rst = conn.Execute(SQL_Table)

'TRATA ERRO
If Err.Number <> 0 Then
	STD_Erro "Erro conn.Execute: #"  & SQL_Table, "AtualizaTagContagem"
	Err.Clear
End If

NomeTag = rst.fields(0).value
If IsDate(SmartTags("Field_2")) Then SmartTags("Field_1") = Date Else SmartTags("Field_1") = SmartTags(NomeTag)
SmartTags("TagDescricaoContador")=NomeTag


'Close data source
conn.close



Set rst = Nothing
Set conn = Nothing


SmartTags("TITULO-MsgZerarContagem")=  "Zerar Contador Estágio " & SmartTags("cicloEstagioIndex")
SmartTags("MSG_ZERAR_CONTAGEM")="Zerar " & TagIndex & " ?"


End Sub