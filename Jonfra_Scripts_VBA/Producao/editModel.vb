Sub editModel()
'////////////////////////////////////////////////////////////////
' Permite editar os modelos dos blocos que estão sendo utilizados pelos blocos.
' 
' Todas as alterações são salvas em uma table do SQL Server.
' 
' INPUT NECESSÀRIOS: Modelo ID, STRING MODELO, NOME DO MODELO, DIAMETRO DA PEÇA
' INPUT OPCIONAIS: WWID do operador
'
' Created: 12-10-2020
' Version: v0.7
' Author:  EJS 
'////////////////////////////////////////////////////////////////
Dim strFuncName,Model_ID, SQL_Table, conn, rst
Dim pDATABASE, Reg_Edit_Table
Dim Modelo , NomeModelo, DiametroCamisa, TamanhoBloco

strFuncName = "editModel"

On Error Resume Next
'WWID para teste, porém ao acessar esse número um WWID será inserido.
SmartTags("Ultimo_WWID") = "editModel"

Model_ID = SmartTags("select_ID_Model")
Modelo = SmartTags("edit_ModelString")
NomeModelo = SmartTags("edit_ModelNameString")
DiametroCamisa = SmartTags("edit_diametroCamisa")
TamanhoBloco = SmartTags("edit_tamanhoBloco")

'ABRIR CONEXAO
Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'ABRIR CONEXAO
'Para conexão local (usando a IHM)
conn.Open "DRIVER={SQL Server};" & _
	"SERVER=.\SQLEXPRESS;" & _
	"DATABASE=" & pDATABASE & ";" & _
	"UID=;PWD=;"

'Error routine - Rotina de Erro
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Set conn = Nothing
	Exit Sub
End If

'Caso a ID seja válida então poderá ocorrer a alteranção no Banco de Dados
If Model_ID <> 0 Then
    SQL_Table = "USE hmiDB; " &_
                " UPDATE ModelosBlocos" &_
                " SET ModeloString='" & Modelo & "'," &_
                " NomeModelo='" & NomeModelo & "'," &_
                " DiametroCamisa=" & DiametroCamisa & "," &_
                " TamanhoBloco=" & TamanhoBloco &_
                " WHERE Modelo_id=" & Model_ID & "; "
    
    Reg_Edit_Table =    "USE hmiDB; " &_
                        "INSERT INTO alterProducTable " &_
                        "(comando,wwid,dt_Alteracao) " &_
                        "Values('" & Replace(SQL_Table,"'","''") & "', '" & SmartTags("Ultimo_WWID") & "', " & "GETDATE()" & ");"

'Se o Debug estiver ativado
'showLog  strFuncName & " Select: " & SQL_Table
'EXECUTA COMANDO SQL
    Set rst = conn.Execute(SQL_Table)
    Set rst = conn.Execute(Reg_Edit_Table)
    showLog strFuncName & "Dados Atualizados"
    showLog "SQL Table: " & SQL_Table
    showLog "Reg Table: " & Reg_Edit_Table

End If

'TRATA ERRO
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": conn.Execute: " & SQL_Table
	Err.Clear
	'ENCERRA
	conn.close
	showLog strFuncName & ": Conexão com o MSSQL fechada"
	rst = Nothing
End If

'Fecha todas conexões
rst.close
conn.close
Set rst = Nothing
Set conn = Nothing

End Sub