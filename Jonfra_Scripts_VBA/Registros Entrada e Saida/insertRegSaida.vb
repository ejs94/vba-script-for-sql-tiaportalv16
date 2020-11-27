Sub insertRegSaida()
'////////////////////////////////////////////////////////////////
'      Caso seja necessário que o operador insira um registro de bloco
'   manualmente, essa função será chamada pela IPC para realizar as
'   alterações no banco de dados.
'////////////////////////////////////////////////////////////////

'INSERT INTO RegSaidaBlocos
'    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
'Values
'    ((SELECT TOP 1 Bloco_id
'        FROM RegEntradaBlocos
'        WHERE PNSerialString='FCF294'),
'        'Aprovada', 'Aprovada', 'Aprovada', 'Lib. Operacao', 'Sim', GETDATE());

'Rotina irá escrever no Banco de Dados após confirmação das opções dispobnibilizadas.
Dim strFuncName, Model_ID, SQL_Table, conn, rst

Dim pDATABASE, Reg_Edit_Table

Dim SerialString, MCH250, MCH350, G704, G516

strFuncName = "insertRegSaida"


On Error Resume Next
'WWID para teste, porém ao acessar esse número um WWID será inserido.
If IsNull(SmartTags("Ultimo_WWID")) Then
    SmartTags("Ultimo_WWID") = "TesteVB"
End If

'Recebendo valores das Tags que são existentes na IHM
'TODO: Fazer as tags na IHM com esses nomes e likar elas nas telas.
SerialString = SmartTags("edit_SerialString")
MCH250 = SmartTags("edit_MCH250")
MCH350 = SmartTags("edit_MCH350")
G704 = SmartTags("edit_G704")
G516 = SmartTags("edit_G516")


'ABRIR CONEXAO
Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

'Opção para:
'Conexão local (usando a IHM)
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
If ModeloString <> 0 And SerialString <> "" And DataString <> "" Then
    SQL_Table = "USE hmiDB; " &_
                " INSERT INTO RegSaidaBlocos" &_
                " (Bloco_id, opBB155, opBB165, opBB175, opBB185, inspecao, dt_Saida)" &_
                " Values (" & SerialString & ", " &_
                " (SELECT Modelo_id FROM ModelosBlocos WHERE ModeloString = '" & ModeloString & "'), " &_
                "'" & DataString & "', " &_
                " GETDATE() );"

" INSERT INTO RegSaidaBlocos" &_
" (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)" &_
" Values(
    (SELECT TOP 1 Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='FCF294'),
        'Aprovada', 
        'Aprovada', 
        'Aprovada', 
        'Lib. Operacao', 
        'Sim', 
        GETDATE());



    
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