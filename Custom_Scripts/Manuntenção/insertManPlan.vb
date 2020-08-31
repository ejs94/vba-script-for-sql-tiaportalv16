Sub insertManPlan()
'Rotina irá escrever no Banco de Dados após confirmação das opções dispobnibilizadas.
Dim strFuncName, ManPlan_ID, SQL_Table, conn, rst
Dim pDATABASE, Reg_Edit_Table
Dim Responsavel, Descricao, horasPlanej, dataPlanej, Maquina, tipoManuntenc, Prioridade

pDATABASE = "hmiDB"
strFuncName = "insertManPlan"

On Error Resume Next

Responsavel = SmartTags("edit_respons")
Descricao = SmartTags("edit_descr")
horasPlanej = SmartTags("edit_h_plan")
dataPlanej = SmartTags("edit_dt_mant")

SmartTags("Ultimo_WWID") = "ManPlan"

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
	ShowSystemAlarm "Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Set conn = Nothing
	Exit Sub
End If


' Chaveamento para conversar com o text file do TIA PORTAL
Select Case SmartTags("edit_maqEqu")
    Case 0
        Maquina = "Outro"
    Case 1 
        Maquina = "Est. Entrad."
    Case 2
        Maquina = "Est. Saida."
    Case 3
        Maquina = "MCH250"
    Case 4
        Maquina = "MCH350"
    Case 5
        Maquina = "G704"
    Case 6
        Maquina = "G516"
    Case Else
        Maquina = "Outro"
End Select

Select Case SmartTags("edit_TipoManuten")
    Case 0
        tipoManuntenc = "Outro"
    Case 1
        tipoManuntenc = "Mecânico"
    Case 2
        tipoManuntenc = "Elétrico"
    Case 3
        tipoManuntenc = "Software"
    Case Else
        tipoManuntenc = "Outro"
End Select

Select Case SmartTags("edit_prior")
    Case 0
        Prioridade = "Baixa"
    Case 1
        Prioridade = "Alta"
    Case Else
        Prioridade = "Baixa"
End Select

'Caso a ID seja válida então poderá ocorrer a alteranção no Banco de Dados
If Responsavel <> "" AND Descricao <> "" Then
    SQL_Table = "USE hmiDB; " &_
                " UPDATE manPlanejada" &_
                " SET equip='" & Maquina & "'," &_
                " tipoManunt='" & tipoManuntenc & "'," &_
                " priorid='" & Prioridade & "'," &_
                " resposavel='" & Responsavel & "'," &_
                " descri='" & Descricao & "'," &_
                " hr_planej='" & horasPlanej & "'," &_
                " dia_manunt='" & dataPlanej & "'," &_
                " dt_Ultima_Alter=" & "GETDATE()" &_
                " WHERE manPlan_id=" & ManPlan_ID & "; "

    SQL_Table = "USE hmiDB; " &_
            " INSERT INTO manPlanejada" &_
            " (Modelo_id, ModeloString, NomeModelo)" &_
            " Values (" & Model_ID & ", " &_
            "'" & Modelo & "', " &_
            "'" & NomeModelo & "' );"
    
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




(manPlan_id,equip,tipoManunt,priorid,resposavel,descri,hr_planej,ativo,dia_manunt,dt_Ultima_Alter)