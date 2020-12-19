Sub showAllStatistic()


Dim strFuncName, conn, rst, pDATABASE, i, j
Dim SQL_MCH250, SQL_MCH350, SQL_G516, SQL_G704
Dim aprovada0, aprovada1, aprovada2
Dim refugo0, refugo1, refugo2
Dim medicao0, medicao1, medicao2
Dim libop, trab

strFuncName = "showAllStatistic"
pDATABASE = "hmiDB"


On Error Resume Next


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


SQL_MCH250 = "USE hmiDB; " &_
                " SELECT LTRIM(S.opBB155),COUNT(S.opBB155) " &_
                " FROM RegEntradaBlocos AS B " &_
                " JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id " &_
                " LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id " &_
                " GROUP BY S.opBB155; "

SQL_MCH350 = "USE hmiDB; " &_
                " SELECT LTRIM(S.opBB165),COUNT(S.opBB165) " &_
                " FROM RegEntradaBlocos AS B " &_
                " JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id " &_
                " LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id " &_
                " GROUP BY S.opBB165; "

SQL_G516 = "USE hmiDB; " &_
                " SELECT LTRIM(S.opBB175),COUNT(S.opBB175) " &_
                " FROM RegEntradaBlocos AS B " &_
                " JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id " &_
                " LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id " &_
                " GROUP BY S.opBB175; "

SQL_G704 = "USE hmiDB; " &_
                " SELECT LTRIM(S.opBB185),COUNT(S.opBB185) " &_
                " FROM RegEntradaBlocos AS B " &_
                " JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id " &_
                " LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id " &_
                " GROUP BY S.opBB185; "




'''''' ESTATISTICA DA MCH 350

aprovada0 = 0
aprovada1 = 0
aprovada2 = 0
refugo0 = 0
refugo1 = 0
refugo2 = 0
medicao0 = 0
medicao1 = 0
medicao2 = 0
libop = 0
trab = 0


Set rst = conn.Execute(SQL_MCH250)
If Not (rst.EOF And rst.BOF) Then
	rst.MoveFirst 'PRIMEIRO DADO RECEBIDO 

	'ZERA ITERADOR
	j=0
	
	'VERIFICA QUANTIDADE DE ELEMENTOS NA TABELA
	Do
		j=j+1
		rst.MoveNext
	Loop Until rst.EOF
    
    For i=0 To j
        rst.MoveFirst 'VOLTA AO PRIMEIRO DADO RECEBIDO 
            If rst.EOF Then
                SmartTags("DB_Estatistica_IHM_MCH250TotalRetentivo") = 0
                SmartTags("DB_Estatistica_IHM_MCH250_REFUGO_%_RETENTIVO") = 0
                Else             
                    Select Case rst.Fields(0).Value
                        Case "Aprovada"
                            aprovada0 = rst.Fields(1).Value
                        Case "Aprovada P1"
                            aprovada1 = rst.Fields(1).Value
                        Case "Aprovada P2"
                            aprovada2 = rst.Fields(1).Value
                        Case "Lib. Operacao"
                            libop = rst.Fields(1).Value
                        Case "Medicao"
                            medicao0 = rst.Fields(1).Value
                        Case "Medicao P1"
                            medicao1 = rst.Fields(1).Value
                        Case "Medicao P2"
                            medicao2 = rst.Fields(1).Value
                        Case "Refugo"
                            refugo0 = rst.Fields(1).Value
                        Case "Refugo P1"
                            refugo1 = rst.Fields(1).Value
                        Case "Refugo P2"
                            refugo2 = rst.Fields(1).Value
                        Case "Trabalha"
                            trab = rst.Fields(1).Value
                    End Select
            End If
        rst.MoveNext
    Next
End If
rst.close

SmartTags("DB_Estatistica_IHM_MCH250TotalRetentivo") = aprovada0 + aprovada1 + aprovada2 + refugo0 + refugo1 + refugo2 + medicao0 + medicao1 + medicao2 + libop + trab
SmartTags("DB_Estatistica_IHM_MCH250_REFUGO_%_RETENTIVO") = (refugo0 + refugo1 + refugo2)/( aprovada0 + aprovada1 + aprovada2 + refugo0 + refugo1 + refugo2 + medicao0 + medicao1 + medicao2 + libop + trab )


'''''''''''''' Estatistica da MCH350

aprovada0 = 0
aprovada1 = 0
aprovada2 = 0
refugo0 = 0
refugo1 = 0
refugo2 = 0
medicao0 = 0
medicao1 = 0
medicao2 = 0
libop = 0
trab = 0


Set rst = conn.Execute(SQL_MCH350)
If Not (rst.EOF And rst.BOF) Then
	rst.MoveFirst 'PRIMEIRO DADO RECEBIDO 

	'ZERA ITERADOR
	j=0
	
	'VERIFICA QUANTIDADE DE ELEMENTOS NA TABELA
	Do
		j=j+1
		rst.MoveNext
	Loop Until rst.EOF
	
    For i=0 To j
        rst.MoveFirst 'VOLTA AO PRIMEIRO DADO RECEBIDO 
            If rst.EOF Then
                SmartTags("DB_Estatistica_IHM_MCH350TotalRetentivo") = 0
                SmartTags("DB_Estatistica_IHM_MCH350_REFUGO_%_RETENTIVO") = 0
                Else             
                    Select Case rst.Fields(0).Value
                        Case "Aprovada"
                            aprovada0 = rst.Fields(1).Value
                        Case "Aprovada P1"
                            aprovada1 = rst.Fields(1).Value
                        Case "Aprovada P2"
                            aprovada2 = rst.Fields(1).Value
                        Case "Lib. Operacao"
                            libop = rst.Fields(1).Value
                        Case "Medicao"
                            medicao0 = rst.Fields(1).Value
                        Case "Medicao P1"
                            medicao1 = rst.Fields(1).Value
                        Case "Medicao P2"
                            medicao2 = rst.Fields(1).Value
                        Case "Refugo"
                            refugo0 = rst.Fields(1).Value
                        Case "Refugo P1"
                            refugo1 = rst.Fields(1).Value
                        Case "Refugo P2"
                            refugo2 = rst.Fields(1).Value
                        Case "Trabalha"
                            trab = rst.Fields(1).Value
                    End Select
            End If
        rst.MoveNext
    Next
End If
rst.close

SmartTags("DB_Estatistica_IHM_MCH350TotalRetentivo") = aprovada0 + aprovada1 + aprovada2 + refugo0 + refugo1 + refugo2 + medicao0 + medicao1 + medicao2 + libop + trab
SmartTags("DB_Estatistica_IHM_MCH350_REFUGO_%_RETENTIVO") = (refugo0 + refugo1 + refugo2)/( aprovada0 + aprovada1 + aprovada2 + refugo0 + refugo1 + refugo2 + medicao0 + medicao1 + medicao2 + libop + trab )


'''''''''''''''''''''''''''' Estatistica da Grob G704

aprovada0 = 0
aprovada1 = 0
aprovada2 = 0
refugo0 = 0
refugo1 = 0
refugo2 = 0
medicao0 = 0
medicao1 = 0
medicao2 = 0
libop = 0
trab = 0


Set rst = conn.Execute(SQL_G704)
If Not (rst.EOF And rst.BOF) Then
	rst.MoveFirst 'PRIMEIRO DADO RECEBIDO 

	'ZERA ITERADOR
	j=0
	
	'VERIFICA QUANTIDADE DE ELEMENTOS NA TABELA
	Do
		j=j+1
		rst.MoveNext
	Loop Until rst.EOF
	
    For i=0 To j
        rst.MoveFirst 'VOLTA AO PRIMEIRO DADO RECEBIDO 
            If rst.EOF Then
                SmartTags("DB_Estatistica_IHM_G704TotalRetentivo") = 0
                SmartTags("DB_Estatistica_IHM_G704_REFUGO_%_RETENTIVO") = 0
                Else             
                    Select Case rst.Fields(0).Value
                        Case "Aprovada"
                            aprovada0 = rst.Fields(1).Value
                        Case "Aprovada P1"
                            aprovada1 = rst.Fields(1).Value
                        Case "Aprovada P2"
                            aprovada2 = rst.Fields(1).Value
                        Case "Lib. Operacao"
                            libop = rst.Fields(1).Value
                        Case "Medicao"
                            medicao0 = rst.Fields(1).Value
                        Case "Medicao P1"
                            medicao1 = rst.Fields(1).Value
                        Case "Medicao P2"
                            medicao2 = rst.Fields(1).Value
                        Case "Refugo"
                            refugo0 = rst.Fields(1).Value
                        Case "Refugo P1"
                            refugo1 = rst.Fields(1).Value
                        Case "Refugo P2"
                            refugo2 = rst.Fields(1).Value
                        Case "Trabalha"
                            trab = rst.Fields(1).Value
                    End Select
            End If
        rst.MoveNext
    Next
End If
rst.close

SmartTags("DB_Estatistica_IHM_G704TotalRetentivo") = aprovada0 + aprovada1 + aprovada2 + refugo0 + refugo1 + refugo2 + medicao0 + medicao1 + medicao2 + libop + trab
SmartTags("DB_Estatistica_IHM_G704_REFUGO_%_RETENTIVO") = (refugo0 + refugo1 + refugo2)/( aprovada0 + aprovada1 + aprovada2 + refugo0 + refugo1 + refugo2 + medicao0 + medicao1 + medicao2 + libop + trab )



''''''''''''''''''''''''''' Estatistica da G516


aprovada0 = 0
aprovada1 = 0
aprovada2 = 0
refugo0 = 0
refugo1 = 0
refugo2 = 0
medicao0 = 0
medicao1 = 0
medicao2 = 0
libop = 0
trab = 0


Set rst = conn.Execute(SQL_G516)
If Not (rst.EOF And rst.BOF) Then
	rst.MoveFirst 'PRIMEIRO DADO RECEBIDO 

	'ZERA ITERADOR
	j=0
	
	'VERIFICA QUANTIDADE DE ELEMENTOS NA TABELA
	Do
		j=j+1
		rst.MoveNext
	Loop Until rst.EOF
	
    For i=0 To j
        rst.MoveFirst 'VOLTA AO PRIMEIRO DADO RECEBIDO 
            If rst.EOF Then
                SmartTags("DB_Estatistica_IHM_G516TotalRetentivo") = 0
                SmartTags("DB_Estatistica_IHM_G516_REFUGO_%_RETENTIVO") = 0
                Else             
                    Select Case rst.Fields(0).Value
                        Case "Aprovada"
                            aprovada0 = rst.Fields(1).Value
                        Case "Aprovada P1"
                            aprovada1 = rst.Fields(1).Value
                        Case "Aprovada P2"
                            aprovada2 = rst.Fields(1).Value
                        Case "Lib. Operacao"
                            libop = rst.Fields(1).Value
                        Case "Medicao"
                            medicao0 = rst.Fields(1).Value
                        Case "Medicao P1"
                            medicao1 = rst.Fields(1).Value
                        Case "Medicao P2"
                            medicao2 = rst.Fields(1).Value
                        Case "Refugo"
                            refugo0 = rst.Fields(1).Value
                        Case "Refugo P1"
                            refugo1 = rst.Fields(1).Value
                        Case "Refugo P2"
                            refugo2 = rst.Fields(1).Value
                        Case "Trabalha"
                            trab = rst.Fields(1).Value
                    End Select
            End If
        rst.MoveNext
    Next
End If
rst.close

SmartTags("DB_Estatistica_IHM_G516TotalRetentivo") = aprovada0 + aprovada1 + aprovada2 + refugo0 + refugo1 + refugo2 + medicao0 + medicao1 + medicao2 + libop + trab
SmartTags("DB_Estatistica_IHM_G516_REFUGO_%_RETENTIVO") = (refugo0 + refugo1 + refugo2)/( aprovada0 + aprovada1 + aprovada2 + refugo0 + refugo1 + refugo2 + medicao0 + medicao1 + medicao2 + libop + trab )



''''''''TRATA ERROS''''''''''''''''''''''''''''''''''
If Err.Number <> 0 Then
	ShowSystemAlarm strFuncName & ": Error #" & Err.Number & " " & Err.Description
	Err.Clear
	Exit Sub
End If

'Fecha todas conexões
rst.close
conn.close
Set rst = Nothing
Set conn = Nothing



End Sub