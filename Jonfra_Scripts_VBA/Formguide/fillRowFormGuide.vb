Sub fillRowFormGuide(ByRef conn, ByRef rst, ByRef SQL_StartTime, ByRef SQL_EndTime, ByRef SQL_DayOffset, ByRef IHM_Turno, ByRef IHM_Linha)

Dim SQL_Seriais, tempCountConforme, tempCountNaoConforme, i, aux1, aux2
Dim strFuncName

strFuncName = "fillRowFormGuide"

tempCountConforme = 0
tempCountNaoConforme = 0
aux1 = 0
aux2 = 0

SQL_Seriais = returnSQLString(SQL_StartTime, SQL_EndTime, SQL_DayOffset )

Set rst = conn.Execute(SQL_Seriais)
If Not (rst.EOF And rst.BOF) Then
    rst.MoveFirst
    For i = 1 To 15
        If rst.EOF Then
            SmartTags("DB_Contador_Producao_Dados_Turno_"& IHM_Turno &"_SN_" & IHM_Linha & "{" & i & "}") = ""
	        Else
                SmartTags("DB_Contador_Producao_Dados_Turno_" & IHM_Turno & "_SN_" & IHM_Linha & "{" & i & "}") = rst.Fields(1).Value                

	            If (rst.Fields(2).Value = "Aprovada P1" Or rst.Fields(2).Value = "Aprovada P2" Or rst.Fields(3).Value = "Aprovada P1" Or rst.Fields(3).Value = "Aprovada P2" Or rst.Fields(3).Value = "Aprovada P1" Or rst.Fields(3).Value = "Aprovada P2" Or rst.Fields(4).Value = "Aprovada" Or rst.Fields(5).Value = "Aprovada") And Not (rst.Fields(2).Value = "Refugo P1" Or rst.Fields(2).Value = "Refugo P2" Or rst.Fields(3).Value = "Refugo P1" Or rst.Fields(3).Value = "Refugo P2" Or rst.Fields(4).Value = "Refugo" Or rst.Fields(5).Value = "Refugo") Then
	            	tempCountConforme = 1
	            	Else
	            		tempCountConforme = 0
	            End If                

 	            If rst.Fields(2).Value = "Refugo P1" Or rst.Fields(2).Value = "Refugo P2"  Or rst.Fields(3).Value = "Refugo P1" Or rst.Fields(3).Value = "Refugo P2" Or rst.Fields(4).Value = "Refugo"  Or rst.Fields(5).Value = "Refugo"  Or (rst.Fields(2).Value = "Trabalha" And rst.Fields(3).Value = "Trabalha" And rst.Fields(4).Value = "Trabalha" And rst.Fields(5).Value = "Trabalha") Then
	            	tempCountNaoConforme = 1
	            	Else
	            		tempCountNaoConforme = 0
	            End If

                aux1 = aux1 + tempCountConforme
                aux2 = aux2 + tempCountNaoConforme
                rst.MoveNext
        End If
    Next
End If
rst.close


SmartTags("DB_Contador_Producao_Dados_Turno_" & IHM_Turno & "_Contador_OK{"& IHM_Linha &"}") = aux1
SmartTags("DB_Contador_Producao_Dados_Turno_" & IHM_Turno & "_Contador_LIB OP{" & IHM_Linha & "}") = aux2

showLog strFuncName & " Fora do For: " & tempCountConforme
showLog strFuncName & " Turno:" & IHM_Turno & " Linha: " & IHM_Linha & " SQL_Query: " & SQL_Seriais
showLog strFuncName & " Ok:" & aux1 & " NOk: " & aux2


End Sub