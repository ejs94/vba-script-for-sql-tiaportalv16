Function returnSQLString(ByRef StartTime, ByRef EndTime, ByRef StartOffSet, ByRef EndOffSet)

Dim SQL_Seriais

''''''''''''''''''''''''''' STRINGS SQL ''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Query para buscar as Strings de todos os Turnos
SQL_Seriais = " USE hmiDB; " &_
                " SELECT S.dt_Saida, B.PNSerialString, S.opBB155, S.opBB165, S.opBB175, LTRIM(S.opBB185) " &_
                " FROM RegEntradaBlocos AS B " &_
                " JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id " &_
                " LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id " &_
                " WHERE S.dt_Saida BETWEEN CAST(CONVERT(date,GETDATE()+" & StartOffSet & ") AS varchar)+' "& StartTime &"' AND CAST(CONVERT(date,GETDATE()+" & EndOffSet & ") AS varchar)+' "& EndTime &"'" &_
                " ORDER BY S.dt_Saida; "

returnSQLString = SQL_Seriais


End Function