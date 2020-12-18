Function returnSQLString(ByRef StartTime, ByRef EndTime, ByRef OffSet)

Dim SQL_Seriais

''''''''''''''''''''''''''' STRINGS SQL ''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Query para buscar as Strings de todos os Turnos
SQL_Seriais = " USE hmiDB; " &_
                " SELECT S.dt_Saida,B.PNSerialString " &_
                " FROM RegEntradaBlocos AS B " &_
                " JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id " &_
                " LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id " &_
                " WHERE S.dt_Saida BETWEEN CAST(CONVERT(date,GETDATE()+" & OffSet & ") AS varchar)+' "& StartTime &"' AND CAST(CONVERT(date,GETDATE()) AS varchar)+' "& EndTime &"'" &_
                " ORDER BY S.dt_Saida; "

returnSQLString = SQL_Seriais


End Function