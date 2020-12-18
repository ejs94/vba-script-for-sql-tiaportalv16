Sub returnSQLCount(ByRef StartTime, ByRef EndTime, ByRef OffSet, ByRef SQL_Conforme_NConfome)

''''''''''''''''''''''''''' STRINGS SQL ''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Query para contar o número de peças que foram conforme ou não conforme durante a produção
SQL_Conforme_NConfome = " USE hmiDB;" &_
                                " SELECT" &_
                                " COUNT(CASE " &_
                                " WHEN opBB155 = 'Aprovada P1' or opBB155 = 'Aprovada P2'" &_
                                " OR opBB165 = 'Aprovada P1' or opBB165 = 'Aprovada P2'" &_
                                " OR opBB165 = 'Aprovada P1' or opBB165 = 'Aprovada P2'" &_
                                " OR opBB175 = 'Aprovada'" &_
                                " AND LTRIM(opBB185) = 'Aprovada'" &_
                                " AND NOT (opBB155 = 'Refugo P1' or opBB155 = 'Refugo P2' or opBB165 = 'Refugo P1' or opBB165 = 'Refugo P2' or opBB175 = 'Refugo' or LTRIM(opBB185) = 'Refugo')" &_
			                    " or ( opBB155 = 'Trabalha' AND opBB165 = 'Trabalha' AND opBB175 = 'Trabalha' AND LTRIM(opBB185) = 'Trabalha')"
                                " THEN 1 END) As Conforme," &_
                                " COUNT(CASE WHEN opBB155 = 'Refugo P1' or opBB155 = 'Refugo P2'" &_
                                " or opBB165 = 'Refugo P1' or opBB165 = 'Refugo P2'" &_
                                " or opBB175 = 'Refugo'" &_
                                " or LTRIM(opBB185) = 'Refugo' THEN 1 END) As Nao_Conforme" &_
                                " FROM RegEntradaBlocos AS B" &_
                                " JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id" &_
                                " WHERE S.dt_Saida BETWEEN CAST(CONVERT(date,GETDATE()+" & OffSet & ") AS varchar)+' "& StartTime &"' AND CAST(CONVERT(date,GETDATE()) AS varchar)+' "& EndTime &"';"

End Sub