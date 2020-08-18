--- Query to see the result:
--- Select para cruzar as duas tabelas
SELECT S.Producao_id, B.PNSerialString, M.ModeloString, M.NomeModelo, S.opBB155, S.opBB165, S.opBB175, S.opBB185, S.inspecao, B.dt_Entrada, S.dt_Saida
FROM RegEntradaBlocos AS B
    RIGHT JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id
    INNER JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id;