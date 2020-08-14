--- Query to see the result:
--- Select para cruzar as duas tabelas
SELECT S.Producao_id, B.PNSerialString, B.ModeloString, S.opBB155, S.opBB165, S.opBB175, S.opBB185, S.inspecao, B.dt_Entrada, S.dt_Saida
FROM RegEntradaBlocos AS B
    RIGHT JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id;


--- 