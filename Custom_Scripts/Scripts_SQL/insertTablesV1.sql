--- Registro quando o bloco entra na celula
INSERT INTO RegEntradaBlocos
    (PNSerialString,ModeloString,DataString,dt_Entrada)
Values('333124', '2134421', '16/03/20', GETDATE());

--- Registro quando o bloco sai da celula
INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    (1, 'Okey', 'Okey', 'Okey', 'Okey', 'Nao', GETDATE());


--- Query to see the result:
--- Select para cruzar as duas tabelas
SELECT S.Producao_id, B.PNSerialString, B.ModeloString, S.opBB155, S.opBB165, S.opBB175, S.opBB185, S.inspecao, B.dt_Entrada, S.dt_Saida
FROM RegEntradaBlocos AS B
    RIGHT JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id;