--- Registro quando o bloco entra na celula
INSERT INTO RegEntradaBlocos
    (PNSerialString,ModeloString,DataString,dt_Entrada)
Values('333124', '2134421', '16/03/20', GETDATE());

--- Registro quando o bloco sai da celula
INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    (9, 'Okey', 'Okey', 'Okey', 'Okey', 'Nao', GETDATE());

--- Outro exemplo, já que melhoro ele.
INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    (10, 'Nao Okey', 'Okey', 'Okey', 'Okey', 'Sim', GETDATE());