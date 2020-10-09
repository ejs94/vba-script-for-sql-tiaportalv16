--- Inserindo os modelos dos blocos
INSERT INTO ModelosBlocos
    (Modelo_id,ModeloString,NomeModelo)
VALUES(10, '3938364', 'B4 Mecânico');


--- Registro de Blocos na Entrada
--- Registro quando o bloco entra na celula
INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('FFF124', 10, '16/03/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('FFE364', 30, '17/04/20', GETDATE());

--- Registro de Entrada de Blocos com SubQuery
INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values
    ( 'FFFEEE',
        (SELECT IFNULL(Modelo_id, 'Nao Reg')
        FROM ModeloBlocos
        WHERE ModeloString='364'), --- Se o modelo não estiver na tabela o valor será NULL
        '18/05/20',
        GETDATE());

--- Registro de Blocos na Saída
--- Registro quando o bloco sai da celula
INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='FFF124'),
        'Okey', 'Okey', 'Okey', 'Okey', 'Nao', GETDATE());

--- Registro quando o bloco sai da celula
INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='FFE364'),
        'Okey', 'Okey', 'N Okey', 'Okey', 'Sim', GETDATE());