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

--- FAZER SCIPT DE REGISTRO DE BLOCOS
--- Registro de Entrada de Blocos com SubQuery
USE hmiDB;
INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values
    ( '10123',
        (SELECT Modelo_id
        FROM ModelosBlocos
        WHERE ModeloString='3932'), --- Se o modelo não estiver na tabela o valor será NULL
        '18/09/20',
        GETDATE());

--- Registro de Blocos na Saída
--- Registro quando o bloco sai da celula
INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='FFF124'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());

--- Registro quando o bloco sai da celula
INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='FFE364'),
        'Aprovada', 'Aprovada', 'Refugada', 'Lib. OP', 'Nao', GETDATE());