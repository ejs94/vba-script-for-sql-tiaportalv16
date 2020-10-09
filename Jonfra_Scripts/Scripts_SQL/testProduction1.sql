--- Registros de entrada
INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('FFF224', 11, '19/03/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('FFE464', 11, '18/04/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('FCF294', 12, '19/03/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('AFE434', 13, '19/04/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('F17764', 11, '18/04/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('3CE694', 12, '19/03/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('7FE444', 33, '19/04/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('F13464', 30, '16/04/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('3CF294', 13, '13/03/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('7FE224', 32, '15/04/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('F23864', 33, '16/04/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('33F264', 33, '13/03/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('74E4D4', 12, '15/04/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('7FE334', 32, '15/04/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('FFF999', 12, '16/04/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('FFF998', 12, '13/03/20', GETDATE());

INSERT INTO RegEntradaBlocos
    (PNSerialString,Modelo_id,DataString,dt_Entrada)
Values('FFF987', 13, '15/04/20', GETDATE());


--- Registro de Saida
INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='FFF224'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());


INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='FFE464'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());

INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='FCF294'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());

INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='AFE434'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());

INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='F17764'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());


INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='3CE694'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());

INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='7FE444'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());

INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='F13464'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());

INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='3CF294'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());


INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='7FE224'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());

INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='F23864'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());

INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='33F264'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());

INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='74E4D4'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());


INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='7FE334'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());

INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='FFF999'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());

INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='FFF998'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());

INSERT INTO RegSaidaBlocos
    (Bloco_id,opBB155,opBB165,opBB175,opBB185,inspecao,dt_Saida)
Values
    ((SELECT Bloco_id
        FROM RegEntradaBlocos
        WHERE PNSerialString='FFF987'),
        'Aprovada', 'Aprovada', 'Aprovada', 'Aprovada', 'Nao', GETDATE());
