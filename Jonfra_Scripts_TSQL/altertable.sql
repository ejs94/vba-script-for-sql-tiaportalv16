USE hmiDB;

ALTER TABLE ModelosBlocos ADD TamanhoBloco INT;

UPDATE ModelosBlocos
SET TamanhoBloco = 4
WHERE NomeModelo LIKE 'B4%'; 

UPDATE ModelosBlocos
SET TamanhoBloco = 6
WHERE NomeModelo LIKE 'B6%'; 

ALTER TABLE
ModelosBlocos 
ALTER COLUMN 
TamanhoBloco INT NOT NULL;