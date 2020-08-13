--- Banco de dados ao qual iremos criar a tabela
USE hmiDB;

--- Tabela de registro de blocos na entrada
CREATE TABLE RegEntradaBlocos
(
	Bloco_id INT IDENTITY PRIMARY KEY,
	PNSerialString VARCHAR(6) NOT NULL,
	ModeloString VARCHAR(7) NOT NULL,
	DataString VARCHAR(8),
	dt_Entrada DATETIME,
);

--- Tabela preenchida com a saida dos blocos
CREATE TABLE RegSaidaBlocos
(
	Producao_id INT IDENTITY PRIMARY KEY,
	Bloco_id INT FOREIGN KEY REFERENCES RegEntradaBlocos(Bloco_id)
	ON DELETE CASCADE
    ON UPDATE CASCADE,
	opBB155 VARCHAR
	(14) NOT NULL,
	opBB165 VARCHAR
	(14) NOT NULL,
	opBB175 VARCHAR
	(14) NOT NULL,
	opBB185 VARCHAR
	(14) NOT NULL,
	inspecao VARCHAR(5) DEFAULT 'Nao',
	dt_Saida DATETIME
);