--- Banco de dados ao qual iremos criar a tabela
USE hmiDB;

--- Criando uma tabela baseada num Type DB do TIA PORTAL
CREATE TABLE RastreabilidadeBloco
(
	Dado_id INT IDENTITY(1,1) PRIMARY KEY,
	SerialString VARCHAR(6) NOT NULL,
	ModeloString VARCHAR(7) NOT NULL,
	DataString VARCHAR(8),
	datatimeEntrada DATETIME NOT NULL,
	datatimeSaida DATETIME NOT NULL,
	ultimaOP INT,
	ultimaLeitura DATETIME,
	opBB155 VARCHAR
	(14) NOT NULL,
	opBB165 VARCHAR
	(14) NOT NULL,
	opBB175 VARCHAR
	(14) NOT NULL,
	opBB185 VARCHAR
	(14) NOT NULL,
	tipoCarga int NOT NULL,
);