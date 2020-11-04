--- Banco de dados ao qual iremos criar a tabela
USE hmiDB;

---TODO Fazer um registro de modelo atraves do banco de dados, que se foda o recept do tia portal
--- Tabela de registro de modelos de blocos
CREATE TABLE ModelosBlocos
(
	Modelo_id INT PRIMARY KEY,
	ModeloString VARCHAR(7) NOT NULL UNIQUE,
	NomeModelo VARCHAR(40) NOT NULL UNIQUE,
	DiametroCamisa INT NOT NULL,
);

--- Tabela de registro de blocos na entrada
CREATE TABLE RegEntradaBlocos
(
	Bloco_id INT IDENTITY PRIMARY KEY,
	PNSerialString VARCHAR(6) NOT NULL,
	Modelo_id INT FOREIGN KEY REFERENCES ModelosBlocos(Modelo_id)
	ON DELETE SET NULL
    ON UPDATE CASCADE,
	DataString VARCHAR(8),
	dt_Entrada DATETIME,
);

--- Tabela preenchida com a saida dos blocos
CREATE TABLE RegSaidaBlocos
(
	Producao_id INT IDENTITY PRIMARY KEY,
	Bloco_id INT FOREIGN KEY REFERENCES RegEntradaBlocos(Bloco_id)
	ON DELETE SET NULL
    ON UPDATE CASCADE,
	opBB155 VARCHAR
	(9) NOT NULL,
	opBB165 VARCHAR
	(9) NOT NULL,
	opBB175 VARCHAR
	(9) NOT NULL,
	opBB185 VARCHAR
	(9) NOT NULL,
	inspecao VARCHAR(5) DEFAULT 'Nao',
	dt_Saida DATETIME
);

--- Tabela com as alterações realizada na tabela de produção
CREATE TABLE alterProducTable
(
	Alter_id INT IDENTITY PRIMARY KEY,
	comando VARCHAR(400),
	wwid VARCHAR
	(20) NOT NULL,
	dt_Alteracao DATETIME
);

--- Tabela das manuntenções planejadas
CREATE TABLE manPlanejada
(
	manPlan_id INT IDENTITY PRIMARY KEY,
	equip VARCHAR(20),
	tipoManunt VARCHAR
	(20),
	priorid VARCHAR(20),
	resposavel VARCHAR(20),
	descri VARCHAR(150),
	hr_planej TIME,
	ativo BIT,
	dia_manunt DATE,
	dt_Ultima_Alter DATETIME
);