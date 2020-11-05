USE hmiDB;

--- Contar toda produção durante um periodo de tempo
SELECT COUNT(1)
FROM RegSaidaBlocos
---WHERE BETWEEN DataInicial AND DataFinal
;

--- Contar toda a produção de peça aprovadas durante um periodo de tempo
SELECT COUNT(1)
FROM RegSaidaBlocos;
--- WHERE opBB155 = $Valor OR opBB165 = $Valor OR opBB175 = $Valor OR opBB185 = $Valor;