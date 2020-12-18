--- Retira a data atual e converte em string
SELECT CAST(CONVERT(date,GETDATE()) AS varchar);


--- Essa Query irá retornar todos valores da produção dentro de uma margem de horário
USE hmiDB;
SELECT B.PNSerialString, S.dt_Saida
FROM RegEntradaBlocos AS B
    JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id
    LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id
--- Pega todos os valores dentro de um Intervalo de tempo
WHERE S.dt_Saida BETWEEN CAST(CONVERT(date,GETDATE()) AS varchar)+' 22:00:00' AND CAST(CONVERT(date,GETDATE()) AS varchar)+' 23:00:00' 
ORDER BY S.dt_Saida DESC, S.Producao_id DESC;


--- Essa Query irá retornar TODOS os valores de Peça Conforme e Não Conforme
SELECT 
    COUNT(CASE WHEN opBB155 = 'Aprovada P1' or opBB155 = 'Aprovada P2' THEN 1 END) + 
        COUNT(CASE WHEN opBB165 = 'Aprovada P1' or opBB165 = 'Aprovada P2' THEN 1 END) +
            COUNT(CASE WHEN opBB175 = 'Aprovada' THEN 1 END) + 
                COUNT(CASE WHEN opBB185 = ' Aprovada' THEN 1 END) As Conforme,

    COUNT(CASE WHEN opBB155 != 'Aprovada P1' AND opBB155 != 'Aprovada P2' THEN 1 END) + 
        COUNT(CASE WHEN opBB165 != 'Aprovada P1' AND opBB165 != 'Aprovada P2' THEN 1 END) +
	        COUNT(CASE WHEN opBB175 != 'Aprovada' THEN 1 END) + 
                COUNT(CASE WHEN opBB185 != ' Aprovada' THEN 1 END) As Nao_Conforme


FROM RegEntradaBlocos AS B
    JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id
WHERE S.dt_Saida BETWEEN CAST(CONVERT(date,GETDATE()-1) AS varchar)+' 22:00:00' AND CAST(CONVERT(date,GETDATE()-1) AS varchar)+' 23:00:00'



--- Essa Query irá retornar os valores de Peça Conforme e Não Conforme de opBB155 e opBB165
SELECT 
    COUNT(CASE WHEN opBB155 = 'Aprovada P1' or opBB155 = 'Aprovada P2' THEN 1 END) + 
        COUNT(CASE WHEN opBB165 = 'Aprovada P1' or opBB165 = 'Aprovada P2' THEN 1 END) AS Conforme,
    COUNT(CASE WHEN opBB155 != 'Aprovada P1' AND opBB155 != 'Aprovada P2' THEN 1 END) + 
        COUNT(CASE WHEN opBB165 != 'Aprovada P1' AND opBB165 != 'Aprovada P2' THEN 1 END) AS Nao_Conforme
FROM RegEntradaBlocos AS B
    JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id
WHERE S.dt_Saida BETWEEN CAST(CONVERT(date,GETDATE()-1) AS varchar)+' 22:00:00' AND CAST(CONVERT(date,GETDATE()-1) AS varchar)+' 23:00:00'

--- Essa Query irá retornar os valores de Peça Conformes e Não Conformes de opBB175 e opBB185
SELECT 
	COUNT(CASE WHEN opBB175 = 'Aprovada' THEN 1 END) + 
        COUNT(CASE WHEN opBB185 = ' Aprovada' THEN 1 END) As Conforme,
	COUNT(CASE WHEN opBB175 != 'Aprovada' THEN 1 END) + 
        COUNT(CASE WHEN opBB185 != ' Aprovada' THEN 1 END) As Nao_Conforme
FROM RegEntradaBlocos AS B
    JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id
WHERE S.dt_Saida BETWEEN CAST(CONVERT(date,GETDATE()-1) AS varchar)+' 22:00:00' AND CAST(CONVERT(date,GETDATE()) AS varchar)+' 23:00:00'


--- Essa String que irá ser utilizada no FormGuide
USE hmiDB;

SELECT S.dt_Saida,B.PNSerialString
FROM RegEntradaBlocos AS B
    JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id
    LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id
--- Pega todos os valores dentro de um Intervalo de tempo
WHERE S.dt_Saida BETWEEN CAST(CONVERT(date,GETDATE()) AS varchar)+' 22:00:00' AND CAST(CONVERT(date,GETDATE()) AS varchar)+' 23:00:00' 
ORDER BY S.dt_Saida

USE hmiDB;
SELECT 
	COUNT(CASE 
		WHEN opBB155 = 'Aprovada P1' or opBB155 = 'Aprovada P2'
					OR opBB165 = 'Aprovada P1' or opBB165 = 'Aprovada P2'
					OR opBB165 = 'Aprovada P1' or opBB165 = 'Aprovada P2'
					OR opBB175 = 'Aprovada'
					AND LTRIM(opBB185) = 'Aprovada'
					AND NOT 
						(opBB155 = 'Refugo P1' 
						or opBB155 = 'Refugo P2' 
						or opBB165 = 'Refugo P1' 
						or opBB165 = 'Refugo P2'
						or opBB175 = 'Refugo'
						or LTRIM(opBB185) = 'Refugo') THEN 1 END) As Conforme,
	COUNT(CASE 
		WHEN opBB155 = 'Refugo P1' or opBB155 = 'Refugo P2' 
			or opBB165 = 'Refugo P1' or opBB165 = 'Refugo P2'
			or opBB175 = 'Refugo' 
			or LTRIM(opBB185) = 'Refugo' THEN 1 END) As Nao_Conforme

FROM RegEntradaBlocos AS B
    JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id
WHERE S.dt_Saida BETWEEN CAST(CONVERT(date,GETDATE()) AS varchar)+' 22:00:00' AND CAST(CONVERT(date,GETDATE()) AS varchar)+' 23:00:00'


USE hmiDB;
SELECT COUNT(S.Producao_id)
FROM RegEntradaBlocos AS B
    JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id
WHERE S.dt_Saida BETWEEN CAST(CONVERT(date,GETDATE()) AS varchar)+' 01:30:00' AND CAST(CONVERT(date,GETDATE()) AS varchar)+' 7:00:00'