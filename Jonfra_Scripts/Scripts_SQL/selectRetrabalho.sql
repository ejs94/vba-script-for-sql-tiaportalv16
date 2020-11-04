--- Para o retrabalho é necessário realizar um Query no Banco de Dados, que responda:
--- Quais operações a peça realizou?
--- Qual Data e Hora de Entrada e Saída?
--- Buscar pela Serial da Peça na Esteira de entrada!
--- Sempre pegar o último valor de produção!


SELECT TOP 1
    S.Producao_id,
    B.PNSerialString AS Serial,
    M.ModeloString AS Modelo,
    B.DataString AS 'Data Serial',
    S.opBB155 AS MCH250,
    S.opBB165 AS MCH350,
    S.opBB175 AS G705,
    S.opBB185 AS G516,
    B.dt_Entrada AS Entrada,
    S.dt_Saida AS Saida
FROM RegEntradaBlocos AS B
    JOIN RegSaidaBlocos AS S
    ON B.Bloco_id = S.Bloco_id
    LEFT JOIN ModelosBlocos AS M
    ON B.Modelo_id = M.Modelo_id
--- Aqui a Variavel para buscar pela Serial
WHERE B.PNSerialString = '1000J0'

ORDER BY S.Producao_id DESC;

--- Esse Script será a base para o VBA de Retrabalho!