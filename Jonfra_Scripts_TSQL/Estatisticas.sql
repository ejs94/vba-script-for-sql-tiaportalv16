USE hmiDB;
 SELECT LTRIM(S.opBB155),COUNT(S.opBB155)
 FROM RegEntradaBlocos AS B 
 JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id 	
 LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id 
 ---WHERE S.dt_Saida BETWEEN '2020/12/18 00:00:00' AND '2020/12/18 23:59:00'
 GROUP BY S.opBB155;

SELECT LTRIM(S.opBB165),COUNT(S.opBB165)
 FROM RegEntradaBlocos AS B 
 JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id 	
 LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id 
 ---WHERE S.dt_Saida BETWEEN '2020/12/18 00:00:00' AND '2020/12/18 23:59:00'
 GROUP BY S.opBB165;


SELECT LTRIM(S.opBB175),COUNT(S.opBB175)
 FROM RegEntradaBlocos AS B 
 JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id 	
 LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id 
 ---WHERE S.dt_Saida BETWEEN '2020/12/18 00:00:00' AND '2020/12/18 23:59:00'
 GROUP BY S.opBB175;

SELECT LTRIM(S.opBB175),COUNT(S.opBB175)
 FROM RegEntradaBlocos AS B 
 JOIN RegSaidaBlocos AS S ON B.Bloco_id = S.Bloco_id 	
 LEFT JOIN ModelosBlocos AS M ON B.Modelo_id = M.Modelo_id 
 ---WHERE S.dt_Saida BETWEEN '2020/12/18 00:00:00' AND '2020/12/18 23:59:00'
 GROUP BY S.opBB175;