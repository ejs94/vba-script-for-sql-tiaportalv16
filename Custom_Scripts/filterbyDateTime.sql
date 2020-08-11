SELECT [Dados_id]
      , [Nome]
	  , [Nascimento]
FROM [hmiDB].[dbo].[Dados]
WHERE Nascimento >= '2020-08-7 17:51' AND Nascimento <= '2020-08-10';