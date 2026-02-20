SELECT tmpSQLServerTables.Table_Name, tmpSQLServerTables.Table_Type
FROM tmpSQLServerTables
WHERE (((tmpSQLServerTables.Table_Name) In (SELECT [Table_Name] FROM [tmpSQLServerTables] As Tmp GROUP BY [Table_Name] HAVING Count(*)>1 )))
ORDER BY tmpSQLServerTables.Table_Name;

