SELECT tmpSQLServerTables.ID, tmpSQLServerTables.Table_Name, tmpSQLServerTables.Table_Type
FROM tmpSQLServerTables
WHERE ((([tmpSQLServerTables].[Table_Name] & " " & [tmpSQLServerTables].[Table_Type]) Like '*'))
ORDER BY tmpSQLServerTables.Table_Name;

