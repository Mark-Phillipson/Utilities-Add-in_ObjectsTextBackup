SELECT tmpSQLServerTables.Table_Name
FROM tmpSQLServerTables
WHERE (((tmpSQLServerTables.Table_Type)="TABLE")) OR (((tmpSQLServerTables.Table_Type)="VIEW"))
ORDER BY tmpSQLServerTables.Table_Name;

