SELECT TC.Constraint_Name, CC.Column_Name, TC.table_name
FROM [information_schema].table_constraints AS TC INNER JOIN [information_schema].constraint_column_usage AS CC ON TC.Constraint_Name = CC.Constraint_Name
WHERE cc.Column_Name='Mandatory' And TC.table_name='SQLServerTesting'
ORDER BY TC.Constraint_Name;

