SELECT        o.Name AS ObjectName,
            o.type AS ObjectType,
            s.name AS SchemaOwner,
            ep.name AS PropertyName,
            substring(cast(ep.value AS nvarchar(255)) ,0,255) AS PropertyValue,
            c.name AS ColumnName,
            c.colid AS Ordinal
FROM        sys.objects o INNER JOIN sys.extended_properties ep
            ON o.object_id = ep.major_id
            INNER JOIN sys.schemas s
            ON o.schema_id = s.schema_id
            LEFT JOIN syscolumns c
            ON ep.minor_id = c.colid
            AND ep.major_id = c.id
WHERE        o.type ='U' and ep.name='MS_Caption' and ep.Class_desc='OBJECT_OR_COLUMN'
ORDER BY    SchemaOwner,ObjectName, ObjectType, PropertyName


