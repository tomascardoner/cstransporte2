GRANT CONNECT TO cstransporte
GO

GRANT ALTER ON CuentaCorriente TO cstransporte
GO

-----------------------------
--GRANT PERMISSIONS ON TABLES
-----------------------------
DECLARE TablesCursor CURSOR LOCAL FORWARD_ONLY STATIC FOR SELECT name FROM sysobjects WHERE xtype = 'U' ORDER BY name
DECLARE @TableName sysname

OPEN TablesCursor
FETCH NEXT FROM TablesCursor INTO @TableName

WHILE @@FETCH_STATUS = 0
BEGIN
	EXEC ('GRANT SELECT, INSERT, UPDATE, DELETE ON ' + @TableName + ' TO cstransporte')
	FETCH NEXT FROM TablesCursor INTO @TableName
END

CLOSE TablesCursor

DEALLOCATE TablesCursor


----------------------------------------
--GRANT PERMISSIONS ON STORED PROCEDURES
----------------------------------------
DECLARE StoredProceduresCursor CURSOR LOCAL FORWARD_ONLY STATIC FOR SELECT name FROM sysobjects WHERE xtype = 'P' AND category = 0 ORDER BY name
DECLARE @StoredProcedureName sysname

OPEN StoredProceduresCursor
FETCH NEXT FROM StoredProceduresCursor INTO @StoredProcedureName

WHILE @@FETCH_STATUS = 0
BEGIN
	IF @StoredProcedureName NOT IN ('sp_alterdiagram', 'sp_creatediagram', 'sp_dropdiagram', 'sp_helpdiagramdefinition', 'sp_helpdiagrams', 'sp_renamediagram', 'sp_upgraddiagrams')
		EXEC ('GRANT EXECUTE ON ' + @StoredProcedureName + ' TO cstransporte')
	FETCH NEXT FROM StoredProceduresCursor INTO @StoredProcedureName
END

CLOSE StoredProceduresCursor

DEALLOCATE StoredProceduresCursor

----------------------------------------
--GRANT PERMISSIONS ON FUNCTIONS
----------------------------------------
DECLARE FunctionsCursor CURSOR LOCAL FORWARD_ONLY STATIC FOR SELECT name FROM sysobjects WHERE xtype = 'FN' ORDER BY name
DECLARE @FunctionName sysname

OPEN FunctionsCursor
FETCH NEXT FROM FunctionsCursor INTO @StoredProcedureName

WHILE @@FETCH_STATUS = 0
BEGIN
	IF @FunctionName NOT IN ('fn_diagramobjects')
		EXEC ('GRANT EXECUTE ON ' + @FunctionName + ' TO cstransporte')
	FETCH NEXT FROM FunctionsCursor INTO @FunctionName
END

CLOSE FunctionsCursor

DEALLOCATE FunctionsCursor