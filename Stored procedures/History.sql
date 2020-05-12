USE CSTransporte_History
GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTE_LIST_SALDOANTERIOR
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_CuentaCorriente_List_SaldoAnterior'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_CuentaCorriente_List_SaldoAnterior
GO

CREATE PROCEDURE dbo.sp_CuentaCorriente_List_SaldoAnterior
	@IDPersona int,
	@IDCuentaCorrienteGrupo int,
	@IDCuentaCorrienteCaja int,
	@FilterTipo int,
	@FechaDesde smalldatetime,
	@SaldoAnterior money OUTPUT AS

	SET @SaldoAnterior = (SELECT SUM(CuentaCorriente.Importe)
		FROM CuentaCorriente
		WHERE (@IDPersona IS NULL OR CuentaCorriente.IDPersona = @IDPersona)
			AND (@IDCuentaCorrienteGrupo IS NULL OR CuentaCorriente.IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo)
			AND (@IDCuentaCorrienteCaja IS NULL OR CuentaCorriente.IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja)
			AND (@FilterTipo = 0 OR (@FilterTipo = 1 AND CuentaCorriente.Importe >= 0) OR (@FilterTipo = 2 AND CuentaCorriente.Importe < 0))
			AND CuentaCorriente.FechaHora < @FechaDesde)

GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTE_LIST_MOVIMIENTOS
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_CuentaCorriente_List_Movimientos'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_CuentaCorriente_List_Movimientos
GO

CREATE PROCEDURE dbo.sp_CuentaCorriente_List_Movimientos
	@IDPersona int,
	@IDCuentaCorrienteGrupo int,
	@IDCuentaCorrienteCaja int,
	@FilterTipo int,
	@FechaFiltro int,
	@FechaDesde smalldatetime,
	@FechaHasta smalldatetime AS

	--TODAS LAS FECHAS
	IF @FechaFiltro = 0
		INSERT INTO #CuentaCorriente
			(IDMovimiento, FechaHora, IDCuentaCorrienteGrupo, Grupo, IDCuentaCorrienteCaja, Caja, Persona, Descripcion, Realizado, Pasajero, Importe, ImporteAcumulado)
			SELECT TOP 1000 CuentaCorriente.IDMovimiento, CuentaCorriente.FechaHora, CuentaCorriente.IDCuentaCorrienteGrupo, CuentaCorrienteGrupo.Nombre AS Grupo, CuentaCorriente.IDCuentaCorrienteCaja, CuentaCorrienteCaja.Nombre AS Caja, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, CuentaCorriente.Descripcion, ViajeDetalle.Realizado, Pasajero.Apellido + (CASE ISNULL(Pasajero.Nombre, '') WHEN '' THEN '' ELSE ', ' + Pasajero.Nombre END) AS Pasajero, CuentaCorriente.Importe, NULL
				FROM ((((CuentaCorriente INNER JOIN CSTransporte..CuentaCorrienteGrupo ON CuentaCorriente.IDCuentaCorrienteGrupo = CSTransporte..CuentaCorrienteGrupo.IDCuentaCorrienteGrupo) INNER JOIN CSTransporte..CuentaCorrienteCaja ON CuentaCorriente.IDCuentaCorrienteCaja = CSTransporte..CuentaCorrienteCaja.IDCuentaCorrienteCaja) LEFT JOIN CSTransporte..Persona ON CuentaCorriente.IDPersona = CSTransporte..Persona.IDPersona) LEFT JOIN CSTransporte..Persona AS Pasajero ON CuentaCorriente.IDPersonaOrigen = Pasajero.IDPersona) LEFT JOIN ViajeDetalle ON CuentaCorriente.Viaje_FechaHora = ViajeDetalle.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalle.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalle.Indice	
				WHERE (@IDPersona IS NULL OR CuentaCorriente.IDPersona = @IDPersona)
					AND (@IDCuentaCorrienteGrupo IS NULL OR CuentaCorriente.IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo)
					AND (@IDCuentaCorrienteCaja IS NULL OR CuentaCorriente.IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja)
					AND ((@FilterTipo = 0) OR (@FilterTipo = 1 AND CuentaCorriente.Importe >= 0) OR (@FilterTipo = 2 AND CuentaCorriente.Importe < 0))
	
	--FECHA IGUAL A
	IF @FechaFiltro = 1
		BEGIN
		SET @FechaHasta = dateadd(minute, -1, dateadd(day, 1, @FechaDesde))
		INSERT INTO #CuentaCorriente
			(IDMovimiento, FechaHora, IDCuentaCorrienteGrupo, Grupo, IDCuentaCorrienteCaja, Caja, Persona, Descripcion, Realizado, Pasajero, Importe, ImporteAcumulado)
			SELECT TOP 1000 CuentaCorriente.IDMovimiento, CuentaCorriente.FechaHora, CuentaCorriente.IDCuentaCorrienteGrupo, CuentaCorrienteGrupo.Nombre AS Grupo, CuentaCorriente.IDCuentaCorrienteCaja, CuentaCorrienteCaja.Nombre AS Caja, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, CuentaCorriente.Descripcion, ViajeDetalle.Realizado, Pasajero.Apellido + (CASE ISNULL(Pasajero.Nombre, '') WHEN '' THEN '' ELSE ', ' + Pasajero.Nombre END) AS Pasajero, CuentaCorriente.Importe, NULL
				FROM ((((CuentaCorriente INNER JOIN CSTransporte..CuentaCorrienteGrupo ON CuentaCorriente.IDCuentaCorrienteGrupo = CSTransporte..CuentaCorrienteGrupo.IDCuentaCorrienteGrupo) INNER JOIN CSTransporte..CuentaCorrienteCaja ON CuentaCorriente.IDCuentaCorrienteCaja = CSTransporte..CuentaCorrienteCaja.IDCuentaCorrienteCaja) LEFT JOIN CSTransporte..Persona ON CuentaCorriente.IDPersona = CSTransporte..Persona.IDPersona) LEFT JOIN CSTransporte..Persona AS Pasajero ON CuentaCorriente.IDPersonaOrigen = Pasajero.IDPersona) LEFT JOIN ViajeDetalle ON CuentaCorriente.Viaje_FechaHora = ViajeDetalle.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalle.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalle.Indice
				WHERE (@IDPersona IS NULL OR CuentaCorriente.IDPersona = @IDPersona)
					AND (@IDCuentaCorrienteGrupo IS NULL OR CuentaCorriente.IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo)
					AND (@IDCuentaCorrienteCaja IS NULL OR CuentaCorriente.IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja)
					AND ((@FilterTipo = 0) OR (@FilterTipo = 1 AND CuentaCorriente.Importe >= 0) OR (@FilterTipo = 2 AND CuentaCorriente.Importe < 0))
			 		AND CuentaCorriente.FechaHora BETWEEN @FechaDesde AND @FechaHasta
		END

	--FECHA MAYOR O IGUAL A
	IF @FechaFiltro = 2
		BEGIN
		INSERT INTO #CuentaCorriente
			(IDMovimiento, FechaHora, IDCuentaCorrienteGrupo, Grupo, IDCuentaCorrienteCaja, Caja, Persona, Descripcion, Realizado, Pasajero, Importe, ImporteAcumulado)
			SELECT TOP 1000 CuentaCorriente.IDMovimiento, CuentaCorriente.FechaHora, CuentaCorriente.IDCuentaCorrienteGrupo, CuentaCorrienteGrupo.Nombre AS Grupo, CuentaCorriente.IDCuentaCorrienteCaja, CuentaCorrienteCaja.Nombre AS Caja, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, CuentaCorriente.Descripcion, ViajeDetalle.Realizado, Pasajero.Apellido + (CASE ISNULL(Pasajero.Nombre, '') WHEN '' THEN '' ELSE ', ' + Pasajero.Nombre END) AS Pasajero, CuentaCorriente.Importe, NULL
				FROM ((((CuentaCorriente INNER JOIN CSTransporte..CuentaCorrienteGrupo ON CuentaCorriente.IDCuentaCorrienteGrupo = CSTransporte..CuentaCorrienteGrupo.IDCuentaCorrienteGrupo) INNER JOIN CSTransporte..CuentaCorrienteCaja ON CuentaCorriente.IDCuentaCorrienteCaja = CSTransporte..CuentaCorrienteCaja.IDCuentaCorrienteCaja) LEFT JOIN CSTransporte..Persona ON CuentaCorriente.IDPersona = CSTransporte..Persona.IDPersona) LEFT JOIN CSTransporte..Persona AS Pasajero ON CuentaCorriente.IDPersonaOrigen = Pasajero.IDPersona) LEFT JOIN ViajeDetalle ON CuentaCorriente.Viaje_FechaHora = ViajeDetalle.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalle.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalle.Indice
				WHERE (@IDPersona IS NULL OR CuentaCorriente.IDPersona = @IDPersona)
					AND (@IDCuentaCorrienteGrupo IS NULL OR CuentaCorriente.IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo)
					AND (@IDCuentaCorrienteCaja IS NULL OR CuentaCorriente.IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja)
					AND ((@FilterTipo = 0) OR (@FilterTipo = 1 AND CuentaCorriente.Importe >= 0) OR (@FilterTipo = 2 AND CuentaCorriente.Importe < 0))
					AND CuentaCorriente.FechaHora >= @FechaDesde
		END

	--FECHA MENOR O IGUAL A
	IF @FechaFiltro = 3
		BEGIN
		SET @FechaDesde = dateadd(minute, -1, dateadd(day, 1, @FechaDesde))
		INSERT INTO #CuentaCorriente
			(IDMovimiento, FechaHora, IDCuentaCorrienteGrupo, Grupo, IDCuentaCorrienteCaja, Caja, Persona, Descripcion, Realizado, Pasajero, Importe, ImporteAcumulado)
			SELECT TOP 1000 CuentaCorriente.IDMovimiento, CuentaCorriente.FechaHora, CuentaCorriente.IDCuentaCorrienteGrupo, CuentaCorrienteGrupo.Nombre AS Grupo, CuentaCorriente.IDCuentaCorrienteCaja, CuentaCorrienteCaja.Nombre AS Caja, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, CuentaCorriente.Descripcion, ViajeDetalle.Realizado, Pasajero.Apellido + (CASE ISNULL(Pasajero.Nombre, '') WHEN '' THEN '' ELSE ', ' + Pasajero.Nombre END) AS Pasajero, CuentaCorriente.Importe, NULL
				FROM ((((CuentaCorriente INNER JOIN CSTransporte..CuentaCorrienteGrupo ON CuentaCorriente.IDCuentaCorrienteGrupo = CSTransporte..CuentaCorrienteGrupo.IDCuentaCorrienteGrupo) INNER JOIN CSTransporte..CuentaCorrienteCaja ON CuentaCorriente.IDCuentaCorrienteCaja = CSTransporte..CuentaCorrienteCaja.IDCuentaCorrienteCaja) LEFT JOIN CSTransporte..Persona ON CuentaCorriente.IDPersona = CSTransporte..Persona.IDPersona) LEFT JOIN CSTransporte..Persona AS Pasajero ON CuentaCorriente.IDPersonaOrigen = Pasajero.IDPersona) LEFT JOIN ViajeDetalle ON CuentaCorriente.Viaje_FechaHora = ViajeDetalle.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalle.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalle.Indice
				WHERE (@IDPersona IS NULL OR CuentaCorriente.IDPersona = @IDPersona)
					AND (@IDCuentaCorrienteGrupo IS NULL OR CuentaCorriente.IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo)
					AND (@IDCuentaCorrienteCaja IS NULL OR CuentaCorriente.IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja)
					AND ((@FilterTipo = 0) OR (@FilterTipo = 1 AND CuentaCorriente.Importe >= 0) OR (@FilterTipo = 2 AND CuentaCorriente.Importe < 0))
					AND CuentaCorriente.FechaHora <= @FechaDesde
		END
		
	--FECHA DESDE Y HASTA
	IF @FechaFiltro = 4
		BEGIN
		SET @FechaHasta = dateadd(minute, -1, dateadd(day, 1, @FechaHasta))
		INSERT INTO #CuentaCorriente
			(IDMovimiento, FechaHora, IDCuentaCorrienteGrupo, Grupo, IDCuentaCorrienteCaja, Caja, Persona, Descripcion, Realizado, Pasajero, Importe, ImporteAcumulado)
			SELECT TOP 1000 CuentaCorriente.IDMovimiento, CuentaCorriente.FechaHora, CuentaCorriente.IDCuentaCorrienteGrupo, CuentaCorrienteGrupo.Nombre AS Grupo, CuentaCorriente.IDCuentaCorrienteCaja, CuentaCorrienteCaja.Nombre AS Caja, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, CuentaCorriente.Descripcion, ViajeDetalle.Realizado, Pasajero.Apellido + (CASE ISNULL(Pasajero.Nombre, '') WHEN '' THEN '' ELSE ', ' + Pasajero.Nombre END) AS Pasajero, CuentaCorriente.Importe, NULL
				FROM ((((CuentaCorriente INNER JOIN CSTransporte..CuentaCorrienteGrupo ON CuentaCorriente.IDCuentaCorrienteGrupo = CSTransporte..CuentaCorrienteGrupo.IDCuentaCorrienteGrupo) INNER JOIN CSTransporte..CuentaCorrienteCaja ON CuentaCorriente.IDCuentaCorrienteCaja = CSTransporte..CuentaCorrienteCaja.IDCuentaCorrienteCaja) LEFT JOIN CSTransporte..Persona ON CuentaCorriente.IDPersona = CSTransporte..Persona.IDPersona) LEFT JOIN CSTransporte..Persona AS Pasajero ON CuentaCorriente.IDPersonaOrigen = Pasajero.IDPersona) LEFT JOIN ViajeDetalle ON CuentaCorriente.Viaje_FechaHora = ViajeDetalle.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalle.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalle.Indice
				WHERE (@IDPersona IS NULL OR CuentaCorriente.IDPersona = @IDPersona)
					AND (@IDCuentaCorrienteGrupo IS NULL OR CuentaCorriente.IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo)
					AND (@IDCuentaCorrienteCaja IS NULL OR CuentaCorriente.IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja)
					AND ((@FilterTipo = 0) OR (@FilterTipo = 1 AND CuentaCorriente.Importe >= 0) OR (@FilterTipo = 2 AND CuentaCorriente.Importe < 0))
					AND CuentaCorriente.FechaHora BETWEEN @FechaDesde AND @FechaHasta
		END

GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTE_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_CuentaCorriente_List'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_CuentaCorriente_List
GO

CREATE PROCEDURE dbo.sp_CuentaCorriente_List
	@IDPersona int,
	@IDCuentaCorrienteGrupo int,
	@IDCuentaCorrienteCaja int,
	@FilterTipo int,
	@FechaFiltro int,
	@FechaDesde smalldatetime,
	@FechaHasta smalldatetime AS

	DECLARE @SaldoAnterior money

	SET NOCOUNT ON
	SET XACT_ABORT ON

	BEGIN TRANSACTION

	--TABLA TEMPORARIA
	IF EXISTS (SELECT name FROM sysobjects WHERE  name = N'#CuentaCorriente' AND type = 'U')
		BEGIN
		DELETE FROM #CuentaCorriente
		END
	ELSE
		BEGIN
		CREATE TABLE #CuentaCorriente
			(IDMovimiento int NOT NULL PRIMARY KEY, FechaHora smalldatetime NULL, IDCuentaCorrienteGrupo integer NULL, Grupo varchar(50) NULL, IDCuentaCorrienteCaja integer NULL, Caja varchar(50) NULL, Persona varchar(152) NULL, Descripcion varchar(255) NOT NULL, Realizado bit NULL, Pasajero varchar(152) NULL, Importe money NOT NULL, ImporteAcumulado money NULL)
		END

	--SALDO ANTERIOR
	IF @FechaFiltro > 0
		BEGIN
		IF @FechaFiltro <> 3
			BEGIN
			EXEC sp_CuentaCorriente_List_SaldoAnterior @IDPersona, @IDCuentaCorrienteGrupo, @IDCuentaCorrienteCaja, @FilterTipo, @FechaDesde, @SaldoAnterior OUTPUT
			INSERT INTO #CuentaCorriente
				(IDMovimiento, FechaHora, IDCuentaCorrienteGrupo, Grupo, IDCuentaCorrienteCaja, Caja, Persona, Descripcion, Realizado, Pasajero, Importe, ImporteAcumulado)
				VALUES (0, NULL,  NULL,  NULL, NULL, NULL, NULL, 'SALDO ANTERIOR', NULL, NULL, ISNULL(@SaldoAnterior, 0), ISNULL(@SaldoAnterior, 0))
			END
		END

	--MOVIMIENTOS
	EXEC sp_CuentaCorriente_List_Movimientos @IDPersona, @IDCuentaCorrienteGrupo, @IDCuentaCorrienteCaja, @FilterTipo, @FechaFiltro, @FechaDesde, @FechaHasta

	--DEVUELVO EL RESULTADO
	SET NOCOUNT OFF
	
	SELECT * FROM #CuentaCorriente ORDER BY FechaHora, IDMovimiento

	COMMIT TRANSACTION

	SET XACT_ABORT OFF

GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTE_LISTPERSONAL
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_CuentaCorriente_ListPersonal'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_CuentaCorriente_ListPersonal
GO

CREATE PROCEDURE dbo.sp_CuentaCorriente_ListPersonal
	@IDPersona integer,
	@IDCuentaCorrienteGrupo integer,
	@IDCuentaCorrienteCaja integer,
	@FilterTipo integer,
	@FechaFiltro integer,
	@FechaDesde smalldatetime,
	@FechaHasta smalldatetime AS

	DECLARE @SaldoAnterior money

	DECLARE @Resultset table(IDMovimiento int NOT NULL PRIMARY KEY, FechaHora smalldatetime NULL, IDCuentaCorrienteGrupo integer NULL, Grupo varchar(50) NULL, IDCuentaCorrienteCaja integer NULL, Caja varchar(50) NULL, Persona varchar(152) NULL, Descripcion varchar(255) NOT NULL, Realizado bit NULL, Pasajero varchar(152) NULL, Importe money NOT NULL, ImporteAcumulado money NULL)

	SET NOCOUNT ON
	SET XACT_ABORT ON

	BEGIN TRANSACTION

	--SALDO ANTERIOR
	IF @FechaFiltro > 0
		BEGIN
		IF @FechaFiltro <> 3
			BEGIN
			SET @SaldoAnterior = (SELECT SUM(CuentaCorriente.Importe)
				FROM ((CuentaCorriente INNER JOIN CSTransporte..CuentaCorrienteCaja ON CuentaCorriente.IDCuentaCorrienteCaja = CSTransporte..CuentaCorrienteCaja.IDCuentaCorrienteCaja) LEFT JOIN ViajeDetalle ON CuentaCorriente.Viaje_FechaHora = ViajeDetalle.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalle.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalle.Indice) LEFT JOIN Viaje ON CuentaCorriente.Viaje_FechaHora = Viaje.FechaHora AND CuentaCorriente.Viaje_IDRuta = Viaje.IDRuta AND CuentaCorriente.Viaje_Indice = 0
				WHERE (Viaje.Personal IS NULL OR Viaje.Personal = 0)
					AND CuentaCorriente.SaldoAnterior = 0
					AND (@IDPersona IS NULL OR CuentaCorriente.IDPersona = @IDPersona)
					AND (@IDCuentaCorrienteGrupo IS NULL OR CuentaCorriente.IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo)
					AND (@IDCuentaCorrienteCaja IS NULL OR CuentaCorriente.IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja)
					AND (@FilterTipo = 0 OR (@FilterTipo = 1 AND CuentaCorriente.Importe >= 0) OR (@FilterTipo = 2 AND CuentaCorriente.Importe < 0))
					AND CuentaCorriente.FechaHora < @FechaDesde)
			INSERT INTO @Resultset
				(IDMovimiento, FechaHora, IDCuentaCorrienteGrupo, Grupo, IDCuentaCorrienteCaja, Caja, Persona, Descripcion, Realizado, Pasajero, Importe, ImporteAcumulado)
				VALUES (0, NULL,  NULL, NULL,  NULL, NULL, NULL, 'SALDO ANTERIOR', NULL, NULL, ISNULL(@SaldoAnterior, 0), ISNULL(@SaldoAnterior, 0))
			END
		END

	INSERT INTO @Resultset
		(IDMovimiento, FechaHora, IDCuentaCorrienteGrupo, Grupo, IDCuentaCorrienteCaja, Caja, Persona, Descripcion, Realizado, Pasajero, Importe, ImporteAcumulado)
		SELECT TOP 1000 CuentaCorriente.IDMovimiento, CuentaCorriente.FechaHora, CuentaCorriente.IDCuentaCorrienteGrupo, CuentaCorrienteGrupo.Nombre AS Grupo, CuentaCorriente.IDCuentaCorrienteCaja, CuentaCorrienteCaja.Nombre AS Caja, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, CuentaCorriente.Descripcion, ViajeDetalle.Realizado, Pasajero.Apellido + (CASE ISNULL(Pasajero.Nombre, '') WHEN '' THEN '' ELSE ', ' + Pasajero.Nombre END) AS Pasajero, CuentaCorriente.Importe, NULL
			FROM (((((CuentaCorriente INNER JOIN CSTransporte..CuentaCorrienteGrupo ON CuentaCorriente.IDCuentaCorrienteGrupo = CSTransporte..CuentaCorrienteGrupo.IDCuentaCorrienteGrupo) INNER JOIN CSTransporte..CuentaCorrienteCaja ON CuentaCorriente.IDCuentaCorrienteCaja = CSTransporte..CuentaCorrienteCaja.IDCuentaCorrienteCaja) LEFT JOIN CSTransporte..Persona ON CuentaCorriente.IDPersona = CSTransporte..Persona.IDPersona) LEFT JOIN CSTransporte..Persona AS Pasajero ON CuentaCorriente.IDPersonaOrigen = Pasajero.IDPersona)  LEFT JOIN ViajeDetalle ON CuentaCorriente.Viaje_FechaHora = ViajeDetalle.FechaHora AND CuentaCorriente.Viaje_IDRuta = ViajeDetalle.IDRuta AND CuentaCorriente.Viaje_Indice = ViajeDetalle.Indice) LEFT JOIN Viaje ON CuentaCorriente.Viaje_FechaHora = Viaje.FechaHora AND CuentaCorriente.Viaje_IDRuta = Viaje.IDRuta
			WHERE (Viaje.Personal IS NULL OR Viaje.Personal = 0)
				AND CuentaCorriente.SaldoAnterior = 0
				AND (@IDPersona IS NULL OR CuentaCorriente.IDPersona = @IDPersona)
				AND (@IDCuentaCorrienteGrupo IS NULL OR CuentaCorriente.IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo)
				AND (@IDCuentaCorrienteCaja IS NULL OR CuentaCorriente.IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja)
				AND (@FilterTipo = 0 OR (@FilterTipo = 1 AND CuentaCorriente.Importe >= 0) OR (@FilterTipo = 2 AND CuentaCorriente.Importe < 0))
			 	AND (@FechaFiltro = 0	--ALL
					OR (@FechaFiltro = 1 AND CuentaCorriente.FechaHora BETWEEN @FechaDesde AND dateadd(minute, -1, dateadd(day, 1, @FechaDesde)))	--EQUAL
					OR (@FechaFiltro = 2 AND CuentaCorriente.FechaHora >= @FechaDesde)	--GREATER OR EQUAL
					OR (@FechaFiltro = 3 AND CuentaCorriente.FechaHora <= dateadd(minute, -1, dateadd(day, 1, @FechaDesde)))	--MINOR OR EQUAL
					OR (@FechaFiltro = 4 AND CuentaCorriente.FechaHora BETWEEN @FechaDesde AND dateadd(minute, -1, dateadd(day, 1, @FechaHasta)))	--BETWEEN
					)
			ORDER BY CuentaCorriente.FechaHora, CuentaCorriente.IDMovimiento

	--SALDO ACUMULADO
	DECLARE @IDMovimiento integer
	DECLARE @Importe money
	DECLARE @ImporteAcumulado money
	
	DECLARE ResultsetCursor
		CURSOR LOCAL FORWARD_ONLY STATIC
		FOR SELECT IDMovimiento, Importe
			FROM @Resultset

	SET @ImporteAcumulado = 0
	OPEN ResultsetCursor
	FETCH NEXT FROM ResultsetCursor INTO @IDMovimiento, @Importe
	WHILE @@FETCH_STATUS = 0
		BEGIN
		SET @ImporteAcumulado = @ImporteAcumulado + @Importe
		UPDATE @Resultset
			SET ImporteAcumulado = @ImporteAcumulado
			WHERE IDMovimiento = @IDMovimiento
		FETCH NEXT FROM ResultsetCursor INTO @IDMovimiento, @Importe
	END
	CLOSE ResultsetCursor
	DEALLOCATE ResultsetCursor

	SET NOCOUNT OFF
	
	SELECT * FROM @Resultset

	COMMIT TRANSACTION

	SET XACT_ABORT OFF

GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTE_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name
	   FROM   sysobjects 
	   WHERE  name = N'sp_CuentaCorriente_Data'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_CuentaCorriente_Data
GO

CREATE PROCEDURE dbo.sp_CuentaCorriente_Data
	@IDMovimiento int AS

	SELECT IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM CuentaCorriente
		WHERE IDMovimiento = @IDMovimiento

GO