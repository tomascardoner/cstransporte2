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

	SELECT IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDMedioPago, Cuotas, Operacion, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, Viaje_ConductorNumero, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM CuentaCorriente
		WHERE IDMovimiento = @IDMovimiento

GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTE_UPDATE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_CuentaCorriente_Update'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_CuentaCorriente_Update
GO

CREATE PROCEDURE dbo.sp_CuentaCorriente_Update
	@IDMovimiento int OUTPUT,
	@IDCuentaCorrienteGrupo int,
	@IDCuentaCorrienteCaja int,
	@IDPersona int,
	@FechaHora smalldatetime,
	@Descripcion varchar(255),
	@Importe money,
	@IDMedioPago tinyint,
	@Cuotas tinyint,
	@Operacion varchar(20), 	
	@IDPersonaOrigen int,
	@Notas varchar(8000),
	@SaldoAnterior bit,
	@Viaje_FechaHora smalldatetime,
	@Viaje_IDRuta char(20),
	@Viaje_Indice int,
	@Viaje_ConductorNumero tinyint,
	@IDUsuario smallint,
	@PermiteMultiples bit AS

	SET XACT_ABORT ON

	BEGIN TRANSACTION
	
	SET @IDMedioPago = ISNULL(@IDMedioPago, 1)

	IF @PermiteMultiples = 0 AND @IDMovimiento = 0
		BEGIN
		IF (@Viaje_FechaHora IS NOT NULL) AND (@Viaje_IDRuta IS NOT NULL)
			BEGIN
			SET @IDMovimiento = ISNULL((SELECT IDMovimiento FROM CuentaCorriente WHERE Viaje_FechaHora = @Viaje_FechaHora AND Viaje_IDRuta = @Viaje_IDRuta AND ((Viaje_Indice IS NULL AND @Viaje_Indice IS NULL) OR (Viaje_Indice = @Viaje_Indice)) AND ((Viaje_ConductorNumero IS NULL AND @Viaje_ConductorNumero IS NULL) OR (Viaje_ConductorNumero = @Viaje_ConductorNumero)) AND IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo), 0)
			END
		END

	IF @IDMovimiento = 0
		BEGIN
		SET @IDMovimiento = (SELECT ISNULL(MAX(IDMovimiento), 0) + 1 FROM CuentaCorriente)
		INSERT INTO CuentaCorriente
			(IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, IDMedioPago, Cuotas, Operacion, Importe, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, Viaje_ConductorNumero, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)
			VALUES (@IDMovimiento, @IDCuentaCorrienteGrupo, @IDCuentaCorrienteCaja, @IDPersona, @FechaHora, @Descripcion, @IDMedioPago, @Cuotas, @Operacion, @Importe, @IDPersonaOrigen, @Notas, @SaldoAnterior, @Viaje_FechaHora, @Viaje_IDRuta, @Viaje_Indice, @Viaje_ConductorNumero, getdate(), @IDUsuario, getdate(), @IDUsuario)
		END
	ELSE
		BEGIN
		UPDATE CuentaCorriente
			SET IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja, IDPersona = @IDPersona, FechaHora = @FechaHora, Descripcion = @Descripcion, Importe = @Importe, IDMedioPago = @IDMedioPago, Cuotas = @Cuotas, Operacion = @Operacion, IDPersonaOrigen = @IDPersonaOrigen, Notas = @Notas, SaldoAnterior = @SaldoAnterior, Viaje_FechaHora = @Viaje_FechaHora, Viaje_IDRuta = @Viaje_IDRuta, Viaje_Indice = @Viaje_Indice, Viaje_ConductorNumero = @Viaje_ConductorNumero, FechaHoraModificacion = getdate(), IDUsuarioModificacion = @IDUsuario
			WHERE IDMovimiento = @IDMovimiento
		END

	COMMIT TRANSACTION

	SET XACT_ABORT OFF

GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTE_TRANSFERENCIA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_CuentaCorriente_Transferencia'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_CuentaCorriente_Transferencia
GO

CREATE PROCEDURE dbo.sp_CuentaCorriente_Transferencia
	@IDCuentaCorrienteGrupo_Origen int,
	@IDCuentaCorrienteCaja_Origen int,
	@Descripcion_Origen varchar(255),
	@IDCuentaCorrienteGrupo_Destino int,
	@IDCuentaCorrienteCaja_Destino int,
	@Descripcion_Destino varchar(255),
	@FechaHora smalldatetime,
	@Importe money,
	@IDMedioPago tinyint,
	@IDUsuario smallint AS

	SET XACT_ABORT ON

	BEGIN TRANSACTION

	DECLARE @IDMovimiento int

	--ORIGEN
	SET @IDMovimiento = (SELECT ISNULL(MAX(IDMovimiento), 0) + 1 FROM CuentaCorriente)
	INSERT INTO CuentaCorriente
		(IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDMedioPago, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)
		VALUES (@IDMovimiento, @IDCuentaCorrienteGrupo_Origen, @IDCuentaCorrienteCaja_Origen, NULL, @FechaHora, @Descripcion_Origen, @Importe * -1, @IDMedioPago, NULL, NULL, 0, NULL, NULL, NULL, getdate(), @IDUsuario, getdate(), @IDUsuario)

	--DESTINO
	SET @IDMovimiento = (SELECT ISNULL(MAX(IDMovimiento), 0) + 1 FROM CuentaCorriente)
	INSERT INTO CuentaCorriente
		(IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDMedioPago, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)
		VALUES (@IDMovimiento, @IDCuentaCorrienteGrupo_Destino, @IDCuentaCorrienteCaja_Destino, NULL, @FechaHora, @Descripcion_Destino, @Importe, @IDMedioPago, NULL, NULL, 0, NULL, NULL, NULL, getdate(), @IDUsuario, getdate(), @IDUsuario)

	COMMIT TRANSACTION

	SET XACT_ABORT OFF

GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTE_DATA_VIAJE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_CuentaCorriente_DataByViaje' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_CuentaCorriente_DataByViaje
GO

CREATE PROCEDURE dbo.sp_CuentaCorriente_DataByViaje
	@IDCuentaCorrienteGrupo_FILTER integer,
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20) AS

	SELECT CuentaCorriente.IDMovimiento, CuentaCorriente.IDCuentaCorrienteGrupo, CuentaCorriente.IDCuentaCorrienteCaja, CuentaCorriente.IDPersona, CuentaCorriente.FechaHora, CuentaCorriente.Descripcion, CuentaCorriente.Importe, CuentaCorriente.IDMedioPago, CuentaCorriente.Cuotas, CuentaCorriente.Operacion, CuentaCorriente.IDPersonaOrigen, CuentaCorriente.Notas, CuentaCorriente.SaldoAnterior, CuentaCorriente.Viaje_FechaHora, CuentaCorriente.Viaje_IDRuta, CuentaCorriente.Viaje_Indice, CuentaCorriente.FechaHoraCreacion, CuentaCorriente.IDUsuarioCreacion, CuentaCorriente.FechaHoraModificacion, CuentaCorriente.IDUsuarioModificacion
		FROM CuentaCorriente
		WHERE CuentaCorriente.IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo_FILTER AND CuentaCorriente.Viaje_FechaHora = @FechaHora_FILTER AND CuentaCorriente.Viaje_IDRuta = @IDRuta_FILTER AND CuentaCorriente.Viaje_Indice = 0

GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTE_DATA_VIAJECONDUCTORNUMERO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_CuentaCorriente_DataByViajeConductorNumero' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_CuentaCorriente_DataByViajeConductorNumero
GO

CREATE PROCEDURE dbo.sp_CuentaCorriente_DataByViajeConductorNumero
	@IDCuentaCorrienteGrupo_FILTER integer,
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@ConductorNumero_FILTER tinyint AS

	SELECT IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDMedioPago, Cuotas, Operacion, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, Viaje_ConductorNumero, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM CuentaCorriente
		WHERE IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo_FILTER AND Viaje_FechaHora = @FechaHora_FILTER AND Viaje_IDRuta = @IDRuta_FILTER AND Viaje_Indice = 0 AND Viaje_ConductorNumero = @ConductorNumero_FILTER

GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTE_DATA_VIAJEDETALLE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_CuentaCorriente_DataByViajeDetalle'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_CuentaCorriente_DataByViajeDetalle
GO

CREATE PROCEDURE dbo.sp_CuentaCorriente_DataByViajeDetalle
	@IDCuentaCorrienteGrupo_FILTER integer,
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20),
	@Indice_FILTER integer AS

	SELECT CuentaCorriente.IDMovimiento, CuentaCorriente.IDCuentaCorrienteGrupo, CuentaCorriente.IDCuentaCorrienteCaja, CuentaCorriente.IDPersona, CuentaCorriente.FechaHora, CuentaCorriente.Descripcion, CuentaCorriente.Importe, CuentaCorriente.IDMedioPago, CuentaCorriente.Cuotas, CuentaCorriente.Operacion, CuentaCorriente.IDPersonaOrigen, CuentaCorriente.Notas, CuentaCorriente.SaldoAnterior, CuentaCorriente.Viaje_FechaHora, CuentaCorriente.Viaje_IDRuta, CuentaCorriente.Viaje_Indice, CuentaCorriente.FechaHoraCreacion, CuentaCorriente.IDUsuarioCreacion, CuentaCorriente.FechaHoraModificacion, CuentaCorriente.IDUsuarioModificacion
		FROM CuentaCorriente
		WHERE CuentaCorriente.IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo_FILTER AND CuentaCorriente.Viaje_FechaHora = @FechaHora_FILTER AND CuentaCorriente.Viaje_IDRuta = @IDRuta_FILTER AND CuentaCorriente.Viaje_Indice = @Indice_FILTER

GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTECAJA_SALDOACTUAL
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_CuentaCorrienteCaja_SaldoActual'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_CuentaCorrienteCaja_SaldoActual
GO

CREATE PROCEDURE dbo.sp_CuentaCorrienteCaja_SaldoActual
	@IDCuentaCorrienteCaja int,
	@SaldoActual_Efectivo money OUTPUT,
	@SaldoActual_Tarjeta money OUTPUT AS

BEGIN
	DECLARE @CuentaCorrienteCaja_ID_ViajeDebito int

	SET NOCOUNT ON;

	SET @CuentaCorrienteCaja_ID_ViajeDebito = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'CuentaCorrienteCaja_ID_ViajeDebito')

	SET @SaldoActual_Efectivo = (SELECT ISNULL(SUM(CuentaCorriente.Importe), 0)
									FROM CuentaCorriente INNER JOIN MedioPago ON CuentaCorriente.IDMedioPago = MedioPago.IDMedioPago
									WHERE MedioPago.UtilizaOperacion = 0 AND CuentaCorriente.IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja AND (CuentaCorriente.IDCuentaCorrienteCaja <> @CuentaCorrienteCaja_ID_ViajeDebito OR CuentaCorriente.Importe > 0))

	SET @SaldoActual_Tarjeta = (SELECT ISNULL(SUM(CuentaCorriente.Importe), 0)
									FROM CuentaCorriente INNER JOIN MedioPago ON CuentaCorriente.IDMedioPago = MedioPago.IDMedioPago
									WHERE MedioPago.UtilizaOperacion = 1 AND CuentaCorriente.IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja AND (CuentaCorriente.IDCuentaCorrienteCaja <> @CuentaCorrienteCaja_ID_ViajeDebito OR CuentaCorriente.Importe > 0))

END
GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTEGRUPO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_CuentaCorrienteGrupo_Data' AND type = 'P')
    DROP PROCEDURE usp_CuentaCorrienteGrupo_Data
GO

CREATE PROCEDURE dbo.usp_CuentaCorrienteGrupo_Data 
	@IDCuentaCorrienteGrupo int AS

	SELECT IDCuentaCorrienteGrupo, Nombre, Notas, Ocultar, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM CuentaCorrienteGrupo
		WHERE IDCuentaCorrienteGrupo = @IDCuentaCorrienteGrupo

GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTEGRUPO_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_CuentaCorrienteGrupo_IDMax' AND type = 'P')
    DROP PROCEDURE usp_CuentaCorrienteGrupo_IDMax
GO

CREATE PROCEDURE dbo.usp_CuentaCorrienteGrupo_IDMax AS
	SELECT Max(IDCuentaCorrienteGrupo) AS IDCuentaCorrienteGrupoMax
	FROM CuentaCorrienteGrupo
	 
GO



------------------------------------------------------------------------------------------
-- CUENTACORRIENTECAJA_SALDOACTUAL_TARJETA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_CuentaCorrienteCaja_SaldoActual_Tarjeta'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_CuentaCorrienteCaja_SaldoActual_Tarjeta
GO

CREATE PROCEDURE dbo.sp_CuentaCorrienteCaja_SaldoActual_Tarjeta
	@IDCuentaCorrienteCaja int AS

BEGIN
	DECLARE @CuentaCorrienteCaja_ID_ViajeDebito int

	SET NOCOUNT ON;

	SET @CuentaCorrienteCaja_ID_ViajeDebito = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'CuentaCorrienteCaja_ID_ViajeDebito')

	SELECT MedioPago.IDMedioPago, MedioPago.Nombre AS MedioPagoNombre, ISNULL(SUM(CuentaCorriente.Importe), 0) AS SaldoActual, 0 AS Transferir
		FROM CuentaCorriente INNER JOIN MedioPago ON CuentaCorriente.IDMedioPago = MedioPago.IDMedioPago
		WHERE MedioPago.UtilizaOperacion = 1 AND CuentaCorriente.IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja AND (CuentaCorriente.IDCuentaCorrienteCaja <> @CuentaCorrienteCaja_ID_ViajeDebito OR CuentaCorriente.Importe > 0)
		GROUP BY MedioPago.IDMedioPago, MedioPago.Nombre

END
GO