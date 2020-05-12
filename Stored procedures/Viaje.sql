------------------------------------------------------------------------------------------
-- VIAJE_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Viaje_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Viaje_Data
GO

CREATE PROCEDURE dbo.sp_Viaje_Data
	@FechaHora_FILTER smalldatetime,
	@IDRuta_FILTER char(20) AS

	SELECT IDViaje, FechaHora, IDRuta, RutaOtra, IDPersona, Kilometro, Duracion, Importe, ImporteContado, IDMedioPago, Cuotas, Operacion, IDCuentaCorrienteCaja, Charter, IDVehiculo, IDConductor, AcreditaSueldo, IDConductor2, AcreditaSueldo2, DiaSemanaBase, Estado, AsientoOcupado, Notas, Personal, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion, FechaHoraEnProgreso, IDUsuarioEnProgreso, FechaHoraFinalizado, IDUsuarioFinalizado, FechaHoraCancelado, IDUsuarioCancelado
		FROM Viaje
		WHERE FechaHora = @FechaHora_FILTER AND IDRuta = @IDRuta_FILTER

GO



------------------------------------------------------------------------------------------
-- VIAJE_INSERT
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Viaje_Insert'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Viaje_Insert
GO

CREATE PROCEDURE dbo.sp_Viaje_Insert
    @IDViaje int OUTPUT,
	@FechaHora smalldatetime,
	@IDRuta char(20),
	@RutaOtra varchar(50),
	@IDPersona int,
	@Kilometro smallint,
	@Duracion smallint,
	@Importe money,
	@ImporteContado money,
	@IDMedioPago tinyint,
	@Cuotas tinyint,
	@Operacion varchar(20),
	@IDCuentaCorrienteCaja int,
	@Charter bit,
	@IDVehiculo int,
	@IDConductor int,
	@AcreditaSueldo bit,
	@IDConductor2 int,
	@AcreditaSueldo2 bit,
	@DiaSemanaBase tinyint,
	@Estado char(2),
	@Notas varchar(8000),
	@Personal bit,
	@IDUsuarioCreacion smallint AS

	INSERT INTO Viaje
		(FechaHora, IDRuta, RutaOtra, IDPersona, Kilometro, Duracion, Importe, ImporteContado, IDMedioPago, Cuotas, Operacion, IDCuentaCorrienteCaja, Charter, IDVehiculo, IDConductor, AcreditaSueldo, IDConductor2, AcreditaSueldo2, DiaSemanaBase, Estado, AsientoOcupado, Notas, Personal, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion, FechaHoraEnProgreso, IDUsuarioEnProgreso, FechaHoraFinalizado, IDUsuarioFinalizado, FechaHoraCancelado, IDUsuarioCancelado)
		VALUES (@FechaHora, @IDRuta, @RutaOtra, @IDPersona, @Kilometro, @Duracion, @Importe, @ImporteContado, @IDMedioPago, @Cuotas, @Operacion, @IDCuentaCorrienteCaja, @Charter, @IDVehiculo, @IDConductor, @AcreditaSueldo, @IDConductor2, @AcreditaSueldo2, @DiaSemanaBase, @Estado, 0, @Notas, @Personal, getdate(), @IDUsuarioCreacion, getdate(), @IDUsuarioCreacion, NULL, NULL, NULL, NULL, NULL, NULL)
	
    SELECT @IDViaje = SCOPE_IDENTITY()
GO



------------------------------------------------------------------------------------------
-- VIAJE_UPDATE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'sp_Viaje_Update' AND type = 'P')
    DROP PROCEDURE sp_Viaje_Update
GO

CREATE PROCEDURE dbo.sp_Viaje_Update
	@FechaHora_Original smalldatetime,
	@IDRuta_Original char(20),
	@FechaHora smalldatetime,
	@IDRuta char(20),
	@RutaOtra varchar(50),
	@IDPersona int,
	@Kilometro smallint,
	@Duracion smallint,
	@Importe money,
	@ImporteContado money,
	@IDMedioPago tinyint,
	@Cuotas tinyint,
	@Operacion varchar(20),
	@IDCuentaCorrienteCaja int,
	@Charter bit,
	@IDVehiculo int,
	@IDConductor int,
	@AcreditaSueldo bit,
	@IDConductor2 int,
	@AcreditaSueldo2 bit,
	@DiaSemanaBase tinyint,
	@Estado char(2),
	@Notas varchar(8000),
	@Personal bit,
	@IDUsuarioModificacion smallint,
	@FechaHoraEnProgreso smalldatetime,
	@IDUsuarioEnProgreso smallint,
	@FechaHoraFinalizado smalldatetime,
	@IDUsuarioFinalizado smallint,
	@FechaHoraCancelado smalldatetime,
	@IDUsuarioCancelado smallint AS

	SET XACT_ABORT ON

	BEGIN TRANSACTION

	ALTER TABLE CuentaCorriente NOCHECK CONSTRAINT FK__Viaje__CuentaCorriente

	UPDATE Viaje
		SET FechaHora = @FechaHora, IDRuta = @IDRuta, RutaOtra = @RutaOtra, IDPersona = @IDPersona, Kilometro = @Kilometro, Duracion = @Duracion, Importe = @Importe, ImporteContado = @ImporteContado, IDMedioPago = @IDMedioPago, Cuotas = @Cuotas, Operacion = @Operacion, IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja, Charter = @Charter, IDVehiculo = @IDVehiculo, IDConductor = @IDConductor, AcreditaSueldo = @AcreditaSueldo, IDConductor2 = @IDConductor2, AcreditaSueldo2 = @AcreditaSueldo2, DiaSemanaBase = @DiaSemanaBase, Estado = @Estado, AsientoOcupado = 0, Notas = @Notas, Personal = @Personal, FechaHoraModificacion = getdate(), IDUsuarioModificacion = @IDUsuarioModificacion, FechaHoraEnProgreso = @FechaHoraEnProgreso, IDUsuarioEnProgreso = @IDUsuarioEnProgreso, FechaHoraFinalizado = @FechaHoraFinalizado, IDUsuarioFinalizado = @IDUsuarioFinalizado, FechaHoraCancelado = @FechaHoraCancelado, IDUsuarioCancelado = @IDUsuarioCancelado
		WHERE FechaHora = @FechaHora_Original AND IDRuta = @IDRuta_Original

	IF @FechaHora_Original <> @FechaHora OR @IDRuta_Original <> @IDRuta
		BEGIN
		UPDATE CuentaCorriente
			SET Viaje_FechaHora = @FechaHora, Viaje_IDRuta = @IDRuta
			WHERE Viaje_FechaHora = @FechaHora_Original AND Viaje_IDRuta = @IDRuta_Original
		END

	ALTER TABLE CuentaCorriente WITH CHECK CHECK CONSTRAINT FK__Viaje__CuentaCorriente

	COMMIT TRANSACTION

	SET XACT_ABORT OFF

GO



------------------------------------------------------------------------------------------
-- VIAJE_DELETE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Viaje_Delete'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Viaje_Delete
GO

CREATE PROCEDURE dbo.sp_Viaje_Delete
	@FechaHora smalldatetime,
	@IDRuta char(20) AS

	SET XACT_ABORT ON
	
	BEGIN TRANSACTION

    DECLARE @IDViajeDetalle int

	--ELIMINO LOS DETALLES DEL VIAJE CONEXION
	DELETE ViajeDetalle_Conexion
		FROM ViajeDetalle_Conexion INNER JOIN ViajeDetalle ON ViajeDetalle_Conexion.IDViajeDetalle = ViajeDetalle.IDViajeDetalle
		WHERE ViajeDetalle.FechaHora = @FechaHora AND ViajeDetalle.IDRuta = @IDRuta
	DELETE ViajeDetalle_Conexion
		FROM ViajeDetalle_Conexion INNER JOIN ViajeDetalle ON ViajeDetalle_Conexion.Conexion_IDViajeDetalle = ViajeDetalle.IDViajeDetalle
		WHERE ViajeDetalle.FechaHora = @FechaHora AND ViajeDetalle.IDRuta = @IDRuta

    --ELIMINO LOS DETALLES DEL VIAJE
	DELETE FROM ViajeDetalle
		WHERE FechaHora = @FechaHora AND IDRuta = @IDRuta

	--ELIMINO LOS DETALLES DEL VIAJE COMISION
	DELETE FROM ViajeDetalle_Comision
		WHERE FechaHora = @FechaHora AND IDRuta = @IDRuta

	--ELIMINO LOS MOVIMIENTOS DE LA CUENTA CORRIENTE
	DELETE FROM CuentaCorriente
		WHERE Viaje_FechaHora = @FechaHora AND Viaje_IDRuta = @IDRuta

	--ELIMINO EL VIAJE
	DELETE FROM Viaje
		WHERE FechaHora = @FechaHora AND IDRuta = @IDRuta

	COMMIT TRANSACTION

	SET XACT_ABORT OFF

GO
