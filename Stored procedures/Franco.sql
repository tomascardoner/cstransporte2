------------------------------------------------------------------------------------------
-- FRANCO_UPDATE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Franco_Update' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Franco_Update
GO

CREATE PROCEDURE dbo.sp_Franco_Update 
	@FechaOriginal smalldatetime,
	@IDPersonaOriginal int,
	@Fecha smalldatetime,
	@IDPersona int,
	@Importe smallmoney,
	@IDMovimientoCuentaCorriente int,
	@IDUsuario smallint AS

	SET XACT_ABORT ON

	BEGIN TRANSACTION

	DECLARE @IDMovimiento int
	DECLARE @IDCuentaCorrienteGrupo int
	DECLARE @IDCuentaCorrienteCaja int

	SET @IDMovimiento = (SELECT ISNULL(IDMovimientoCuentaCorriente, 0) FROM Franco WHERE Fecha = @FechaOriginal AND IDPersona = @IDPersonaOriginal)

	IF @IDMovimientoCuentaCorriente = 0
		BEGIN
		SET @IDCuentaCorrienteGrupo = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'CuentaCorrienteGrupo_ID_Sueldo')
		SET @IDCuentaCorrienteCaja = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'CuentaCorrienteCaja_ID_ViajeDebito')

		SET @IDMovimiento = (SELECT ISNULL(MAX(IDMovimiento), 0) + 1 FROM CuentaCorriente)
		INSERT INTO CuentaCorriente
			(IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)
			VALUES (@IDMovimiento, @IDCuentaCorrienteGrupo, @IDCuentaCorrienteCaja, @IDPersona, getdate(), 'Sueldo por Franco: ' + CONVERT(char(10), @Fecha, 103), @Importe, NULL, NULL, 0, NULL, NULL, NULL, getdate(), @IDUsuario, getdate(), @IDUsuario)

		INSERT INTO Franco
			(Fecha, IDPersona, Importe, IDMovimientoCuentaCorriente, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)
			VALUES (@Fecha, @IDPersona, @Importe, @IDMovimiento, getdate(), @IDUsuario, getdate(), @IDUsuario)
		END
	ELSE
		BEGIN
		UPDATE CuentaCorriente
			SET IDPersona = @IDPersona, Descripcion = 'Sueldo por Franco: ' + CONVERT(char(10), @Fecha, 103), Importe = @Importe, FechaHoraModificacion = getdate(), IDUsuarioModificacion = @IDUsuario
			WHERE IDMovimiento = @IDMovimiento

		UPDATE Franco
			SET Fecha = @Fecha, IDPersona = @IDPersona, Importe = @Importe, FechaHoraModificacion = getdate(), IDUsuarioModificacion = @IDUsuario
			WHERE Fecha = @FechaOriginal AND IDPersona = @IDPersonaOriginal
		END

	COMMIT TRANSACTION

	SET XACT_ABORT OFF

GO



