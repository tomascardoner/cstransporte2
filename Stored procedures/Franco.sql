------------------------------------------------------------------------------------------
-- FRANCO_UPDATE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'sp_Franco_Update' AND type = 'P')
    DROP PROCEDURE sp_Franco_Update
GO

CREATE PROCEDURE dbo.sp_Franco_Update 
	@FechaOriginal smalldatetime,
	@IDPersonaOriginal int,
	@Fecha smalldatetime,
	@IDPersona int,
	@Importe smallmoney,
	@IDUsuario smallint AS

	SET XACT_ABORT ON

	BEGIN TRANSACTION

	DECLARE @IDMovimientoCuentaCorriente int
	DECLARE @IDCuentaCorrienteGrupo int
	DECLARE @IDCuentaCorrienteCaja int

	IF @IDPersonaOriginal = 0
		BEGIN
		-- Es un franco nuevo
		IF @Importe IS NOT NULL
			BEGIN
			-- El importe no es null, entonces, agrego el movimiento de cuenta corriente
			SET @IDCuentaCorrienteGrupo = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'CuentaCorrienteGrupo_ID_Sueldo')
			SET @IDCuentaCorrienteCaja = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'CuentaCorrienteCaja_ID_ViajeDebito')
			SET @IDMovimientoCuentaCorriente = (SELECT ISNULL(MAX(IDMovimiento), 0) + 1 FROM CuentaCorriente)
			INSERT INTO CuentaCorriente
				(IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)
				VALUES (@IDMovimientoCuentaCorriente, @IDCuentaCorrienteGrupo, @IDCuentaCorrienteCaja, @IDPersona, getdate(), 'Sueldo por Franco: ' + CONVERT(char(10), @Fecha, 103), @Importe, NULL, NULL, 0, NULL, NULL, NULL, getdate(), @IDUsuario, getdate(), @IDUsuario)
			END

		-- Agrego el franco
		INSERT INTO Franco
			(Fecha, IDPersona, Importe, IDMovimientoCuentaCorriente, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)
			VALUES (@Fecha, @IDPersona, @Importe, @IDMovimientoCuentaCorriente, getdate(), @IDUsuario, getdate(), @IDUsuario)
		END
	ELSE
		BEGIN
		-- Es un franco existente
		-- Busco el movimiento de cuenta corriente original
		SET @IDMovimientoCuentaCorriente = (SELECT ISNULL(IDMovimientoCuentaCorriente, 0) FROM Franco WHERE Fecha = @FechaOriginal AND IDPersona = @IDPersonaOriginal)
		IF @Importe IS NULL
			BEGIN
			-- Borro el movimiento de cuenta corriente porque ya no es necesario
			IF @IDMovimientoCuentaCorriente > 0
				BEGIN
				UPDATE Franco
					SET IDMovimientoCuentaCorriente = NULL
					WHERE Fecha = @FechaOriginal AND IDPersona = @IDPersonaOriginal
				DELETE
					FROM CuentaCorriente
					WHERE IDMovimiento = @IDMovimientoCuentaCorriente
				SET @IDMovimientoCuentaCorriente = NULL
				END
			END
		ELSE
			BEGIN
			IF @IDMovimientoCuentaCorriente = 0
				BEGIN
				-- Agrego el movimiento de cuenta corriente
				SET @IDCuentaCorrienteGrupo = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'CuentaCorrienteGrupo_ID_Sueldo')
				SET @IDCuentaCorrienteCaja = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'CuentaCorrienteCaja_ID_ViajeDebito')
				SET @IDMovimientoCuentaCorriente = (SELECT ISNULL(MAX(IDMovimiento), 0) + 1 FROM CuentaCorriente)
				INSERT INTO CuentaCorriente
					(IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)
					VALUES (@IDMovimientoCuentaCorriente, @IDCuentaCorrienteGrupo, @IDCuentaCorrienteCaja, @IDPersona, getdate(), 'Sueldo por Franco: ' + CONVERT(char(10), @Fecha, 103), @Importe, NULL, NULL, 0, NULL, NULL, NULL, getdate(), @IDUsuario, getdate(), @IDUsuario)
				END
			ELSE
				-- Actualizo el movimiento de cuenta corriente
				UPDATE CuentaCorriente
					SET IDPersona = @IDPersona, Descripcion = 'Sueldo por Franco: ' + CONVERT(char(10), @Fecha, 103), Importe = @Importe, FechaHoraModificacion = getdate(), IDUsuarioModificacion = @IDUsuario
					WHERE IDMovimiento = @IDMovimientoCuentaCorriente
			END

		-- Actualizo el franco
		UPDATE Franco
			SET Fecha = @Fecha, IDPersona = @IDPersona, Importe = @Importe, IDMovimientoCuentaCorriente = @IDMovimientoCuentaCorriente, FechaHoraModificacion = getdate(), IDUsuarioModificacion = @IDUsuario
			WHERE Fecha = @FechaOriginal AND IDPersona = @IDPersonaOriginal
		END

	COMMIT TRANSACTION

	SET XACT_ABORT OFF

GO