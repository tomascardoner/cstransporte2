USE CSTransporte_History 
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

	SELECT IDMovimiento, IDCuentaCorrienteGrupo, IDCuentaCorrienteCaja, IDPersona, FechaHora, Descripcion, Importe, IDPersonaOrigen, Notas, SaldoAnterior, Viaje_FechaHora, Viaje_IDRuta, Viaje_Indice, Viaje_ConductorNumero, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM CuentaCorriente
		WHERE IDMovimiento = @IDMovimiento

GO