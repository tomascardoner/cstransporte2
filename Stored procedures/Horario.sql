------------------------------------------------------------------------------------------
-- HORARIO_LIST_DIASEMANA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Horario_List_DiaSemana' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Horario_List_DiaSemana
GO

CREATE PROCEDURE dbo.sp_Horario_List_DiaSemana
	@DiaSemana_FILTER tinyint AS

	SELECT DiaSemana, Hora, IDRuta, IDConductor, IDConductor2, IDVehiculo, ConductorImporteTramoCompleto, ConductorImporteTramo1, ConductorImporteTramo2, Notas, Personal, FechaHoraCreacion
		FROM Horario
		WHERE Horario.DiaSemana = @DiaSemana_FILTER AND Horario.Activo = 1

GO
