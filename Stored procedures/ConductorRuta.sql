------------------------------------------------------------------------------------------
-- CONDUCTORRUTA_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM   sysobjects WHERE  name = N'usp_ConductorRuta_Data' AND type = 'P')
    DROP PROCEDURE usp_ConductorRuta_Data
GO

CREATE PROCEDURE dbo.usp_ConductorRuta_Data 
	@IDPersona int,
	@IDRuta char(20) AS

	SELECT IDPersona, IDRuta, ConductorImporteTramoCompleto, ConductorImporteTramo1, ConductorImporteTramo2, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM ConductorRuta
		WHERE IDPersona = @IDPersona AND IDRuta = @IDRuta

GO