------------------------------------------------------------------------------------------
-- RUTA_ALLDATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_Ruta_AllData' AND type = 'P')
    DROP PROCEDURE usp_Ruta_AllData
GO

CREATE PROCEDURE dbo.usp_Ruta_AllData AS

	SELECT IDRuta, Nombre, IDOrigen, IDDestino, IDRutaGrupo, Kilometro, Duracion, LimiteCancelacionIDLugar, LimiteCancelacionDuracion, Permite2Conductores, ConductorImporteTramoCompleto, ConductorImporteTramo1, ConductorImporteTramo2, Notas, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM Ruta
		ORDER BY Nombre

GO



------------------------------------------------------------------------------------------
-- RUTA_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'usp_Ruta_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE usp_Ruta_Data
GO

CREATE PROCEDURE dbo.usp_Ruta_Data
	@IDRuta char(20) AS

	SELECT IDRuta, Nombre, IDOrigen, IDDestino, IDRutaGrupo, Kilometro, Duracion, LimiteCancelacionIDLugar, LimiteCancelacionDuracion, Permite2Conductores, ConductorImporteTramoCompleto, ConductorImporteTramo1, ConductorImporteTramo2, Notas, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM Ruta
		WHERE IDRuta = @IDRuta

GO


------------------------------------------------------------------------------------------
-- RUTA_STATISTICS
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Ruta_Statistics' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Ruta_Statistics
GO

CREATE PROCEDURE dbo.sp_Ruta_Statistics
	@IDRuta_FILTER char(20) AS

	SELECT Count(IDRuta) AS CantidadLugares, Min(Indice) AS IndiceMinimo, Max(Indice) AS IndiceMaximo
		FROM RutaDetalle
		WHERE IDRuta = @IDRuta_FILTER

GO