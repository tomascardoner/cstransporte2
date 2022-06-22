------------------------------------------------------------------------------------------
-- RUTADETALLE_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_RutaDetalle_Data' AND type = 'P')
    DROP PROCEDURE usp_RutaDetalle_Data
GO

CREATE PROCEDURE dbo.usp_RutaDetalle_Data
	@IDRuta char(20),
	@IDLugar int AS

	SELECT IDRuta, IDLugar, Indice, IDLugarGrupo, Kilometro, Duracion, Espera, HoraInicio, HoraFin, DistanciaNotificacion, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM RutaDetalle
		WHERE IDRuta = @IDRuta AND IDLugar = @IDLugar

GO



------------------------------------------------------------------------------------------
-- RUTADETALLE_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'usp_RutaDetalle_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE usp_RutaDetalle_List
GO

CREATE PROCEDURE dbo.usp_RutaDetalle_List
	@IDRuta_FILTER char(20) AS

	SELECT IDRuta, IDLugar, Indice, IDLugarGrupo, Kilometro, Duracion, Espera, HoraInicio, HoraFin, DistanciaNotificacion, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM RutaDetalle
		WHERE IDRuta = @IDRuta_FILTER
		ORDER BY Indice

GO



------------------------------------------------------------------------------------------
-- RUTADETALLE_INDICEMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'usp_RutaDetalle_IndiceMax' 
	   AND 	  type = 'P')
    DROP PROCEDURE usp_RutaDetalle_IndiceMax
GO

CREATE PROCEDURE dbo.usp_RutaDetalle_IndiceMax
	@IDRuta char(20) AS

	SELECT Max(Indice) AS IndiceMax
		FROM RutaDetalle
		WHERE IDRuta = @IDRuta

GO
