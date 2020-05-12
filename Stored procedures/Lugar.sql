------------------------------------------------------------------------------------------
-- LUGAR_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'sp_Lugar_Data' AND type = 'P')
    DROP PROCEDURE sp_Lugar_Data
GO

CREATE PROCEDURE dbo.sp_Lugar_Data
	@IDLugar_FILTER int AS

	SELECT IDLugar, Nombre, NombreCorto, UbicacionLatitud, UbicacionLongitud, Activo, Notas, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM Lugar
		WHERE IDLugar = @IDLugar_FILTER

GO



------------------------------------------------------------------------------------------
-- LUGAR_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'sp_Lugar_IDMax' AND type = 'P')
    DROP PROCEDURE sp_Lugar_IDMax
GO

CREATE PROCEDURE dbo.sp_Lugar_IDMax AS
	SELECT Max(Lugar.IDLugar) AS IDLugarMax
	FROM Lugar
	 
GO



------------------------------------------------------------------------------------------
-- LUGARGRUPO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_LugarGrupo_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_LugarGrupo_Data
GO

CREATE PROCEDURE dbo.sp_LugarGrupo_Data 
	@IDLugarGrupo_FILTER int AS

	SELECT LugarGrupo.IDLugarGrupo, LugarGrupo.Nombre, LugarGrupo.Activo, LugarGrupo.Notas, LugarGrupo.FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM LugarGrupo
		WHERE LugarGrupo.IDLugarGrupo = @IDLugarGrupo_FILTER

GO



------------------------------------------------------------------------------------------
-- LUGARGRUPO_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_LugarGrupo_IDMax' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_LugarGrupo_IDMax
GO

CREATE PROCEDURE dbo.sp_LugarGrupo_IDMax AS
	SELECT Max(LugarGrupo.IDLugarGrupo) AS IDLugarGrupoMax
	FROM LugarGrupo
	 
GO