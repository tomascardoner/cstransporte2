------------------------------------------------------------------------------------------
-- VEHICULOMANTENIMIENTO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_VehiculoMantenimiento_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_VehiculoMantenimiento_Data
GO

CREATE PROCEDURE dbo.sp_VehiculoMantenimiento_Data 
	@IDVehiculo_FILTER int,
	@IDVehiculoMantenimientoGrupo_FILTER int AS

	SELECT VehiculoMantenimiento.IDVehiculo, VehiculoMantenimiento.IDVehiculoMantenimientoGrupo, VehiculoMantenimiento.Tipo, VehiculoMantenimiento.KilometrajeLapso, VehiculoMantenimiento.KilometrajePreaviso, VehiculoMantenimiento.DiasLapso, VehiculoMantenimiento.DiasPreaviso, VehiculoMantenimiento.FechaFecha, VehiculoMantenimiento.FechaPreaviso, VehiculoMantenimiento.Notas, VehiculoMantenimiento.Activo, VehiculoMantenimiento.FechaHoraCreacion, VehiculoMantenimiento.IDUsuarioCreacion, VehiculoMantenimiento.FechaHoraModificacion, VehiculoMantenimiento.IDUsuarioModificacion
		FROM VehiculoMantenimiento
		WHERE VehiculoMantenimiento.IDVehiculo = @IDVehiculo_FILTER AND VehiculoMantenimiento.IDVehiculoMantenimientoGrupo = @IDVehiculoMantenimientoGrupo_FILTER

GO



------------------------------------------------------------------------------------------
-- VEHICULOMANTENIMIENTOACCION_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_VehiculoMantenimientoAccion_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_VehiculoMantenimientoAccion_Data
GO

CREATE PROCEDURE dbo.sp_VehiculoMantenimientoAccion_Data 
	@IDVehiculoMantenimientoAccion_FILTER int AS

	SELECT VehiculoMantenimientoAccion.IDVehiculoMantenimientoAccion, VehiculoMantenimientoAccion.IDVehiculo, VehiculoMantenimientoAccion.IDVehiculoMantenimientoGrupo, VehiculoMantenimientoAccion.IDConductor, VehiculoMantenimientoAccion.FechaHora, VehiculoMantenimientoAccion.Kilometraje, VehiculoMantenimientoAccion.Litros, VehiculoMantenimientoAccion.Importe, VehiculoMantenimientoAccion.Notas, VehiculoMantenimientoAccion.FechaHoraCreacion, VehiculoMantenimientoAccion.IDUsuarioCreacion, VehiculoMantenimientoAccion.FechaHoraModificacion, VehiculoMantenimientoAccion.IDUsuarioModificacion
		FROM VehiculoMantenimientoAccion
		WHERE VehiculoMantenimientoAccion.IDVehiculoMantenimientoAccion = @IDVehiculoMantenimientoAccion_FILTER

GO



------------------------------------------------------------------------------------------
-- VEHICULOMANTENIMIENTOACCION_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_VehiculoMantenimientoAccion_IDMax' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_VehiculoMantenimientoAccion_IDMax
GO

CREATE PROCEDURE dbo.sp_VehiculoMantenimientoAccion_IDMax AS
	SELECT Max(VehiculoMantenimientoAccion.IDVehiculoMantenimientoAccion) AS IDVehiculoMantenimientoAccionMax
	FROM VehiculoMantenimientoAccion
	 
GO



------------------------------------------------------------------------------------------
-- VEHICULOMANTENIMIENTOGRUPO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_VehiculoMantenimientoGrupo_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_VehiculoMantenimientoGrupo_Data
GO

CREATE PROCEDURE dbo.sp_VehiculoMantenimientoGrupo_Data 
	@IDVehiculoMantenimientoGrupo_FILTER int AS

	SELECT VehiculoMantenimientoGrupo.IDVehiculoMantenimientoGrupo, VehiculoMantenimientoGrupo.Nombre, VehiculoMantenimientoGrupo.Activo, VehiculoMantenimientoGrupo.Notas, VehiculoMantenimientoGrupo.FechaHoraCreacion, VehiculoMantenimientoGrupo.IDUsuarioCreacion, VehiculoMantenimientoGrupo.FechaHoraModificacion, VehiculoMantenimientoGrupo.IDUsuarioModificacion
		FROM VehiculoMantenimientoGrupo
		WHERE VehiculoMantenimientoGrupo.IDVehiculoMantenimientoGrupo = @IDVehiculoMantenimientoGrupo_FILTER

GO



------------------------------------------------------------------------------------------
-- VEHICULOMANTENIMIENTOGRUPO_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_VehiculoMantenimientoGrupo_IDMax' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_VehiculoMantenimientoGrupo_IDMax
GO

CREATE PROCEDURE dbo.sp_VehiculoMantenimientoGrupo_IDMax AS
	SELECT Max(VehiculoMantenimientoGrupo.IDVehiculoMantenimientoGrupo) AS IDVehiculoMantenimientoGrupoMax
	FROM VehiculoMantenimientoGrupo
	 
GO



------------------------------------------------------------------------------------------
-- VEHICULOMANTENIMIENTO_COPY
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_VehiculoMantenimiento_Copy' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_VehiculoMantenimiento_Copy
GO

CREATE PROCEDURE dbo.sp_VehiculoMantenimiento_Copy
	@IDVehiculoOrigen int,
	@IDVehiculoDestino int,
	@IDUsuario smallint AS

	INSERT INTO VehiculoMantenimiento
		(IDVehiculo, IDVehiculoMantenimientoGrupo, Tipo, KilometrajeLapso, KilometrajePreaviso, DiasLapso, DiasPreaviso, FechaFecha, FechaPreaviso, Notas, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)
		SELECT @IDVehiculoDestino, IDVehiculoMantenimientoGrupo, Tipo, KilometrajeLapso, KilometrajePreaviso, DiasLapso, DiasPreaviso, FechaFecha, FechaPreaviso, Notas, Activo, getdate(), @IDUsuario, getdate(), @IDUsuario
			FROM VehiculoMantenimiento
			WHERE VehiculoMantenimiento.IDVehiculo = @IDVehiculoOrigen
				AND IDVehiculoMantenimientoGrupo NOT IN(SELECT IDVehiculoMantenimientoGrupo FROM VehiculoMantenimiento WHERE IDVehiculo = @IDVehiculoDestino)

GO