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


------------------------------------------------------------------------------------------
-- RUTADETALLE_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_RutaDetalle_Data' AND type = 'P')
    DROP PROCEDURE usp_RutaDetalle_Data
GO

CREATE PROCEDURE dbo.usp_RutaDetalle_Data
	@IDRuta char(20),
	@IDLugar int AS

	SELECT IDRuta, IDLugar, Indice, IDLugarGrupo, Kilometro, Duracion, Espera, HoraInicio, HoraFin, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM RutaDetalle
		WHERE IDRuta = @IDRuta AND IDLugar = @IDLugar

GO



------------------------------------------------------------------------------------------
-- RUTADETALLE_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_RutaDetalle_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_RutaDetalle_List
GO

CREATE PROCEDURE dbo.sp_RutaDetalle_List
	@IDRuta_FILTER char(20) AS

	SELECT IDRuta, IDLugar, Indice, IDLugarGrupo, Kilometro, Duracion, Espera, HoraInicio, HoraFin, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
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



-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	01/02/2019 09:43
-- Updated:	
-- Description: Obtiene los datos de la Ruta-LugarGrupo
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaLugarGrupo_Get') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaLugarGrupo_Get
GO

CREATE PROCEDURE dbo.usp_RutaLugarGrupo_Get
	@IDRuta char(20),
    @IDLugarGrupo int
AS

BEGIN
	SET NOCOUNT ON;

	SELECT RutaLugarGrupo.IDRuta, RutaLugarGrupo.IDLugarGrupo, RutaLugarGrupo.IDLugarPredeterminado
		FROM RutaLugarGrupo
		WHERE RutaLugarGrupo.IDRuta = @IDRuta AND RutaLugarGrupo.IDLugarGrupo = @IDLugarGrupo 

END
GO



-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	01/02/2019 09:46
-- Updated:	
-- Description: Agrega un Ruta-LugarGrupo
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaLugarGrupo_Add') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaLugarGrupo_Add
GO

CREATE PROCEDURE dbo.usp_RutaLugarGrupo_Add
	@IDRuta char(20),
	@IDLugarGrupo int,
	@IDLugarPredeterminado int
AS

BEGIN
	SET NOCOUNT ON;

    INSERT INTO RutaLugarGrupo
        (IDRuta, IDLugarGrupo, IDLugarPredeterminado)
        VALUES (@IDRuta, @IDLugarGrupo, @IDLugarPredeterminado)
END
GO



-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	01/02/2019 09:48
-- Updated:	
-- Description: Actualiza los datos de la Ruta-LugarGrupo
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaLugarGrupo_Update') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaLugarGrupo_Update
GO

CREATE PROCEDURE dbo.usp_RutaLugarGrupo_Update
	@IDRuta char(20),
	@IDLugarGrupo int,
	@IDLugarPredeterminado int
AS

BEGIN
	SET NOCOUNT ON;

    UPDATE RutaLugarGrupo
        SET IDLugarPredeterminado = @IDLugarPredeterminado
        WHERE IDRuta = @IDRuta AND IDLugarGrupo = @IDLugarGrupo
END
GO



-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	01/02/2019 09:51
-- Updated:	
-- Description: Elimina una Ruta-LugarGrupo
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaLugarGrupo_Delete') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaLugarGrupo_Delete
GO

CREATE PROCEDURE dbo.usp_RutaLugarGrupo_Delete 
	@IDRuta char(20),
	@IDLugarGrupo int
AS
	
BEGIN
	SET NOCOUNT ON;

    DELETE
        FROM RutaLugarGrupo
        WHERE IDRuta = @IDRuta AND IDLugarGrupo = @IDLugarGrupo

END
GO
