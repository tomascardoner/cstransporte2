------------------------------------------------------------------------------------------
-- USUARIO_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM   sysobjects WHERE  name = N'usp_Usuario_List' AND type = 'P')
    DROP PROCEDURE usp_Usuario_List
GO

CREATE PROCEDURE dbo.usp_Usuario_List 
	AS

	SELECT IDUsuario, LoginName, Nombre, Password, Descripcion, IDUsuarioGrupo, IDEmpresa, IDPersona, Notas, Activo, Semaforo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM Usuario
		WHERE IDUsuario > 1 AND Activo = 1
		ORDER BY Nombre

GO



------------------------------------------------------------------------------------------
-- USUARIO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM   sysobjects WHERE  name = N'usp_Usuario_Data' AND type = 'P')
    DROP PROCEDURE usp_Usuario_Data
GO

CREATE PROCEDURE dbo.usp_Usuario_Data 
	@IDUsuario smallint AS

	SELECT IDUsuario, LoginName, Nombre, Password, Descripcion, IDUsuarioGrupo, IDEmpresa, IDPersona, Notas, Activo, Semaforo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM Usuario
		WHERE IDUsuario = @IDUsuario

GO



------------------------------------------------------------------------------------------
-- USUARIO_DATA_BYLOGINNAME
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM   sysobjects WHERE  name = N'usp_Usuario_Data_ByLoginName' AND type = 'P')
    DROP PROCEDURE usp_Usuario_Data_ByLoginName
GO

CREATE PROCEDURE dbo.usp_Usuario_Data_ByLoginName 
	@LoginName varchar(30) AS

	SELECT IDUsuario, LoginName, Nombre, Password, Descripcion, IDUsuarioGrupo, IDEmpresa, IDPersona, Notas, Activo, Semaforo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM Usuario
		WHERE LoginName = @LoginName

GO



------------------------------------------------------------------------------------------
-- USUARIO_DATA_BYPERSONA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM   sysobjects WHERE  name = N'usp_Usuario_Data_ByPersona' AND type = 'P')
    DROP PROCEDURE usp_Usuario_Data_ByPersona
GO

CREATE PROCEDURE dbo.usp_Usuario_Data_ByPersona 
	@IDPersona int AS

	SELECT IDUsuario, LoginName, Nombre, Password, Descripcion, IDUsuarioGrupo, IDEmpresa, IDPersona, Notas, Activo, Semaforo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM Usuario
		WHERE IDPersona = @IDPersona

GO



------------------------------------------------------------------------------------------
-- USUARIO_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE  name = N'usp_Usuario_IDMax' AND type = 'P')
    DROP PROCEDURE usp_Usuario_IDMax
GO

CREATE PROCEDURE dbo.usp_Usuario_IDMax AS
	SELECT Max(Usuario.IDUsuario) AS IDUsuarioMax
	FROM Usuario

GO



------------------------------------------------------------------------------------------
-- USUARIO_SEMAFORO_UPDATE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM   sysobjects WHERE  name = N'usp_Usuario_Semaforo_Update' AND type = 'P')
    DROP PROCEDURE usp_Usuario_Semaforo_Update
GO

CREATE PROCEDURE dbo.usp_Usuario_Semaforo_Update 
	@IDUsuario smallint AS

	UPDATE Usuario
		SET Semaforo = (ABS(CHECKSUM(NewId())) % 100000000)
		WHERE IDUsuario = @IDUsuario

GO



------------------------------------------------------------------------------------------
-- USUARIOGRUPO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'usp_UsuarioGrupo_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE usp_UsuarioGrupo_Data
GO

CREATE PROCEDURE dbo.usp_UsuarioGrupo_Data
	@IDUsuarioGrupo tinyint AS

	SELECT UsuarioGrupo.IDUsuarioGrupo, UsuarioGrupo.Nombre, UsuarioGrupo.Notas, UsuarioGrupo.Activo, UsuarioGrupo.FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM UsuarioGrupo
		WHERE UsuarioGrupo.IDUsuarioGrupo = @IDUsuarioGrupo

GO



------------------------------------------------------------------------------------------
-- USUARIOGRUPO_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'usp_UsuarioGrupo_IDMax' 
	   AND 	  type = 'P')
    DROP PROCEDURE usp_UsuarioGrupo_IDMax
GO

CREATE PROCEDURE dbo.usp_UsuarioGrupo_IDMax  AS
	SELECT Max(UsuarioGrupo.IDUsuarioGrupo) AS IDUsuarioGrupoMax
		FROM UsuarioGrupo

GO


------------------------------------------------------------------------------------------
-- USUARIOGRUPOPERMISO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_UsuarioGrupoPermiso_Data' AND type = 'P')
    DROP PROCEDURE usp_UsuarioGrupoPermiso_Data
GO

CREATE PROCEDURE dbo.usp_UsuarioGrupoPermiso_Data
	@IDUsuarioGrupo tinyint,
	@IDPermiso char(100) AS

	SELECT IDUsuarioGrupo, IDPermiso, FechaHoraCreacion, IDUsuarioCreacion
		FROM UsuarioGrupoPermiso
		WHERE IDUsuarioGrupo = @IDUsuarioGrupo AND IDPermiso = @IDPermiso

GO



------------------------------------------------------------------------------------------
-- USUARIOGRUPOPERMISO_ALLDATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_UsuarioGrupoPermiso_AllData' AND type = 'P')
    DROP PROCEDURE usp_UsuarioGrupoPermiso_AllData
GO

CREATE PROCEDURE dbo.usp_UsuarioGrupoPermiso_AllData
	@IDUsuarioGrupo int AS

	SELECT IDUsuarioGrupo, IDPermiso, FechaHoraCreacion, IDUsuarioCreacion
		FROM UsuarioGrupoPermiso
		WHERE IDUsuarioGrupo = @IDUsuarioGrupo

GO