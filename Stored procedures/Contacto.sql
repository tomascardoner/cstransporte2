------------------------------------------------------------------------------------------
-- CONTACTO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_Contacto_Data' AND type = 'P')
    DROP PROCEDURE usp_Contacto_Data
GO

CREATE PROCEDURE dbo.usp_Contacto_Data
	@IDContacto int AS

	SELECT IDContacto, Titulo, Apellido, Nombre, Compania, TituloLaboral, IDTelefono1Tipo, Telefono1TipoOtro, Telefono1Area, Telefono1Numero, IDTelefono2Tipo, Telefono2TipoOtro, Telefono2Area, Telefono2Numero, IDTelefono3Tipo, Telefono3TipoOtro, Telefono3Area, Telefono3Numero, IDTelefono4Tipo, Telefono4TipoOtro, Telefono4Area, Telefono4Numero, IDTelefono5Tipo, Telefono5TipoOtro, Telefono5Area, Telefono5Numero, IDContactoGrupo, DomicilioLaboralCalle1, DomicilioLaboralNumero, DomicilioLaboralPiso, DomicilioLaboralDepartamento, DomicilioLaboralCalle2, DomicilioLaboralCalle3, DomicilioLaboralCodigoPostal, DomicilioLaboralIDProvincia, DomicilioLaboralIDLocalidad, DomicilioLaboralMailing, DomicilioParticularCalle1, DomicilioParticularNumero, DomicilioParticularPiso, DomicilioParticularDepartamento, DomicilioParticularCalle2, DomicilioParticularCalle3, DomicilioParticularCodigoPostal, DomicilioParticularIDProvincia, DomicilioParticularIDLocalidad, DomicilioParticularMailing, DomicilioOtroNombre, DomicilioOtroCalle1, DomicilioOtroNumero, DomicilioOtroPiso, DomicilioOtroDepartamento, DomicilioOtroCalle2, DomicilioOtroCalle3, DomicilioOtroCodigoPostal, DomicilioOtroIDProvincia, DomicilioOtroIDLocalidad, DomicilioOtroMailing, Email1, Email1Nombre, Email2, Email2Nombre, Email3, Email3Nombre, PaginaWeb, SobreNombre, FechaNacimiento, Asistente, Notas, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM Contacto
		WHERE IDContacto = @IDContacto

GO



------------------------------------------------------------------------------------------
-- CONTACTO_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_Contacto_IDMax' AND type = 'P')
    DROP PROCEDURE usp_Contacto_IDMax
GO

CREATE PROCEDURE dbo.usp_Contacto_IDMax AS

	SELECT Max(IDContacto) AS IDContactoMax
		FROM Contacto

GO



------------------------------------------------------------------------------------------
-- CONTACTOGRUPO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_ContactoGrupo_Data' AND type = 'P')
    DROP PROCEDURE usp_ContactoGrupo_Data
GO

CREATE PROCEDURE dbo.usp_ContactoGrupo_Data
	@IDContactoGrupo int  AS

	SELECT IDContactoGrupo, Nombre, Notas, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM ContactoGrupo
		WHERE IDContactoGrupo = @IDContactoGrupo

GO



------------------------------------------------------------------------------------------
-- CONTACTOGRUPO_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_ContactoGrupo_IDMax' AND type = 'P')
    DROP PROCEDURE usp_ContactoGrupo_IDMax
GO

CREATE PROCEDURE dbo.usp_ContactoGrupo_IDMax AS

	SELECT Max(IDContactoGrupo) AS IDContactoGrupoMax 
		FROM ContactoGrupo

GO



------------------------------------------------------------------------------------------
-- TITULO_CHECK
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_Titulo_Check' AND type = 'P')
    DROP PROCEDURE usp_Titulo_Check
GO

CREATE PROCEDURE dbo.usp_Titulo_Check 
	@Nombre varchar(50) AS

	DECLARE @IDTitulo int

	IF (SELECT COUNT(IDTitulo) FROM Titulo WHERE Nombre = @Nombre) = 0
		BEGIN
			SET @IDTitulo = ISNULL((SELECT Max(Titulo.IDTitulo) FROM Titulo), 0) + 1
			INSERT INTO Titulo VALUES (@IDTitulo, @Nombre)
		END

GO
