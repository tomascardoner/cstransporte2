------------------------------------------------------------------------------------------
-- PERSONA_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Persona_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Persona_Data
GO

CREATE PROCEDURE dbo.sp_Persona_Data 
	@IDPersona_FILTER int AS

	SELECT Persona.IDPersona, Persona.Apellido, Persona.Nombre, Persona.DomicilioCalle1, Persona.DomicilioNumero, Persona.DomicilioPiso, Persona.DomicilioDepartamento, Persona.DomicilioCalle2, Persona.DomicilioCalle3, Persona.CodigoPostal, Persona.IDProvincia, Persona.IDLocalidad, Persona.IDDocumentoTipo, Persona.DocumentoNumero, Persona.IDCondicionIVA, Persona.FechaNacimiento, Persona.Email, Persona.IDTelefono1Tipo, Persona.Telefono1TipoOtro, Persona.Telefono1Area, Persona.Telefono1Numero, Persona.IDTelefono2Tipo, Persona.Telefono2TipoOtro, Persona.Telefono2Area, Persona.Telefono2Numero, Persona.IDTelefono3Tipo, Persona.Telefono3TipoOtro, Persona.Telefono3Area, Persona.Telefono3Numero, Persona.IDTelefono4Tipo, Persona.Telefono4TipoOtro, Persona.Telefono4Area, Persona.Telefono4Numero, Persona.IDTelefono5Tipo, Persona.Telefono5TipoOtro, Persona.Telefono5Area, Persona.Telefono5Numero, Persona.EntidadTipo, Persona.IDPersonaACargo, Persona.IDPersonaCuentaCorriente, Persona.SueldoImporte, Persona.SueldoDia, Persona.PermiteViajarSinPagar, Persona.HabilitadoViajar, Persona.HabilitadoInternet, Persona.Notas, Persona.Activo, Persona.ListaPasajero, Persona.FechaHoraCreacion, Persona.IDUsuarioCreacion, Persona.FechaHoraModificacion, Persona.IDUsuarioModificacion
		FROM Persona
		WHERE Persona.IDPersona = @IDPersona_FILTER

GO



------------------------------------------------------------------------------------------
-- PERSONA_LISTGRID
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_Persona_ListGrid' AND type = 'P')
    DROP PROCEDURE usp_Persona_ListGrid
GO

CREATE PROCEDURE dbo.usp_Persona_ListGrid
	@FirstLetter char(1),
	@EntidadTipo char(2),
	@Activo bit AS

	IF @FirstLetter IS NULL
		BEGIN
		SELECT Persona.IDPersona, Persona.Apellido, Persona.Nombre, dbo.udf_GetEntidadTipoNombre(Persona.EntidadTipo) AS EntidadTipo, dbo.udf_GetDocumentoTipoYNumero(DocumentoTipo.Nombre, Persona.DocumentoNumero) AS Documento, dbo.udf_GetBooleanoSiNo(Persona.Activo) AS Activo, dbo.udf_GetBooleanoSi(ListaPasajero) AS ListaPasajero
			FROM Persona LEFT JOIN DocumentoTipo ON Persona.IDDocumentoTipo = DocumentoTipo.IDDocumentoTipo
			WHERE LEFT(Apellido, 1) NOT IN('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z')
				AND (@EntidadTipo IS NULL OR EntidadTipo = @EntidadTipo)
				AND (@Activo IS NULL OR Persona.Activo = @Activo)
		END
	ELSE
		BEGIN
		SELECT Persona.IDPersona, Persona.Apellido, Persona.Nombre, dbo.udf_GetEntidadTipoNombre(Persona.EntidadTipo) AS EntidadTipo, dbo.udf_GetDocumentoTipoYNumero(DocumentoTipo.Nombre, Persona.DocumentoNumero) AS Documento, dbo.udf_GetBooleanoSiNo(Persona.Activo) AS Activo, dbo.udf_GetBooleanoSi(ListaPasajero) AS ListaPasajero
			FROM Persona WITH (INDEX = IX__Apellido) LEFT JOIN DocumentoTipo ON Persona.IDDocumentoTipo = DocumentoTipo.IDDocumentoTipo
			WHERE (Apellido LIKE @FirstLetter + '%')
				AND (@EntidadTipo IS NULL OR EntidadTipo = @EntidadTipo)
				AND (@Activo IS NULL OR Persona.Activo = @Activo)
		END

GO



------------------------------------------------------------------------------------------
-- PERSONA_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Persona_IDMax' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Persona_IDMax
GO

CREATE PROCEDURE dbo.sp_Persona_IDMax AS
	SELECT Max(Persona.IDPersona) AS IDPersonaMax
		FROM Persona

GO



------------------------------------------------------------------------------------------
-- PERSONA_BYDOCUMENTO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_Persona_ByDocumento' AND type = 'P')
    DROP PROCEDURE usp_Persona_ByDocumento
GO

CREATE PROCEDURE dbo.usp_Persona_ByDocumento
	@IDDocumentoTipo tinyint,
	@DocumentoNumero varchar(15) AS

	SELECT Persona.IDPersona, Persona.Apellido, Persona.Nombre, Persona.DomicilioCalle1, Persona.DomicilioNumero, Persona.DomicilioPiso, Persona.DomicilioDepartamento, Persona.DomicilioCalle2, Persona.DomicilioCalle3, Persona.CodigoPostal, Persona.IDProvincia, Persona.IDLocalidad, Persona.IDDocumentoTipo, Persona.DocumentoNumero, Persona.IDCondicionIVA, Persona.FechaNacimiento, Persona.Email, Persona.IDTelefono1Tipo, Persona.Telefono1TipoOtro, Persona.Telefono1Area, Persona.Telefono1Numero, Persona.IDTelefono2Tipo, Persona.Telefono2TipoOtro, Persona.Telefono2Area, Persona.Telefono2Numero, Persona.IDTelefono3Tipo, Persona.Telefono3TipoOtro, Persona.Telefono3Area, Persona.Telefono3Numero, Persona.IDTelefono4Tipo, Persona.Telefono4TipoOtro, Persona.Telefono4Area, Persona.Telefono4Numero, Persona.IDTelefono5Tipo, Persona.Telefono5TipoOtro, Persona.Telefono5Area, Persona.Telefono5Numero, Persona.EntidadTipo, Persona.IDPersonaACargo, Persona.IDPersonaCuentaCorriente, Persona.SueldoImporte, Persona.SueldoDia, Persona.PermiteViajarSinPagar, Persona.HabilitadoInternet, Persona.Notas, Persona.Activo, Persona.ListaPasajero, Persona.FechaHoraCreacion, Persona.IDUsuarioCreacion, Persona.FechaHoraModificacion, Persona.IDUsuarioModificacion
		FROM Persona
		WHERE Persona.IDDocumentoTipo = @IDDocumentoTipo AND Persona.DocumentoNumero = @DocumentoNumero

GO



------------------------------------------------------------------------------------------
-- PERSONA_BYDOCUMENTO_RESUMEN
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE  name = N'usp_Persona_ByDocumento_Resumen' AND type = 'P')
    DROP PROCEDURE usp_Persona_ByDocumento_Resumen
GO

CREATE PROCEDURE dbo.usp_Persona_ByDocumento_Resumen
	@IDDocumentoTipo tinyint,
	@DocumentoNumero varchar(15) AS

	SELECT Persona.IDPersona, Persona.Apellido, Persona.Nombre, Persona.IDDocumentoTipo, Persona.DocumentoNumero, Persona.Email, Persona.IDTelefono1Tipo, Telefono1Tipo.Nombre AS Telefono1Tipo, Persona.Telefono1TipoOtro, Persona.Telefono1Area, Persona.Telefono1Numero, Persona.IDTelefono2Tipo, Telefono2Tipo.Nombre AS Telefono2Tipo, Persona.Telefono2TipoOtro, Persona.Telefono2Area, Persona.Telefono2Numero, Persona.IDTelefono3Tipo, Telefono3Tipo.Nombre AS Telefono3Tipo, Persona.Telefono3TipoOtro, Persona.Telefono3Area, Persona.Telefono3Numero, Persona.IDTelefono4Tipo, Telefono4Tipo.Nombre AS Telefono4Tipo, Persona.Telefono4TipoOtro, Persona.Telefono4Area, Persona.Telefono4Numero, Persona.IDTelefono5Tipo, Telefono5Tipo.Nombre AS Telefono5Tipo, Persona.Telefono5TipoOtro, Persona.Telefono5Area, Persona.Telefono5Numero, Parametro.NumeroEntero AS IDTelefonoTipoOtro, (SELECT Sum(Importe) AS SaldoActual FROM CuentaCorriente WHERE CuentaCorriente.IDPersona = (CASE ISNULL(Persona.IDPersonaCuentaCorriente, 0) WHEN 0 THEN Persona.IDPersona ELSE Persona.IDPersonaCuentaCorriente END)) AS SaldoActual 
		FROM ((((Persona LEFT JOIN TelefonoTipo AS Telefono1Tipo ON Persona.IDTelefono1Tipo = Telefono1Tipo.IDTelefonoTipo) LEFT JOIN TelefonoTipo AS Telefono2Tipo ON Persona.IDTelefono2Tipo = Telefono2Tipo.IDTelefonoTipo) LEFT JOIN TelefonoTipo AS Telefono3Tipo ON Persona.IDTelefono3Tipo = Telefono3Tipo.IDTelefonoTipo) LEFT JOIN TelefonoTipo AS Telefono4Tipo ON Persona.IDTelefono4Tipo = Telefono4Tipo.IDTelefonoTipo) LEFT JOIN TelefonoTipo AS Telefono5Tipo ON Persona.IDTelefono5Tipo = Telefono5Tipo.IDTelefonoTipo, Parametro
		WHERE Persona.IDDocumentoTipo = @IDDocumentoTipo AND Persona.DocumentoNumero = @DocumentoNumero AND Parametro.IDParametro = 'TelefonoTipo_ID_Otro'

GO



------------------------------------------------------------------------------------------
-- PERSONA_SALDOACTUAL
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Persona_SaldoActual' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Persona_SaldoActual
GO

CREATE PROCEDURE dbo.sp_Persona_SaldoActual 
	@IDPersona_FILTER int AS

	SELECT Sum(Importe) AS SaldoActual
		FROM CuentaCorriente
		WHERE CuentaCorriente.IDPersona = @IDPersona_FILTER

GO



------------------------------------------------------------------------------------------
-- PERSONA_SALDOACTUAL_EXCEPTOVIAJE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Persona_SaldoActual_ExceptoViaje' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Persona_SaldoActual_ExceptoViaje
GO

CREATE PROCEDURE dbo.sp_Persona_SaldoActual_ExceptoViaje 
	@IDPersona_FILTER int,
	@Viaje_FechaHora_FILTER smalldatetime, 
	@Viaje_IDRuta_FILTER char(20),
	@Viaje_Indice_FILTER integer AS

	SELECT Sum(Importe) AS SaldoActual
		FROM CuentaCorriente
		WHERE CuentaCorriente.IDPersona = @IDPersona_FILTER
			AND (NOT (CuentaCorriente.Viaje_FechaHora = @Viaje_FechaHora_FILTER AND CuentaCorriente.Viaje_IDRuta = @Viaje_IDRuta_FILTER AND CuentaCorriente.Viaje_Indice = @Viaje_Indice_FILTER)
			OR (CuentaCorriente.Viaje_FechaHora IS NULL AND CuentaCorriente.Viaje_IDRuta IS NULL AND CuentaCorriente.Viaje_Indice IS NULL))

GO



------------------------------------------------------------------------------------------
-- PERSONAHORARIO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_PersonaHorario_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_PersonaHorario_Data
GO

CREATE PROCEDURE dbo.sp_PersonaHorario_Data 
	@IDPersona_FILTER int,
	@DiaSemana_FILTER tinyint,
	@Hora_FILTER char(8),
	@IDRuta_FILTER char(20) AS

	SELECT PersonaHorario.IDPersona, PersonaHorario.DiaSemana, PersonaHorario.Hora, PersonaHorario.IDRuta, PersonaHorario.FechaDesde, PersonaHorario.FechaHasta, PersonaHorario.IDOrigen, PersonaHorario.Sube, PersonaHorario.IDDestino, PersonaHorario.Baja, PersonaHorario.FechaHoraCreacion, PersonaHorario.IDUsuarioCreacion, PersonaHorario.FechaHoraModificacion, PersonaHorario.IDUsuarioModificacion
		FROM PersonaHorario
		WHERE PersonaHorario.IDPersona = @IDPersona_FILTER AND PersonaHorario.DiaSemana = @DiaSemana_FILTER AND convert(char(8), PersonaHorario.Hora, 108) = @Hora_FILTER AND PersonaHorario.IDRuta = @IDRuta_FILTER

GO



------------------------------------------------------------------------------------------
-- PERSONAHORARIO_BYHORARIO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_PersonaHorario_ByHorario' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_PersonaHorario_ByHorario
GO

CREATE PROCEDURE dbo.sp_PersonaHorario_ByHorario 
	@DiaSemanaBase_FILTER tinyint,
	@Hora_FILTER char(8),
	@IDRuta_FILTER char(20),
	@Fecha_FILTER char(10) AS

	SELECT PersonaHorario.IDPersona
		FROM PersonaHorario
		WHERE PersonaHorario.DiaSemana = @DiaSemanaBase_FILTER
			AND convert(char(8), PersonaHorario.Hora, 108) = @Hora_FILTER
			AND PersonaHorario.IDRuta = @IDRuta_FILTER
			AND (PersonaHorario.FechaDesde IS NULL OR @Fecha_FILTER >= convert(char(10), PersonaHorario.FechaDesde, 111))
			AND (PersonaHorario.FechaHasta IS NULL OR @Fecha_FILTER <= convert(char(10), PersonaHorario.FechaHasta, 111))

GO



------------------------------------------------------------------------------------------
-- PERSONARUTA_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_PersonaRuta_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_PersonaRuta_Data
GO

CREATE PROCEDURE dbo.sp_PersonaRuta_Data 
	@IDPersona_FILTER int,
	@IDRuta_FILTER char(20) AS

	SELECT PersonaRuta.IDPersona, PersonaRuta.IDRuta, PersonaRuta.IDOrigen, PersonaRuta.Sube, PersonaRuta.IDDestino, PersonaRuta.Baja, PersonaRuta.IDListaPrecio, PersonaRuta.FechaHoraCreacion, PersonaRuta.IDUsuarioCreacion, PersonaRuta.FechaHoraModificacion, PersonaRuta.IDUsuarioModificacion
		FROM PersonaRuta
		WHERE PersonaRuta.IDPersona = @IDPersona_FILTER AND PersonaRuta.IDRuta = @IDRuta_FILTER

GO



------------------------------------------------------------------------------------------
-- CONDUCTORRUTA_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ConductorRuta_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ConductorRuta_Data
GO

CREATE PROCEDURE dbo.sp_ConductorRuta_Data 
	@IDPersona_FILTER int,
	@IDRuta_FILTER char(20) AS

	SELECT ConductorRuta.IDPersona, ConductorRuta.IDRuta, ConductorRuta.ConductorImporteTramoCompleto, ConductorRuta.ConductorImporteTramo1, ConductorRuta.ConductorImporteTramo2, ConductorRuta.FechaHoraCreacion, ConductorRuta.IDUsuarioCreacion, ConductorRuta.FechaHoraModificacion, ConductorRuta.IDUsuarioModificacion
		FROM ConductorRuta
		WHERE ConductorRuta.IDPersona = @IDPersona_FILTER AND ConductorRuta.IDRuta = @IDRuta_FILTER

GO



------------------------------------------------------------------------------------------
-- PERSONA_ASISTENCIA_VERIFICAR
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Persona_Inasistencias' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Persona_Inasistencias
GO

CREATE PROCEDURE dbo.sp_Persona_Inasistencias
	@IDPersona_FILTER int,
	@FechaHora_Desde smalldatetime AS

	SELECT COUNT(Indice) AS Inasistencias
		FROM Viaje INNER JOIN ViajeDetalle ON Viaje.FechaHora = ViajeDetalle.FechaHora AND Viaje.IDRuta = ViajeDetalle.IDRuta
		WHERE (Viaje.Estado = 'EP' OR Viaje.Estado = 'FI')
			AND ViajeDetalle.OcupanteTipo = 'PA'
			AND ViajeDetalle.IDPersona = @IDPersona_FILTER
			AND ViajeDetalle.FechaHora >= @FechaHora_Desde
			AND ViajeDetalle.FechaHora < getdate()
			AND ViajeDetalle.Estado = '1CO'
			AND (ViajeDetalle.Realizado IS NULL OR ViajeDetalle.Realizado = 0)

GO



------------------------------------------------------------------------------------------
-- TELEFONOTIPO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM   sysobjects WHERE name = N'usp_TelefonoTipo_Data' AND type = 'P')
    DROP PROCEDURE usp_TelefonoTipo_Data
GO

CREATE PROCEDURE dbo.usp_TelefonoTipo_Data
	@IDTelefonoTipo tinyint AS

	SELECT TelefonoTipo.IDTelefonoTipo, TelefonoTipo.Nombre, TelefonoTipo.DiscadoPrefijo, TelefonoTipo.DiscadoSufijo, TelefonoTipo.Notas, TelefonoTipo.Activo, TelefonoTipo.FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM TelefonoTipo
		WHERE TelefonoTipo.IDTelefonoTipo = @IDTelefonoTipo

GO



------------------------------------------------------------------------------------------
-- TELEFONOTIPO_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE  name = N'usp_TelefonoTipo_IDMax' AND type = 'P')
    DROP PROCEDURE usp_TelefonoTipo_IDMax
GO

CREATE PROCEDURE dbo.usp_TelefonoTipo_IDMax  AS
	DECLARE @IDTelefonoTipoOtro tinyint

	SET @IDTelefonoTipoOtro = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'TelefonoTipo_ID_Otro')

	SELECT Max(TelefonoTipo.IDTelefonoTipo) AS IDTelefonoTipoMax
		FROM TelefonoTipo
		WHERE IDTelefonoTipo <> @IDTelefonoTipoOtro
GO



------------------------------------------------------------------------------------------
-- PERSONA_CALLERID_SEARCH
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Persona_CallerID_Search' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Persona_CallerID_Search
GO

CREATE PROCEDURE dbo.sp_Persona_CallerID_Search
	@TelefonoAreaLocal varchar(5),
	@TelefonoNumero varchar(21) AS

	SELECT Persona.IDPersona, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, Persona.IDTelefono1Tipo, Telefono1Tipo.Nombre AS Telefono1TipoNombre, Persona.Telefono1TipoOtro, Persona.Telefono1Area, Persona.Telefono1Numero, Persona.IDTelefono2Tipo, Telefono2Tipo.Nombre AS Telefono2TipoNombre, Persona.Telefono2TipoOtro, Persona.Telefono2Area, Persona.Telefono2Numero, Persona.IDTelefono3Tipo, Telefono3Tipo.Nombre AS Telefono3TipoNombre, Persona.Telefono3TipoOtro, Persona.Telefono3Area, Persona.Telefono3Numero, Persona.IDTelefono4Tipo, Telefono4Tipo.Nombre AS Telefono4TipoNombre, Persona.Telefono4TipoOtro, Persona.Telefono4Area, Persona.Telefono4Numero, Persona.IDTelefono5Tipo, Telefono5Tipo.Nombre AS Telefono5TipoNombre, Persona.Telefono5TipoOtro, Persona.Telefono5Area, Persona.Telefono5Numero
		FROM ((((Persona LEFT JOIN TelefonoTipo AS Telefono1Tipo ON Persona.IDTelefono1Tipo = Telefono1Tipo.IDTelefonoTipo)  LEFT JOIN TelefonoTipo AS Telefono2Tipo ON Persona.IDTelefono2Tipo = Telefono2Tipo.IDTelefonoTipo) LEFT JOIN TelefonoTipo AS Telefono3Tipo ON Persona.IDTelefono3Tipo = Telefono3Tipo.IDTelefonoTipo) LEFT JOIN TelefonoTipo AS Telefono4Tipo ON Persona.IDTelefono4Tipo = Telefono4Tipo.IDTelefonoTipo) LEFT JOIN TelefonoTipo AS Telefono5Tipo ON Persona.IDTelefono5Tipo = Telefono5Tipo.IDTelefonoTipo
		WHERE (ISNULL(Persona.Telefono1Area, @TelefonoAreaLocal) + Persona.Telefono1Numero = @TelefonoNumero) 
			OR (ISNULL(Persona.Telefono2Area, @TelefonoAreaLocal) + Persona.Telefono2Numero = @TelefonoNumero)
			OR (ISNULL(Persona.Telefono3Area, @TelefonoAreaLocal) + Persona.Telefono3Numero = @TelefonoNumero)
			OR (ISNULL(Persona.Telefono4Area, @TelefonoAreaLocal) + Persona.Telefono4Numero = @TelefonoNumero)
			OR (ISNULL(Persona.Telefono5Area, @TelefonoAreaLocal) + Persona.Telefono5Numero = @TelefonoNumero)
		ORDER BY Persona
GO




------------------------------------------------------------------------------------------
-- DOCUMENTOTIPO_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_DocumentoTipo_IDMax' AND type = 'P')
    DROP PROCEDURE usp_DocumentoTipo_IDMax
GO

CREATE PROCEDURE dbo.usp_DocumentoTipo_IDMax AS
	SELECT Max(DocumentoTipo.IDDocumentoTipo) AS IDDocumentoTipoMax
	FROM DocumentoTipo
	 
GO



------------------------------------------------------------------------------------------
-- DOCUMENTOTIPO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_DocumentoTipo_Data' AND type = 'P')
    DROP PROCEDURE usp_DocumentoTipo_Data
GO

CREATE PROCEDURE dbo.usp_DocumentoTipo_Data
	@IDDocumentoTipo_FILTER tinyint AS

	SELECT IDDocumentoTipo, Nombre, Notas, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM DocumentoTipo
		WHERE DocumentoTipo.IDDocumentoTipo = @IDDocumentoTipo_FILTER

GO



------------------------------------------------------------------------------------------
-- LOCALIDAD_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Localidad_Data'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Localidad_Data
GO

CREATE PROCEDURE dbo.sp_Localidad_Data
	@IDProvincia_FILTER char(1),
	@IDLocalidad_FILTER int AS

	SELECT Localidad.IDProvincia, Localidad.IDLocalidad, Localidad.Nombre, Localidad.CodigoPostal
		FROM Localidad
		WHERE Localidad.IDProvincia = @IDProvincia_FILTER AND Localidad.IDLocalidad = @IDLocalidad_FILTER

GO



------------------------------------------------------------------------------------------
-- PERSONARESPUESTA_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_PersonaRespuesta_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_PersonaRespuesta_Data
GO

CREATE PROCEDURE dbo.sp_PersonaRespuesta_Data
	@IDPersona_FILTER int,
	@FechaHora_FILTER smalldatetime AS

	SELECT PersonaRespuesta.IDPersona, PersonaRespuesta.FechaHora, PersonaRespuesta.Respuesta, PersonaRespuesta.Activo, PersonaRespuesta.FechaHoraCreacion, PersonaRespuesta.IDUsuarioCreacion, PersonaRespuesta.FechaHoraModificacion, PersonaRespuesta.IDUsuarioModificacion
		FROM PersonaRespuesta
		WHERE PersonaRespuesta.IDPersona = @IDPersona_FILTER AND PersonaRespuesta.FechaHora = @FechaHora_FILTER

GO



------------------------------------------------------------------------------------------
-- PERSONARESPUESTA_ACTIVATEALL
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_PersonaRespuesta_ActivateAll'
	   AND 	  type = 'P')
    DROP PROCEDURE sp_PersonaRespuesta_ActivateAll
GO

CREATE PROCEDURE dbo.sp_PersonaRespuesta_ActivateAll
	@IDPersona_FILTER int,
	@Valor bit AS

	UPDATE PersonaRespuesta
		SET Activo = @Valor
		WHERE IDPersona = @IDPersona_FILTER

GO



------------------------------------------------------------------------------------------
-- PERSONA_ADMINISTRATIVOSUELDO_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Persona_AdministrativoSueldo_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Persona_AdministrativoSueldo_List
GO

CREATE PROCEDURE dbo.sp_Persona_AdministrativoSueldo_List AS

	DECLARE @DiaHoy tinyint
	DECLARE @DiaAyer tinyint
	DECLARE @DiaAnteayer tinyint
	DECLARE @DiaUltimoDelMes tinyint

	SET @DiaHoy = (SELECT DATEPART(day, getdate()))
	SET @DiaAyer = (SELECT DATEPART(day, DATEADD(day, -1, getdate())))
	SET @DiaAnteayer = (SELECT DATEPART(day, DATEADD(day, -2, getdate())))
	SET @DiaUltimoDelMes = DATEPART(day, DATEADD(month, 1, getdate() - DATEPART(day, getdate())))

	SELECT Persona.IDPersona, Persona.Apellido, Persona.Nombre, Persona.SueldoImporte, Persona.SueldoDia
		FROM Persona
		WHERE (Persona.EntidadTipo = 'PA' OR Persona.EntidadTipo = 'PO')
			AND Persona.SueldoImporte IS NOT NULL
			AND (Persona.SueldoDia = @DiaHoy OR Persona.SueldoDia = @DiaAyer OR Persona.SueldoDia = @DiaAnteayer OR (Persona.SueldoDia = 99 AND (@DiaHoy = @DiaUltimoDelMes OR @DiaAyer = @DiaUltimoDelMes OR @DiaAnteayer = @DiaUltimoDelMes)))
		ORDER BY Persona.IDPersona

GO



------------------------------------------------------------------------------------------
-- PERSONA_ADMINISTRATIVOSUELDO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Persona_AdministrativoSueldo_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Persona_AdministrativoSueldo_Data
GO

CREATE PROCEDURE dbo.sp_Persona_AdministrativoSueldo_Data
	@IDPersona int,
	@MovimientoCreado bit OUTPUT AS

	DECLARE @CuentaCorrienteGrupo_ID_Sueldo int

	SET @CuentaCorrienteGrupo_ID_Sueldo = (SELECT NumeroEntero FROM Parametro WHERE IDParametro = 'CuentaCorrienteGrupo_ID_Sueldo')

	IF (SELECT COUNT(IDMovimiento) FROM CuentaCorriente WHERE CuentaCorriente.IDPersona = @IDPersona AND CuentaCorriente.IDCuentaCorrienteGrupo = @CuentaCorrienteGrupo_ID_Sueldo AND datediff(day, CuentaCorriente.FechaHora, getdate()) >= 0 AND datediff(day, CuentaCorriente.FechaHora, getdate()) <= 2) = 0
		SET @MovimientoCreado = 0
	ELSE
		SET @MovimientoCreado = 1

GO



------------------------------------------------------------------------------------------
-- REGISTROLLAMADA_UPDATE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_RegistroLlamada_Update' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_RegistroLlamada_Update
GO

CREATE PROCEDURE dbo.sp_RegistroLlamada_Update
	@FechaHora smalldatetime,
	@TelefonoNumero varchar(21) AS

	INSERT INTO RegistroLlamada
		(FechaHora, TelefonoNumero)
		VALUES (@FechaHora, @TelefonoNumero)

GO



------------------------------------------------------------------------------------------
-- PERSONA_DESACTIVARINACTIVOS
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Persona_DesactivarInactivos' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Persona_DesactivarInactivos
GO

CREATE PROCEDURE dbo.sp_Persona_DesactivarInactivos
	@Meses tinyint AS

	DECLARE @Persona table (IDPersona int NOT NULL)
	
	--PERSONAS EN VIAJEDETALLE
	INSERT INTO @Persona
		SELECT DISTINCT IDPersona
			FROM ViajeDetalle
			WHERE FechaHora >= DATEADD(month, -@Meses, getdate())

	--PERSONAS CTA. CTE. EN VIAJEDETALLE
	INSERT INTO @Persona
		SELECT DISTINCT IDPersonaCuentaCorriente
			FROM ViajeDetalle
			WHERE IDPersonaCuentaCorriente IS NOT NULL AND FechaHora >= DATEADD(month, -@Meses, getdate())

	--PERSONAS EN VIAJE (VIAJE ESPECIAL)
	INSERT INTO @Persona
		SELECT DISTINCT IDPersona
			FROM Viaje
			WHERE IDPersona IS NOT NULL AND FechaHora >= DATEADD(month, -@Meses, getdate())

	--CONDUCTORES EN VIAJE
	INSERT INTO @Persona
		SELECT DISTINCT IDConductor
			FROM Viaje
			WHERE IDConductor IS NOT NULL AND FechaHora >= DATEADD(month, -@Meses, getdate())

	--CONDUCTORES 2 EN VIAJE
	INSERT INTO @Persona
		SELECT DISTINCT IDConductor2
			FROM Viaje
			WHERE IDConductor2 IS NOT NULL AND FechaHora >= DATEADD(month, -@Meses, getdate())

	--PERSONAS EN CTA. CTE.
	INSERT INTO @Persona
		SELECT DISTINCT IDPersona
			FROM CuentaCorriente
			WHERE IDPersona IS NOT NULL AND FechaHora >= DATEADD(month, -@Meses, getdate())

	--PERSONAS EN CTA. CTE. ORIGEN
	INSERT INTO @Persona
		SELECT DISTINCT IDPersonaOrigen
			FROM CuentaCorriente
			WHERE IDPersonaOrigen IS NOT NULL AND FechaHora >= DATEADD(month, -@Meses, getdate())

	--SON LOS QUE NO ESTAN EN LA LISTA, POR LO TANTO LOS DESACTIVO
	UPDATE Persona
		SET Activo = 0
		WHERE IDPersona NOT IN (SELECT IDPersona FROM @Persona)

GO



------------------------------------------------------------------------------------------
-- PERSONA BUSCAR POR DOCUMENTO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'uspPersonaBuscarPorDocumento' AND type = 'P')
    DROP PROCEDURE uspPersonaBuscarPorDocumento
GO

CREATE PROCEDURE dbo.uspPersonaBuscarPorDocumento 
	@IDDocumentoTipo tinyint,
	@DocumentoNumero varchar(15) AS

	SELECT IDPersona, Apellido, Nombre, Activo
		FROM Persona
		WHERE IDDocumentoTipo = @IDDocumentoTipo AND DocumentoNumero = @DocumentoNumero
GO



------------------------------------------------------------------------------------------
-- PERSONA BUSCAR POR APELLIDO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'uspPersonaBuscarPorApellido' AND type = 'P')
    DROP PROCEDURE uspPersonaBuscarPorApellido
GO

CREATE PROCEDURE dbo.uspPersonaBuscarPorApellido 
	@Apellido varchar(100),
	@Nombre varchar(100) AS

	SELECT TOP 1 IDPersona, Apellido, Nombre, Activo
		FROM Persona
		WHERE Apellido LIKE @Apellido + '%' AND Nombre LIKE @Nombre + '%'
		ORDER BY Apellido, Nombre
GO