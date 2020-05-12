------------------------------------------------------------------------------------------
-- DATE_ISLEAPYEAR
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'udf_Date_IsLeapYear'
	   AND 	  type = 'FN')
    DROP FUNCTION udf_Date_IsLeapYear
GO


CREATE FUNCTION dbo.udf_Date_IsLeapYear 
	(@Year smallint)
	RETURNS bit AS

	BEGIN

		DECLARE @Result bit

		IF @Year % 4 = 0
			BEGIN
			IF @Year % 100 = 0
				BEGIN
				IF @Year % 400 = 0
					BEGIN
					SET @Result = 1
					END
				ELSE
					BEGIN
					SET @Result = 0
					END
				END
			ELSE
				BEGIN
				SET @Result = 1
				END
			END
		ELSE
			BEGIN
			SET @Result = 0
			END

		RETURN @Result
	END

GO



------------------------------------------------------------------------------------------
-- DATE_GETLASTDAYOFMONTH
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'udf_Date_GetLastDayOfMonth'
	   AND 	  type = 'FN')
    DROP FUNCTION udf_Date_GetLastDayOfMonth
GO


CREATE FUNCTION dbo.udf_Date_GetLastDayOfMonth 
	(@Month tinyint, @Year smallint)
	RETURNS tinyint AS

	BEGIN
		DECLARE @Result tinyint

		IF @Month IN(1, 3, 5, 7, 8, 10, 12)
			-- Enero, Marzo, Mayo, Julio, Agosto, Octubre, Diciembre
			BEGIN
			SET @Result = 31
			END
		ELSE
			BEGIN
			IF @Month IN(4, 6, 9, 11)
				-- Abril, Junio, Septiembre, Noviembre
				BEGIN
				SET @Result = 30
				END
			ELSE
				-- Febrero (ver si es bisiesto)
				BEGIN
				IF dbo.udf_Date_IsLeapYear(@Year) = 1
					BEGIN
					SET @Result = 29
					END
				ELSE
					BEGIN
					SET @Result = 28
					END
				END
			END

		RETURN @Result
	END

GO



------------------------------------------------------------------------------------------
-- DOMICILIO_GETFORDISPLAY
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'udf_Domicilio_GetShort'
	   AND 	  type = 'FN')
    DROP FUNCTION udf_Domicilio_GetShort
GO


CREATE FUNCTION dbo.udf_Domicilio_GetShort 
	(@DomicilioCalle1 varchar(50), @DomicilioNumero varchar(10), @DomicilioPiso varchar(10), @DomicilioDepartamento varchar(10), @DomicilioCalle2 varchar(50), @DomicilioCalle3 varchar(50))
	RETURNS varchar(4000) AS

	BEGIN
		DECLARE @Result varchar(400)

		IF @DomicilioCalle1 IS NOT NULL
			BEGIN
			SET @Result = @DomicilioCalle1
			IF @DomicilioNumero IS NOT NULL
				BEGIN
				SET @Result = @Result + ' ' + @DomicilioNumero
				END
			IF @DomicilioPiso IS NOT NULL
				BEGIN
				IF ISNUMERIC(@DomicilioPiso) = 1
					BEGIN
					SET @Result = @Result + ' P.' + @DomicilioPiso
					END
				ELSE
					BEGIN
					SET @Result = @Result + ' ' + @DomicilioPiso
					END
				END
			END
			IF @DomicilioDepartamento IS NOT NULL
				BEGIN
				SET @Result = @Result + ' "' + @DomicilioDepartamento + '"'
				END
			IF @DomicilioCalle2 IS NOT NULL
				BEGIN
				IF @DomicilioCalle3 IS NOT NULL
					BEGIN
					SET @Result = @Result + ' e/' + @DomicilioCalle2 + ' y ' + @DomicilioCalle3
					END
				ELSE
					BEGIN
					SET @Result = @Result + ' y ' + @DomicilioCalle2
					END
				END
		RETURN @Result
	END

GO




------------------------------------------------------------------------------------------
-- GETVIAJETRAMONUMERO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'udf_GetViajeTramoNumero'
	   AND 	  type = 'FN')
    DROP FUNCTION udf_GetViajeTramoNumero
GO

CREATE FUNCTION dbo.udf_GetViajeTramoNumero
	(@IDConductor int,
	@ViajeIDConductor int,
	@ViajeIDConductor2 int)
	RETURNS tinyint AS
	
	BEGIN
		DECLARE @Result tinyint
		
		IF ISNULL(@ViajeIDConductor2, 0) = 0
			BEGIN
			SET @Result = 0
			END
		ELSE
			BEGIN
			IF @IDConductor = @ViajeIDConductor
				BEGIN
				SET @Result = 1
				END
			ELSE IF @IDConductor = @ViajeIDConductor2
				BEGIN
				SET @Result = 2
				END
			END
			
		RETURN @Result
	END
GO




------------------------------------------------------------------------------------------
-- GETVIAJETRAMONOMBRE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'udf_GetViajeTramoNombre'
	   AND 	  type = 'FN')
    DROP FUNCTION udf_GetViajeTramoNombre
GO

CREATE FUNCTION dbo.udf_GetViajeTramoNombre
	(@IDConductor int,
	@ViajeIDConductor int,
	@ViajeIDConductor2 int)
	RETURNS varchar(8) AS
	
	BEGIN
		DECLARE @Result varchar(8)
		
		IF ISNULL(@ViajeIDConductor2, 0) = 0
			BEGIN
			SET @Result = 'Completo'
			END
		ELSE
			BEGIN
			IF @IDConductor = @ViajeIDConductor
				BEGIN
				SET @Result = 'Tramo 1'
				END
			ELSE IF @IDConductor = @ViajeIDConductor2
				BEGIN
				SET @Result = 'Tramo 2'
				END
			END
			
		RETURN @Result
	END
GO



------------------------------------------------------------------------------------------
-- GETVIAJETRAMOIMPORTE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'udf_GetViajeTramoImporte'
	   AND 	  type = 'FN')
    DROP FUNCTION udf_GetViajeTramoImporte
GO

CREATE FUNCTION dbo.udf_GetViajeTramoImporte
	(@ConductorRutaImporte smallmoney,
	@HorarioImporte smallmoney,
	@RutaImporte smallmoney)
	RETURNS smallmoney AS
	
	BEGIN
		DECLARE @Result smallmoney
		
		IF ISNULL(@ConductorRutaImporte, 0) <> 0
			BEGIN
			SET @Result = @ConductorRutaImporte
			END
		ELSE
			BEGIN
			IF ISNULL(@HorarioImporte, 0) <> 0
				BEGIN
				SET @Result = @HorarioImporte
				END
			ELSE
				BEGIN
				IF ISNULL(@RutaImporte, 0) <> 0
					BEGIN
					SET @Result = @RutaImporte
					END
				ELSE
					BEGIN
					SET @Result = 0
					END
				END
			END


		RETURN @Result
	END
GO



------------------------------------------------------------------------------------------
-- GETENTIDADTIPONOMBRE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'udf_GetEntidadTipoNombre' AND type = 'FN')
    DROP FUNCTION udf_GetEntidadTipoNombre
GO

CREATE FUNCTION dbo.udf_GetEntidadTipoNombre
	(@EntidadTipo char(2))
	RETURNS varchar(14) AS
	
	BEGIN
		DECLARE @Result varchar(14)
		
		IF @EntidadTipo = 'PC'
			SET @Result = 'Cliente'
		IF @EntidadTipo = 'PO'
			SET @Result = 'Conductor'
		IF @EntidadTipo = 'PA'
			SET @Result = 'Administrativo'
			
		RETURN @Result
	END
GO



------------------------------------------------------------------------------------------
-- GETDOCUMENTOTIPOYNUMERO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'udf_GetDocumentoTipoYNumero' AND type = 'FN')
    DROP FUNCTION udf_GetDocumentoTipoYNumero
GO

CREATE FUNCTION dbo.udf_GetDocumentoTipoYNumero
	(@DocumentoTipoNombre varchar(10),
	 @DocumentoNumero varchar(15))
	RETURNS varchar(25) AS
	
	BEGIN
		DECLARE @Result varchar(25)
		
		IF ISNULL(@DocumentoNumero, '') = ''
			SET @Result = ''
		ELSE
			IF ISNULL(@DocumentoTipoNombre, '') = ''
				SET @Result = @DocumentoNumero
			ELSE
				SET @Result = @DocumentoTipoNombre + ': ' + @DocumentoNumero
			
		RETURN @Result
	END
GO



------------------------------------------------------------------------------------------
-- GETBOOLEANOSINO
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'udf_GetBooleanoSiNo' AND type = 'FN')
    DROP FUNCTION udf_GetBooleanoSiNo
GO

CREATE FUNCTION dbo.udf_GetBooleanoSiNo
	(@Valor bit)
	RETURNS varchar(2) AS
	
	BEGIN
		DECLARE @Result varchar(2)
		
		IF ISNULL(@Valor, '') = ''
			SET @Result = ''
		ELSE
			IF @Valor = 1
				SET @Result = 'Sí'
			ELSE
				SET @Result = 'No'
			
		RETURN @Result
	END
GO



------------------------------------------------------------------------------------------
-- GETBOOLEANOSI
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'udf_GetBooleanoSi' AND type = 'FN')
    DROP FUNCTION udf_GetBooleanoSi
GO

CREATE FUNCTION dbo.udf_GetBooleanoSi
	(@Valor bit)
	RETURNS varchar(2) AS
	
	BEGIN
		DECLARE @Result varchar(2)
		
		IF ISNULL(@Valor, '') = ''
			SET @Result = ''
		ELSE
			IF @Valor = 1
				SET @Result = 'Sí'
			ELSE
				SET @Result = ''
			
		RETURN @Result
	END
GO



------------------------------------------------------------------------------------------
-- GETENTIDADAPELLIDOYNOMBRE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'udf_GetEntidadApellidoYNombre' AND type = 'FN')
    DROP FUNCTION udf_GetEntidadApellidoYNombre
GO

CREATE FUNCTION dbo.udf_GetEntidadApellidoYNombre
	(@Apellido varchar(100),
	 @Nombre varchar(50))
	RETURNS varchar(150) AS
	
	BEGIN
		DECLARE @Result varchar(150)
		
		IF ISNULL(@Apellido, '') = ''
			SET @Result = ''
		ELSE
			IF ISNULL(@Nombre, '') = ''
				SET @Result = @Apellido
			ELSE
				SET @Result = @Apellido + ', ' + @Nombre
			
		RETURN @Result
	END
GO



------------------------------------------------------------------------------------------
-- GETPASAJEROSUBEBAJA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'udf_GetPasajeroSubeOBaja' AND type = 'FN')
    DROP FUNCTION udf_GetPasajeroSubeOBaja
GO

CREATE FUNCTION dbo.udf_GetPasajeroSubeOBaja
	(@IDRuta char(20),
	 @SubeBaja varchar(50),
	 @IDLugar int,
	 @IDLugarRuta int,
	 @LugarNombre varchar(100),
	 @LugarNombreCorto varchar(30))
	RETURNS varchar(100) AS
	
	BEGIN
		DECLARE @IDRutaOtra char(20)
		DECLARE @Result varchar(50)

		SET @IDRutaOtra = (SELECT Texto FROM Parametro WHERE IDParametro = 'Ruta_ID_Otra')
		
		IF @IDRuta = @IDRutaOtra
			SET @Result = ''
		ELSE
			IF @SubeBaja IS NULL
				IF @IDLugar = @IDLugarRuta
					SET @Result = ''
				ELSE
					IF @LugarNombreCorto IS NULL
						SET @Result = @LugarNombre
					ELSE
						SET @Result = @LugarNombreCorto
			ELSE
				SET @Result = @SubeBaja
			
		RETURN @Result
	END
GO