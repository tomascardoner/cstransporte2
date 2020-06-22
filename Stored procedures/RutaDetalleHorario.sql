-- =============================================
-- Author: Tomás A. Cardoner
-- Created:	21/06/2020
-- Updated:	
-- Description: Obtiene la lista de Horarios del Detalle de Ruta
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaDetalleHorario_List') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaDetalleHorario_List
GO

CREATE PROCEDURE dbo.usp_RutaDetalleHorario_List
	@IDRuta char(20),
	@IDLugar int
AS

BEGIN
	SET NOCOUNT ON;

	SELECT IDRutaDetalleHorario, DiaSemanaNumero, DiaSemana, HoraInicio, HoraFin
		FROM RutaDetalleHorario
		WHERE IDRuta = @IDRuta AND IDLugar = @IDLugar
		ORDER BY DiaSemanaNumero, HoraInicio

END
GO



-- =============================================
-- Author: Tomás A. Cardoner
-- Created:	02/06/2020
-- Updated:	
-- Description: Obtiene los datos de la RutaDetalleHorario
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaDetalleHorario_Get') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaDetalleHorario_Get
GO

CREATE PROCEDURE dbo.usp_RutaDetalleHorario_Get
	@IDRutaDetalleHorario int
AS

BEGIN
	SET NOCOUNT ON;

	SELECT IDRutaDetalleHorario, IDRuta, IDLugar, DiaSemanaNumero, DiaSemana, HoraInicio, HoraFin
		FROM RutaDetalleHorario
		WHERE IDRutaDetalleHorario = @IDRutaDetalleHorario

END
GO



-- =============================================
-- Author: Tomás A. Cardoner
-- Created:	02/06/2020
-- Updated:	
-- Description: Agrega un Horario al Detalle de Ruta
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaDetalleHorario_Add') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaDetalleHorario_Add
GO

CREATE PROCEDURE dbo.usp_RutaDetalleHorario_Add
	@IDRutaDetalleHorario int OUT,
	@IDRuta char(20),
	@IDLugar int,
	@DiaSemanaNumero int,
	@DiaSemana varchar(50),
	@HoraInicio time(7),
	@HoraFin time(7)
AS

BEGIN
	SET NOCOUNT ON;

    INSERT INTO RutaDetalleHorario
        (IDRuta, IDLugar, DiaSemanaNumero, DiaSemana, HoraInicio, HoraFin)
        VALUES (@IDRuta, @IDLugar, @DiaSemanaNumero, @DiaSemana, @HoraInicio, @HoraFin)

	SET @IDRutaDetalleHorario = @@IDENTITY

END
GO



-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	02/06/2020
-- Updated:	
-- Description: Actualiza los datos del Horario del Detalle de Ruta
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaDetalleHorario_Update') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaDetalleHorario_Update
GO

CREATE PROCEDURE dbo.usp_RutaDetalleHorario_Update
	@IDRutaDetalleHorario int,
	@IDRuta char(20),
	@IDLugar int,
	@DiaSemanaNumero int,
	@DiaSemana varchar(50),
	@HoraInicio time(7),
	@HoraFin time(7)
AS

BEGIN
	SET NOCOUNT ON;

    UPDATE RutaDetalleHorario
        SET IDRuta = @IDRuta, IDLugar = @IDLugar, DiaSemanaNumero = @DiaSemanaNumero, DiaSemana = @DiaSemana, HoraInicio = @HoraInicio, HoraFin = @HoraFin
        WHERE IDRutaDetalleHorario = @IDRutaDetalleHorario

END
GO



-- =============================================
-- Author: Tomás A. Cardoner
-- Created:	02/06/2020
-- Updated:	
-- Description: Elimina el Horario del Detalle de Ruta
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaDetalleHorario_Delete') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaDetalleHorario_Delete
GO

CREATE PROCEDURE dbo.usp_RutaDetalleHorario_Delete
	@IDRutaDetalleHorario int
AS
	
BEGIN
	SET NOCOUNT ON;

    DELETE
        FROM RutaDetalleHorario
        WHERE IDRutaDetalleHorario = @IDRutaDetalleHorario

END
GO