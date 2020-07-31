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



-- =============================================
-- Author: Tomás A. Cardoner
-- Created:	20/06/2020
-- Updated:	31/07/2020 - se cambio la funcionalidad ya que el horario es de exclusión y no de disponibilidad como estaba originalmente
-- Description: Verifica si un lugar de un ruta está disponible para el horario especificado
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaDetalle_HorarioDisponible') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaDetalle_HorarioDisponible
GO

CREATE PROCEDURE dbo.usp_RutaDetalle_HorarioDisponible
	@IDRuta char(20),
	@IDLugar int,
	@DiaSemana tinyint,
	@Hora time(0),
	@Disponible bit OUT
AS
	
BEGIN
	SET NOCOUNT ON;

	DECLARE @HoraInicio time(0)
	DECLARE @HoraFin time(0)

	-- Obtengo el horario en el detalle de la ruta
    SELECT @HoraInicio = HoraInicio, @HoraFin = HoraFin
		FROM RutaDetalle
		WHERE IDRuta = @IDRuta AND IDLugar = @IDLugar

	IF @HoraInicio IS NOT NULL AND @HoraFin IS NOT NULL
		BEGIN
		-- Se especificó un horario de exclusión
		IF @Hora BETWEEN @HoraInicio AND @HoraFin
			SET @Disponible = 0
		ELSE
			SET @Disponible = 1
		END
	ELSE
		BEGIN
		-- Busco si hay exclusiones especificadas en la tabla de Horarios del Detalle de la Ruta
		IF (SELECT COUNT(*) FROM RutaDetalleHorario WHERE IDRuta = @IDRuta AND IDLugar = @IDLugar) > 0
			-- Hay especificación de exclusiones de horarios, hay que buscar si está excluido para el horario del viaje
			BEGIN
			IF (SELECT COUNT(*)
					FROM RutaDetalleHorario
					WHERE IDRuta = @IDRuta AND IDLugar = @IDLugar
						AND DiaSemanaNumero = @DiaSemana
						AND @Hora BETWEEN HoraInicio AND HoraFin) > 0
				SET @Disponible = 0
			ELSE
				SET @Disponible = 1
			END
		ELSE
			-- No hay especificación de horarios para el lugar, por lo tanto está disponible siempre
			SET @Disponible = 1
		END

END
GO