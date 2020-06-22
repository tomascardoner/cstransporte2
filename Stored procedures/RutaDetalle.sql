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
	   WHERE  name = N'usp_RutaDetalle_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE usp_RutaDetalle_List
GO

CREATE PROCEDURE dbo.usp_RutaDetalle_List
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
-- Author: Tomás A. Cardoner
-- Created:	20/06/2020
-- Updated:	
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
		-- Se especificó un horario
		IF @Hora BETWEEN @HoraInicio AND @HoraFin
			SET @Disponible = 1
		ELSE
			SET @Disponible = 0
		END
	ELSE
		BEGIN
		-- Busco si hay horarios especificados en la tabla de Horarios del Detalle de la Ruta
		IF (SELECT COUNT(*) FROM RutaDetalleHorario WHERE IDRuta = @IDRuta AND IDLugar = @IDLugar) > 0
			-- Hay especificación de horarios, hay que buscar si está disponible para el horario del viaje
			BEGIN
			IF (SELECT COUNT(*)
					FROM RutaDetalleHorario
					WHERE IDRuta = @IDRuta AND IDLugar = @IDLugar
						AND (DiaSemanaNumero = 0 OR DiaSemanaNumero = @DiaSemana)
						AND @Hora BETWEEN HoraInicio AND HoraFin) > 0
				SET @Disponible = 1
			ELSE
				SET @Disponible = 0
			END
		ELSE
			-- No hay especificación de horarios para el lugar, por lo tanto está disponible siempre
			SET @Disponible = 1
		END

END
GO