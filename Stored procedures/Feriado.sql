------------------------------------------------------------------------------------------
-- FERIADO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Feriado_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Feriado_Data
GO

CREATE PROCEDURE dbo.sp_Feriado_Data
	@Fecha_FILTER char(10) AS

	SELECT Feriado.Fecha, Feriado.Nombre, Feriado.FechaHoraCreacion, Feriado.IDUsuarioCreacion, Feriado.FechaHoraModificacion, Feriado.IDUsuarioModificacion
		FROM Feriado
		WHERE convert(char(10), Feriado.Fecha, 111) = @Fecha_FILTER

GO



------------------------------------------------------------------------------------------
-- FERIADO_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Feriado_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Feriado_List
GO

CREATE PROCEDURE dbo.sp_Feriado_List
	@Anio_FILTER smallint AS

	SELECT Feriado.Fecha, Feriado.Nombre, Feriado.FechaHoraCreacion, Feriado.IDUsuarioCreacion, Feriado.FechaHoraModificacion, Feriado.IDUsuarioModificacion
		FROM Feriado
		WHERE @Anio_FILTER IS NULL OR datepart(year, Feriado.Fecha) = @Anio_FILTER
		ORDER BY Feriado.Fecha

GO



------------------------------------------------------------------------------------------
-- FERIADO_VIAJEDETALLE_LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Feriado_ViajeDetalle_List' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Feriado_ViajeDetalle_List
GO

CREATE PROCEDURE dbo.sp_Feriado_ViajeDetalle_List
	@IDPersona_FILTER int AS

	SELECT Feriado.Nombre, ViajeDetalle.FechaHora, ViajeDetalle.IDRuta
		FROM Feriado, ViajeDetalle
		WHERE ViajeDetalle.IDPersona = @IDPersona_FILTER
				AND convert(char(10), Feriado.Fecha, 111) > convert(char(10), getdate(), 111)
				AND convert(char(10), Feriado.Fecha, 111) <= convert(char(10), dateadd(Day, 7, getdate()), 111) AND convert(char(10), Feriado.Fecha, 111) = convert(char(10), ViajeDetalle.FechaHora, 111)
				AND ViajeDetalle.Estado <> '3CA'
		ORDER BY Feriado.Fecha, ViajeDetalle.FechaHora, ViajeDetalle.IDRuta

GO