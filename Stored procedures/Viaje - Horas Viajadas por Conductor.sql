-- =============================================
-- Author:		 Tomás A. Cardoner
-- Creation:     2021-09-03
-- Modification: 
-- Description:	 Devuelve las horas y minutos trabajados por cada conductor
-- =============================================
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'uspHorasViajadasPorConductor' AND type = 'P')
    DROP PROCEDURE uspHorasViajadasPorConductor
GO

CREATE PROCEDURE dbo.uspHorasViajadasPorConductor
	@FechaDesde smalldatetime,
	@FechaHasta smalldatetime,
	@IDConductor int,
	@IDRuta char(20)
	AS

	DECLARE @ViajePermite2Conductores bit

	SET @ViajePermite2Conductores = (SELECT SiNo FROM Parametro WHERE IDParametro = 'Viaje_Permite_2_Conductores')
	
	IF @ViajePermite2Conductores = 0
		SELECT c.Apellido + ', ' + c.Nombre AS Conductor, SUM(v.Duracion) AS DuracionTotal
			FROM Viaje AS v
				INNER JOIN Ruta AS r ON v.IDRuta = r.IDRuta
				INNER JOIN Persona AS c ON v.IDConductor = c.IDPersona
			WHERE v.Duracion IS NOT NULL
				AND (@FechaDesde IS NULL OR v.FechaHora >= @FechaDesde)
				AND (@FechaHasta IS NULL OR v.FechaHora <= DATETIMEFROMPARTS(YEAR(@FechaHasta), MONTH(@FechaHasta), DAY(@FechaHasta), 23, 59, 59, 99))
				AND (@IDConductor IS NULL OR v.IDConductor = @IDConductor)
				AND (@IDRuta IS NULL OR v.IDRuta = @IDRuta)
				AND v.Estado = 'FI'
			GROUP BY c.Apellido + ', ' + c.Nombre
GO



-- =============================================
-- Author:		 Tomás A. Cardoner
-- Creation:     2021-09-03
-- Modification: 
-- Description:	 Devuelve el detalle de las las horas y minutos trabajados por cada conductor
-- =============================================
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'uspHorasViajadasPorConductorDetalle' AND type = 'P')
    DROP PROCEDURE uspHorasViajadasPorConductorDetalle
GO

CREATE PROCEDURE dbo.uspHorasViajadasPorConductorDetalle
	@FechaDesde smalldatetime,
	@FechaHasta smalldatetime,
	@IDConductor int,
	@IDRuta char(20)
	AS

	DECLARE @ViajePermite2Conductores bit

	SET @ViajePermite2Conductores = (SELECT SiNo FROM Parametro WHERE IDParametro = 'Viaje_Permite_2_Conductores')
	
	IF @ViajePermite2Conductores = 0
		SELECT c.Apellido + ', ' + c.Nombre AS Conductor, v.FechaHora, v.IDRuta, v.Duracion
			FROM Viaje AS v
				INNER JOIN Ruta AS r ON v.IDRuta = r.IDRuta
				INNER JOIN Persona AS c ON v.IDConductor = c.IDPersona
			WHERE v.Duracion IS NOT NULL
				AND (@FechaDesde IS NULL OR v.FechaHora >= @FechaDesde)
				AND (@FechaHasta IS NULL OR v.FechaHora <= DATETIMEFROMPARTS(YEAR(@FechaHasta), MONTH(@FechaHasta), DAY(@FechaHasta), 23, 59, 59, 99))
				AND (@IDConductor IS NULL OR v.IDConductor = @IDConductor)
				AND (@IDRuta IS NULL OR v.IDRuta = @IDRuta)
				AND v.Estado = 'FI'
GO