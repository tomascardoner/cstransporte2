-- =============================================
-- Author:		Tomás A. Cardoner
-- Create date: 2020-09-01 13:08
-- Description:	Obtiene las reservas web
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'uspObtenerReservasWebPagadasConMP') AND type in (N'P', N'PC'))
	 DROP PROCEDURE uspObtenerReservasWebPagadasConMP
GO

CREATE PROCEDURE uspObtenerReservasWebPagadasConMP
	@FechaHoraDesde smalldatetime,
	@FechaHoraHasta smalldatetime
	AS

	BEGIN
		-- SET NOCOUNT ON added to prevent extra result sets from interfering with SELECT statements.
		SET NOCOUNT ON;

		SELECT vd.IDViajeDetalle, vd.FechaHora, vd.IDRuta, vd.Orden, dbo.udf_GetEntidadApellidoYNombre(p.Apellido, p.Nombre) AS ApellidoNombre, vd.FacturaNumero
			FROM ViajeDetalle AS vd
                INNER JOIN Persona AS p ON vd.IDPersona = p.IDPersona
			WHERE vd.FechaHora BETWEEN @FechaHoraDesde AND @FechaHoraHasta
				AND vd.IDUsuarioCreacion = 151 AND vd.ImporteContado > 0 AND vd.Estado = '1CO' AND vd.OcupanteTipo = 'PA'
				AND vd.IDMedioPago = 9
	END
GO

GRANT EXECUTE ON dbo.uspObtenerReservasWebPagadasConMP TO cstransporte
GO