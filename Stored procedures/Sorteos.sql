USE CSTransporte_LobosBus
GO

DECLARE @FechaHoraDesde smalldatetime = '2021-01-18 00:00:00'
DECLARE @FechaHoraHasta smalldatetime = '2021-01-24 23:59:59'

SELECT anu.Surname, anu.Names, LOWER(anu.Email) AS Email, vd.FechaHoraCreacion, vd.FechaHora, RTRIM(vd.IDRuta) AS Ruta, vd.Indice
	FROM ViajeDetalle as vd
		INNER JOIN CSTransporte_Web_LobosBus.dbo.Reserva AS r ON vd.IDViajeDetalle = r.idViajeDetalle
		INNER JOIN CSTransporte_Web_LobosBus.dbo.AspNetUsers AS anu ON r.idAspNetUserCliente = anu.Id
	WHERE vd.FechaHoraCreacion BETWEEN @FechaHoraDesde AND @FechaHoraHasta
        AND ((vd.ImporteContado <> 0 AND anu.PuedePagarEnOficina = 0) OR anu.PuedePagarEnOficina = 1)
        AND vd.IDUsuarioCreacion = 151
        AND vd.OcupanteTipo = 'PA'
		AND vd.Estado = '1CO'
	ORDER BY vd.FechaHoraCreacion