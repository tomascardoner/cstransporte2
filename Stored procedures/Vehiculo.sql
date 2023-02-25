------------------------------------------------------------------------------------------
-- VEHICULO_UTILIZACION
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Vehiculo_Utilizacion' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Vehiculo_Utilizacion
GO

CREATE PROCEDURE dbo.sp_Vehiculo_Utilizacion
	@Fecha smalldatetime AS

	SELECT Viaje.IDVehiculo, Viaje.FechaHora, RTRIM(Viaje.IDRuta) AS IDRuta, Viaje.Duracion, Viaje.Estado
		FROM Viaje
		WHERE CONVERT(CHAR(10), Viaje.FechaHora, 111) <= CONVERT(CHAR(10), @Fecha, 111)
			AND (CONVERT(CHAR(10), DATEADD(minute, Viaje.Duracion, Viaje.FechaHora), 111) >= CONVERT(CHAR(10), @Fecha, 111) OR (Viaje.Estado = 'EP' AND DATEDIFF(MINUTE, @Fecha, GETDATE()) >= 0))
			AND Viaje.Estado <> 'CA'
		ORDER BY Viaje.FechaHora

GO



------------------------------------------------------------------------------------------
-- VEHICULO LIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_Vehiculo_List' AND type = 'P')
    DROP PROCEDURE usp_Vehiculo_List
GO

CREATE PROCEDURE dbo.usp_Vehiculo_List
	@Activo bit AS

	SELECT v.IDVehiculo, v.Nombre, v.Marca, v.Modelo, v.Dominio, vc.Nombre AS Configuracion, v.Activo
		FROM Vehiculo AS v LEFT JOIN CSTransporte..VehiculoConfiguracion AS vc ON v.IDVehiculoConfiguracion = vc.IDVehiculoConfiguracion
		WHERE @Activo IS NULL OR v.Activo = @Activo
		ORDER BY vc.Nombre
GO


------------------------------------------------------------------------------------------
-- VEHICULO CONFIGURACION
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_VehiculoConfiguracion_List' AND type = 'P')
    DROP PROCEDURE usp_VehiculoConfiguracion_List
GO

CREATE PROCEDURE dbo.usp_VehiculoConfiguracion_List
	@MostrarNinguno bit AS

	IF @MostrarNinguno = 0
		SELECT vc.IDVehiculoConfiguracion, vc.Nombre
			FROM CSTransporte..VehiculoConfiguracion AS vc
			WHERE vc.Activo = 1
			ORDER BY vc.Nombre
	ELSE
		(SELECT 0 AS IDVehiculoConfiguracion, '« Ninguna »' AS Nombre, 1 AS Orden
			FROM Vehiculo)
		UNION
		(SELECT vc.IDVehiculoConfiguracion, vc.Nombre, 2 AS Orden
			FROM CSTransporte..VehiculoConfiguracion AS vc
			WHERE vc.Activo = 1)
		ORDER BY Nombre

GO