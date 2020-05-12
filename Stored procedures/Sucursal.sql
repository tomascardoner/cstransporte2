------------------------------------------------------------------------------------------
-- SUCURSAL_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM   sysobjects WHERE  name = N'usp_Sucursal_Data' AND type = 'P')
    DROP PROCEDURE usp_Sucursal_Data
GO

CREATE PROCEDURE dbo.usp_Sucursal_Data 
	@IDSucursal char(3) AS

	SELECT IDSucursal, Nombre, CodigoFacturacion, Email, Notas, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM Sucursal
		WHERE IDSucursal = @IDSucursal

GO
