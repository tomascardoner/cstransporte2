------------------------------------------------------------------------------------------
-- ALARMA_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM   sysobjects WHERE  name = N'usp_Alarma_Data' AND type = 'P')
    DROP PROCEDURE usp_Alarma_Data
GO

CREATE PROCEDURE dbo.usp_Alarma_Data 
	@IDAlarma int AS

	SELECT IDAlarma, Nombre, Fecha, Preaviso, Notas, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM Alarma
		WHERE IDAlarma = @IDAlarma

GO



------------------------------------------------------------------------------------------
-- ALARMA_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_Alarma_IDMax' AND type = 'P')
    DROP PROCEDURE usp_Alarma_IDMax
GO

CREATE PROCEDURE dbo.usp_Alarma_IDMax AS
	SELECT Max(IDAlarma) AS IDAlarmaMax
	FROM Alarma

GO
