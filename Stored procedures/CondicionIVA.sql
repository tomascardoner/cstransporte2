------------------------------------------------------------------------------------------
-- CONDICIONIVA_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM   sysobjects WHERE  name = N'usp_CondicionIVA_Data' AND type = 'P')
    DROP PROCEDURE usp_CondicionIVA_Data
GO

CREATE PROCEDURE dbo.usp_CondicionIVA_Data 
	@IDCondicionIVA int AS

	SELECT IDCondicionIVA, Abreviatura, Nombre, DiscriminaIVA, PorcentajeRI, PorcentajeRNI, Notas, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM CondicionIVA
		WHERE IDCondicionIVA = @IDCondicionIVA

GO



------------------------------------------------------------------------------------------
-- CONDICIONIVA_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM sysobjects WHERE name = N'usp_CondicionIVA_IDMax' AND type = 'P')
    DROP PROCEDURE usp_CondicionIVA_IDMax
GO

CREATE PROCEDURE dbo.usp_CondicionIVA_IDMax AS
	SELECT Max(IDCondicionIVA) AS IDCondicionIVAMax
	FROM CondicionIVA

GO
