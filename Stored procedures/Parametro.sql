------------------------------------------------------------------------------------------
-- PARAMETRO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM   sysobjects WHERE  name = N'usp_Parametro_AllData' AND type = 'P')
    DROP PROCEDURE usp_Parametro_AllData
GO

CREATE PROCEDURE dbo.usp_Parametro_AllData AS

	SELECT IDParametro, Texto, NumeroEntero, NumeroDecimal, Moneda, FechaHora, SiNo
		FROM Parametro

GO