------------------------------------------------------------------------------------------
-- SEMAFOROGENERAL_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name FROM   sysobjects WHERE  name = N'usp_SemaforoGeneral_Data' AND type = 'P')
    DROP PROCEDURE usp_SemaforoGeneral_Data
GO

CREATE PROCEDURE dbo.usp_SemaforoGeneral_Data AS

	SELECT IDSemaforo, ValorTimer, ExtraTexto, ExtraNumero, ExtraFecha, ExtraSiNo
		FROM SemaforoGeneral

GO