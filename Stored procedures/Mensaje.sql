------------------------------------------------------------------------------------------
-- MENSAJE_GETLIST
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_Mensaje_GetList' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_Mensaje_GetList
GO

CREATE PROCEDURE dbo.sp_Mensaje_GetList
	@IDUsuario smallint AS

	SELECT Mensaje.IDMensaje, Mensaje.Mensaje
		FROM (Mensaje LEFT JOIN Usuario ON Mensaje.IDUsuarioGrupo = Usuario.IDUsuarioGrupo) LEFT JOIN Mensaje_Usuario ON Mensaje_Usuario.IDMensaje = Mensaje.IDMensaje AND Mensaje_Usuario.IDUsuario = Usuario.IDUsuario
		WHERE (Mensaje.IDUsuarioGrupo IS NULL OR Usuario.IDUsuario = @IDUsuario)
			AND (Mensaje.FechaInicio IS NULL OR getdate() >= Mensaje.FechaInicio)
			AND (Mensaje.FechaFin IS NULL OR getdate() <= Mensaje.FechaFin)
			AND (Mensaje_Usuario.LeidoVeces IS NULL OR Mensaje_Usuario.LeidoVeces < Mensaje.RepetirVeces)
			AND Mensaje.Activo = 1
GO
