------------------------------------------------------------------------------------------
-- PERSONAALARMA_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_PersonaAlarma_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_PersonaAlarma_Data
GO

CREATE PROCEDURE dbo.sp_PersonaAlarma_Data 
	@IDPersona_FILTER int,
	@IDPersonaAlarmaGrupo_FILTER int AS

	SELECT PersonaAlarma.IDPersona, PersonaAlarma.IDPersonaAlarmaGrupo, PersonaAlarma.Fecha, PersonaAlarma.Preaviso, PersonaAlarma.Notas, PersonaAlarma.Activo, PersonaAlarma.FechaHoraCreacion, PersonaAlarma.IDUsuarioCreacion, PersonaAlarma.FechaHoraModificacion, PersonaAlarma.IDUsuarioModificacion
		FROM PersonaAlarma
		WHERE PersonaAlarma.IDPersona = @IDPersona_FILTER AND PersonaAlarma.IDPersonaAlarmaGrupo = @IDPersonaAlarmaGrupo_FILTER

GO



------------------------------------------------------------------------------------------
-- PERSONAALARMAGRUPO_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_PersonaAlarmaGrupo_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_PersonaAlarmaGrupo_Data
GO

CREATE PROCEDURE dbo.sp_PersonaAlarmaGrupo_Data 
	@IDPersonaAlarmaGrupo_FILTER int AS

	SELECT PersonaAlarmaGrupo.IDPersonaAlarmaGrupo, PersonaAlarmaGrupo.Nombre, PersonaAlarmaGrupo.Activo, PersonaAlarmaGrupo.Notas, PersonaAlarmaGrupo.FechaHoraCreacion, PersonaAlarmaGrupo.IDUsuarioCreacion, PersonaAlarmaGrupo.FechaHoraModificacion, PersonaAlarmaGrupo.IDUsuarioModificacion
		FROM PersonaAlarmaGrupo
		WHERE PersonaAlarmaGrupo.IDPersonaAlarmaGrupo = @IDPersonaAlarmaGrupo_FILTER

GO



------------------------------------------------------------------------------------------
-- PERSONAALARMAGRUPO_IDMAX
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_PersonaAlarmaGrupo_IDMax' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_PersonaAlarmaGrupo_IDMax
GO

CREATE PROCEDURE dbo.sp_PersonaAlarmaGrupo_IDMax AS
	SELECT Max(PersonaAlarmaGrupo.IDPersonaAlarmaGrupo) AS IDPersonaAlarmaGrupoMax
	FROM PersonaAlarmaGrupo
	 
GO
