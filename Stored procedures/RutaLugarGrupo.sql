-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	01/02/2019 09:43
-- Updated:	
-- Description: Obtiene los datos de la Ruta-LugarGrupo
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaLugarGrupo_Get') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaLugarGrupo_Get
GO

CREATE PROCEDURE dbo.usp_RutaLugarGrupo_Get
	@IDRuta char(20),
    @IDLugarGrupo int
AS

BEGIN
	SET NOCOUNT ON;

	SELECT RutaLugarGrupo.IDRuta, RutaLugarGrupo.IDLugarGrupo, RutaLugarGrupo.IDLugarPredeterminado
		FROM RutaLugarGrupo
		WHERE RutaLugarGrupo.IDRuta = @IDRuta AND RutaLugarGrupo.IDLugarGrupo = @IDLugarGrupo 

END
GO



-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	01/02/2019 09:46
-- Updated:	
-- Description: Agrega un Ruta-LugarGrupo
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaLugarGrupo_Add') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaLugarGrupo_Add
GO

CREATE PROCEDURE dbo.usp_RutaLugarGrupo_Add
	@IDRuta char(20),
	@IDLugarGrupo int,
	@IDLugarPredeterminado int
AS

BEGIN
	SET NOCOUNT ON;

    INSERT INTO RutaLugarGrupo
        (IDRuta, IDLugarGrupo, IDLugarPredeterminado)
        VALUES (@IDRuta, @IDLugarGrupo, @IDLugarPredeterminado)
END
GO



-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	01/02/2019 09:48
-- Updated:	
-- Description: Actualiza los datos de la Ruta-LugarGrupo
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaLugarGrupo_Update') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaLugarGrupo_Update
GO

CREATE PROCEDURE dbo.usp_RutaLugarGrupo_Update
	@IDRuta char(20),
	@IDLugarGrupo int,
	@IDLugarPredeterminado int
AS

BEGIN
	SET NOCOUNT ON;

    UPDATE RutaLugarGrupo
        SET IDLugarPredeterminado = @IDLugarPredeterminado
        WHERE IDRuta = @IDRuta AND IDLugarGrupo = @IDLugarGrupo
END
GO



-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	01/02/2019 09:51
-- Updated:	
-- Description: Elimina una Ruta-LugarGrupo
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_RutaLugarGrupo_Delete') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_RutaLugarGrupo_Delete
GO

CREATE PROCEDURE dbo.usp_RutaLugarGrupo_Delete 
	@IDRuta char(20),
	@IDLugarGrupo int
AS
	
BEGIN
	SET NOCOUNT ON;

    DELETE
        FROM RutaLugarGrupo
        WHERE IDRuta = @IDRuta AND IDLugarGrupo = @IDLugarGrupo

END
GO