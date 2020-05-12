-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	01/05/2014 12:59:42
-- Updated:	
-- Description: Obtiene los datos del Medio de Pago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_MedioPago_Get') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_MedioPago_Get
GO

CREATE PROCEDURE dbo.usp_MedioPago_Get
	@IDMedioPago tinyint 
AS

BEGIN
	SET NOCOUNT ON;

	SELECT MedioPago.IDMedioPago, MedioPago.Abreviatura, MedioPago.Nombre, MedioPago.UtilizaOperacion, MedioPago.IDMedioPagoPlan, MedioPago.IDCuentaCorrienteCaja, MedioPago.Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM MedioPago
		WHERE MedioPago.IDMedioPago = @IDMedioPago 

END
GO



-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	01/05/2014 12:59:42
-- Updated:	
-- Description: Agrega un Medio de Pago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_MedioPago_Add') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_MedioPago_Add
GO

CREATE PROCEDURE dbo.usp_MedioPago_Add
	@IDMedioPago tinyint OUTPUT,
	@Abreviatura varchar(4),
	@Nombre varchar(50),
	@UtilizaOperacion bit,
	@IDMedioPagoPlan tinyint,
	@IDCuentaCorrienteCaja int,
	@Activo bit,
	@IDUsuario smallint
AS

BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION

			SET @IDMedioPago = (SELECT ISNULL(MAX(MedioPago.IDMedioPago), 0) + 1 FROM MedioPago)
		
			INSERT INTO MedioPago
				(IDMedioPago, Abreviatura, Nombre, UtilizaOperacion, IDMedioPagoPlan, IDCuentaCorrienteCaja, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)
				VALUES (@IDMedioPago, @Abreviatura, @Nombre, @UtilizaOperacion, @IDMedioPagoPlan, @IDCuentaCorrienteCaja, @Activo, getdate(), @IDUsuario, getdate(), @IDUsuario)
	
		COMMIT TRANSACTION
	END TRY
	
	BEGIN CATCH
		IF @@TRANCOUNT > 0
			ROLLBACK TRANSACTION

		DECLARE @ErrorMessage NVARCHAR(4000);
		DECLARE @ErrorSeverity INT;
		DECLARE @ErrorState INT;

		SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();

		RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState)
	END CATCH

END
GO



-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	01/05/2014 12:59:42
-- Updated:	
-- Description: Actualiza los datos del Medio de Pago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_MedioPago_Update') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_MedioPago_Update
GO

CREATE PROCEDURE dbo.usp_MedioPago_Update
	@IDMedioPago tinyint,
	@Abreviatura varchar(4),
	@Nombre varchar(50),
	@UtilizaOperacion bit,
	@IDMedioPagoPlan tinyint,
	@IDCuentaCorrienteCaja int,
	@Activo bit,
	@IDUsuario smallint
AS

BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION	

				UPDATE MedioPago
					SET Abreviatura = @Abreviatura, Nombre = @Nombre, UtilizaOperacion = @UtilizaOperacion, IDMedioPagoPlan = @IDMedioPagoPlan, IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja, Activo = @Activo, FechaHoraModificacion = getdate(), IDUsuarioModificacion = @IDUsuario
					WHERE IDMedioPago = @IDMedioPago 

		COMMIT TRANSACTION
	END TRY
	
	BEGIN CATCH
		IF @@TRANCOUNT > 0
			ROLLBACK TRANSACTION

		DECLARE @ErrorMessage NVARCHAR(4000);
		DECLARE @ErrorSeverity INT;
		DECLARE @ErrorState INT;

		SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();

		RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState)
	END CATCH

END
GO



-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	01/05/2014 12:59:42
-- Updated:	
-- Description: Elimina un Medio de Pago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_MedioPago_Delete') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_MedioPago_Delete
GO

CREATE PROCEDURE dbo.usp_MedioPago_Delete 
	@IDMedioPago tinyint AS
	
BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION

			DELETE
				FROM MedioPago
				WHERE IDMedioPago = @IDMedioPago 

		COMMIT TRANSACTION
	END TRY
	
	BEGIN CATCH
		IF @@TRANCOUNT > 0
			ROLLBACK TRANSACTION

		DECLARE @ErrorMessage NVARCHAR(4000);
		DECLARE @ErrorSeverity INT;
		DECLARE @ErrorState INT;

		SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();

		RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState)
	END CATCH

END
GO



-- =============================================
-- Author:  Tomás A. Cardoner
-- Created: 11/05/2014 12:37:54
-- Updated:	
-- Description: Obtiene los datos del Plan de Medios de Pago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_MedioPagoPlan_Get') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_MedioPagoPlan_Get
GO

CREATE PROCEDURE dbo.usp_MedioPagoPlan_Get
	@IDMedioPagoPlan tinyint
AS

BEGIN
	SET NOCOUNT ON;

	SELECT MedioPagoPlan.IDMedioPagoPlan, MedioPagoPlan.Nombre, MedioPagoPlan.Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM MedioPagoPlan
		WHERE MedioPagoPlan.IDMedioPagoPlan = @IDMedioPagoPlan 

END
GO



-- =============================================
-- Author:  Tomás A. Cardoner
-- Created: 11/05/2014 12:37:54
-- Updated: 
-- Description: Agrega un Plan de Medios de Pago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_MedioPagoPlan_Add') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_MedioPagoPlan_Add
GO

CREATE PROCEDURE dbo.usp_MedioPagoPlan_Add
	@IDMedioPagoPlan tinyint OUTPUT, 
	@Nombre varchar(50),
	@Activo bit,
	@IDUsuario smallint
AS

BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION

			SET @IDMedioPagoPlan = (SELECT ISNULL(MAX(MedioPagoPlan.IDMedioPagoPlan), 0) + 1 FROM MedioPagoPlan)
		
			INSERT INTO MedioPagoPlan
				(IDMedioPagoPlan, Nombre, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)
				VALUES (@IDMedioPagoPlan, @Nombre, @Activo, getdate(), @IDUsuario, getdate(), @IDUsuario)
	
		COMMIT TRANSACTION
	END TRY
	
	BEGIN CATCH
		IF @@TRANCOUNT > 0
			ROLLBACK TRANSACTION

		DECLARE @ErrorMessage NVARCHAR(4000);
		DECLARE @ErrorSeverity INT;
		DECLARE @ErrorState INT;

		SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();

		RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState)
	END CATCH

END
GO



-- =============================================
-- Author:  Tomás A. Cardoner
-- Created: 11/05/2014 12:37:54
-- Updated: 
-- Description: Actualiza los datos del Plan de Medios de Pago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_MedioPagoPlan_Update') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_MedioPagoPlan_Update
GO

CREATE PROCEDURE dbo.usp_MedioPagoPlan_Update
	@IDMedioPagoPlan tinyint, 
	@Nombre varchar(50), 
	@Activo bit,
	@IDUsuario smallint
AS

BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION	

				UPDATE MedioPagoPlan
					SET Nombre = @Nombre, Activo = @Activo, FechaHoraModificacion = getdate(), IDUsuarioModificacion = @IDUsuario
					WHERE IDMedioPagoPlan = @IDMedioPagoPlan 

		COMMIT TRANSACTION
	END TRY
	
	BEGIN CATCH
		IF @@TRANCOUNT > 0
			ROLLBACK TRANSACTION

		DECLARE @ErrorMessage NVARCHAR(4000);
		DECLARE @ErrorSeverity INT;
		DECLARE @ErrorState INT;

		SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();

		RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState)
	END CATCH

END
GO



-- =============================================
-- Author:  Tomás A. Cardoner
-- Created: 11/05/2014 12:37:54
-- Updated: 
-- Description: Elimina un Plan de Medios de Pago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_MedioPagoPlan_Delete') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_MedioPagoPlan_Delete
GO

CREATE PROCEDURE dbo.usp_MedioPagoPlan_Delete 
	@IDMedioPagoPlan tinyint AS
	
BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION

			DELETE
				FROM MedioPagoPlan
				WHERE IDMedioPagoPlan = @IDMedioPagoPlan 

		COMMIT TRANSACTION
	END TRY
	
	BEGIN CATCH
		IF @@TRANCOUNT > 0
			ROLLBACK TRANSACTION

		DECLARE @ErrorMessage NVARCHAR(4000);
		DECLARE @ErrorSeverity INT;
		DECLARE @ErrorState INT;

		SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();

		RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState)
	END CATCH

END
GO



-- =============================================
-- Author:  Tomás A. Cardoner
-- Created: 11/05/2014 18:00:38
-- Updated: 
-- Description: Lista los datos de las Cuota del Plan de Medios de Pago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_MedioPagoPlanCuota_List') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_MedioPagoPlanCuota_List
GO

CREATE PROCEDURE dbo.usp_MedioPagoPlanCuota_List
	@IDMedioPagoPlan tinyint
AS

BEGIN
	SET NOCOUNT ON;

	SELECT MedioPagoPlanCuota.IDMedioPagoPlan, MedioPagoPlanCuota.Cuota, MedioPagoPlanCuota.Coeficiente, MedioPagoPlanCuota.CoeficientePrepago
		FROM MedioPagoPlanCuota
		WHERE MedioPagoPlanCuota.IDMedioPagoPlan = @IDMedioPagoPlan
		ORDER BY MedioPagoPlanCuota.Cuota

END
GO



-- =============================================
-- Author:  Tomás A. Cardoner
-- Created: 11/05/2014 12:45:38
-- Updated: 
-- Description: Obtiene los datos de la Cuota del Plan de Medios de Pago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_MedioPagoPlanCuota_Get') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_MedioPagoPlanCuota_Get
GO

CREATE PROCEDURE dbo.usp_MedioPagoPlanCuota_Get
	@IDMedioPagoPlan tinyint,
	@Cuota tinyint
AS

BEGIN
	SET NOCOUNT ON;

	SELECT MedioPagoPlanCuota.IDMedioPagoPlan, MedioPagoPlanCuota.Cuota, MedioPagoPlanCuota.Coeficiente, MedioPagoPlanCuota.CoeficientePrepago
		FROM MedioPagoPlanCuota
		WHERE MedioPagoPlanCuota.IDMedioPagoPlan = @IDMedioPagoPlan AND MedioPagoPlanCuota.Cuota = @Cuota

END
GO



-- =============================================
-- Author:	Tomás A. Cardoner
-- Created: 11/05/2014 12:45:38
-- Updated: 
-- Description: Agrega una Cuota del Plan de Medios de Pago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_MedioPagoPlanCuota_Add') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_MedioPagoPlanCuota_Add
GO

CREATE PROCEDURE dbo.usp_MedioPagoPlanCuota_Add
	@IDMedioPagoPlan tinyint, 
	@Cuota tinyint, 
	@Coeficiente decimal(5,1),
	@CoeficientePrepago decimal(5,1)
AS

BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION
		
			INSERT INTO MedioPagoPlanCuota
				(IDMedioPagoPlan, Cuota, Coeficiente, CoeficientePrepago)
				VALUES (@IDMedioPagoPlan, @Cuota, @Coeficiente, @CoeficientePrepago)
	
		COMMIT TRANSACTION
	END TRY
	
	BEGIN CATCH
		IF @@TRANCOUNT > 0
			ROLLBACK TRANSACTION

		DECLARE @ErrorMessage NVARCHAR(4000);
		DECLARE @ErrorSeverity INT;
		DECLARE @ErrorState INT;

		SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();

		RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState)
	END CATCH

END
GO



-- =============================================
-- Author:	Tomás A. Cardoner
-- Created: 11/05/2014 12:45:38
-- Updated: 
-- Description: Actualiza los datos de la Cuota del Plan de Medios de Pago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_MedioPagoPlanCuota_Update') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_MedioPagoPlanCuota_Update
GO

CREATE PROCEDURE dbo.usp_MedioPagoPlanCuota_Update
	@IDMedioPagoPlan tinyint,
	@Cuota tinyint,
	@Coeficiente decimal(5,1),
	@CoeficientePrepago decimal(5,1)
AS

BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION	

				UPDATE MedioPagoPlanCuota
					SET Coeficiente = @Coeficiente, CoeficientePrepago = @CoeficientePrepago
					WHERE IDMedioPagoPlan = @IDMedioPagoPlan AND Cuota = @Cuota 

		COMMIT TRANSACTION
	END TRY
	
	BEGIN CATCH
		IF @@TRANCOUNT > 0
			ROLLBACK TRANSACTION

		DECLARE @ErrorMessage NVARCHAR(4000);
		DECLARE @ErrorSeverity INT;
		DECLARE @ErrorState INT;

		SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();

		RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState)
	END CATCH

END
GO



-- =============================================
-- Author:	Tomás A. Cardoner
-- Created: 11/05/2014 12:45:38
-- Updated: 
-- Description: Elimina una Cuota del Plan de Medios de Pago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_MedioPagoPlanCuota_Delete') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_MedioPagoPlanCuota_Delete
GO

CREATE PROCEDURE dbo.usp_MedioPagoPlanCuota_Delete 
	@IDMedioPagoPlan tinyint,
	@Cuota tinyint AS
	
BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION

			DELETE
				FROM MedioPagoPlanCuota
				WHERE IDMedioPagoPlan = @IDMedioPagoPlan AND Cuota = @Cuota 

		COMMIT TRANSACTION
	END TRY
	
	BEGIN CATCH
		IF @@TRANCOUNT > 0
			ROLLBACK TRANSACTION

		DECLARE @ErrorMessage NVARCHAR(4000);
		DECLARE @ErrorSeverity INT;
		DECLARE @ErrorState INT;

		SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();

		RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState)
	END CATCH

END
GO