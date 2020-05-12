-- =============================================
-- Author: Tomás A. Cardoner
-- Created:	30/04/2014 21:43:11
-- Updated:	
-- Description: Obtiene los datos del Prepago de la Persona
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_PersonaPrepago_Get') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_PersonaPrepago_Get
GO

CREATE PROCEDURE dbo.usp_PersonaPrepago_Get
	@IDPersona int, 
	@IDRutaGrupo int, 
	@FechaInicio smalldatetime 
AS

BEGIN
	SET NOCOUNT ON;

	SELECT PersonaPrepago.IDPersona, PersonaPrepago.IDRutaGrupo, PersonaPrepago.FechaInicio, PersonaPrepago.FechaFin, PersonaPrepago.IDListaPrecio, PersonaPrepago.ImporteOriginal, PersonaPrepago.Importe, PersonaPrepago.IDMedioPago, PersonaPrepago.Cuotas, PersonaPrepago.Operacion, PersonaPrepago.FacturaNumero, PersonaPrepago.IDCuentaCorrienteCaja, PersonaPrepago.IDMovimiento_Credito, PersonaPrepago.IDMovimiento_Debito, PersonaPrepago.FechaHoraCreacion, PersonaPrepago.IDUsuarioCreacion, PersonaPrepago.FechaHoraModificacion, PersonaPrepago.IDUsuarioModificacion
		FROM PersonaPrepago
		WHERE PersonaPrepago.IDPersona = @IDPersona AND PersonaPrepago.IDRutaGrupo = @IDRutaGrupo AND PersonaPrepago.FechaInicio = @FechaInicio 

END
GO



-- =============================================
-- Author: Tomás A. Cardoner
-- Created:	30/04/2014 21:43:11
-- Updated:	
-- Description: Agrega un Prepago de una Persona
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_PersonaPrepago_Add') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_PersonaPrepago_Add
GO

CREATE PROCEDURE dbo.usp_PersonaPrepago_Add
	@IDPersona int,
	@IDRutaGrupo int,
	@FechaInicio smalldatetime,
	@FechaFin smalldatetime,
	@IDListaPrecio int,
	@ImporteOriginal smallmoney,
	@Importe smallmoney,
	@IDMedioPago tinyint,
	@Cuotas tinyint,
	@Operacion varchar(20),
	@FacturaNumero varchar(20),
	@IDCuentaCorrienteCaja int,
	@IDMovimiento_Credito int,
	@IDMovimiento_Debito int,
	@IDUsuario smallint
AS

BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION
		
			INSERT INTO PersonaPrepago
				(IDPersona, IDRutaGrupo, FechaInicio, FechaFin, IDListaPrecio, ImporteOriginal, Importe, IDMedioPago, Cuotas, Operacion, FacturaNumero, IDCuentaCorrienteCaja, IDMovimiento_Credito, IDMovimiento_Debito, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)
				VALUES (@IDPersona, @IDRutaGrupo, @FechaInicio, @FechaFin, @IDListaPrecio, @ImporteOriginal, @Importe, @IDMedioPago, @Cuotas, @Operacion, @FacturaNumero, @IDCuentaCorrienteCaja, @IDMovimiento_Credito, @IDMovimiento_Debito, GETDATE(), @IDUsuario, GETDATE(), @IDUsuario)
	
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
-- Author: Tomás A. Cardoner
-- Created:	30/04/2014 21:43:11
-- Updated:	
-- Description: Actualiza los datos del PersonaPrepago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_PersonaPrepago_Update') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_PersonaPrepago_Update
GO

CREATE PROCEDURE dbo.usp_PersonaPrepago_Update
	@IDRutaGrupo_Original int,
	@FechaInicio_Original smalldatetime,
	@IDPersona int,
	@IDRutaGrupo int,
	@FechaInicio smalldatetime,
	@FechaFin smalldatetime,
	@IDListaPrecio int,
	@ImporteOriginal smallmoney,
	@Importe smallmoney,
	@IDMedioPago tinyint,
	@Cuotas tinyint,
	@Operacion varchar(20),
	@FacturaNumero varchar(20),
	@IDCuentaCorrienteCaja int,
	@IDMovimiento_Credito int,
	@IDMovimiento_Debito int,
	@IDUsuario smallint
AS

BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION	

				UPDATE PersonaPrepago
					SET IDRutaGrupo = @IDRutaGrupo, FechaInicio = @FechaInicio, FechaFin = @FechaFin, IDListaPrecio = @IDListaPrecio, ImporteOriginal = @ImporteOriginal, Importe = @Importe, IDMedioPago = @IDMedioPago, Cuotas = @Cuotas, Operacion = @Operacion, FacturaNumero = @FacturaNumero, IDCuentaCorrienteCaja = @IDCuentaCorrienteCaja, IDMovimiento_Credito = @IDMovimiento_Credito, IDMovimiento_Debito = @IDMovimiento_Debito, FechaHoraModificacion = GETDATE(), IDUsuarioModificacion = @IDUsuario
					WHERE IDPersona = @IDPersona AND IDRutaGrupo = @IDRutaGrupo_Original AND FechaInicio = @FechaInicio_Original

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
-- Created:	30/04/2014 21:43:11
-- Updated:	
-- Description: Elimina un PersonaPrepago
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_PersonaPrepago_Delete') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_PersonaPrepago_Delete
GO

CREATE PROCEDURE dbo.usp_PersonaPrepago_Delete 
	@IDPersona int, 
	@IDRutaGrupo int, 
	@FechaInicio smalldatetime AS
	
BEGIN
	DECLARE @IDMovimiento_Credito int
	DECLARE @IDMovimiento_Debito int
	
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION
			--GUARDO LOS ID DE LOS MOVIMIENTOS DE CUENTA CORRIENTE PARA ELIMINARLOS
			SELECT @IDMovimiento_Credito = ISNULL(IDMovimiento_Credito, 0), @IDMovimiento_Debito = ISNULL(IDMovimiento_Debito, 0)
				FROM PersonaPrepago
				WHERE IDPersona = @IDPersona AND IDRutaGrupo = @IDRutaGrupo AND FechaInicio = @FechaInicio
			
			--ELIMINO EL PREPAGO
			DELETE
				FROM PersonaPrepago
				WHERE IDPersona = @IDPersona AND IDRutaGrupo = @IDRutaGrupo AND FechaInicio = @FechaInicio

			--ELIMINO LOS MOVIMIENTOS DE CUENTA CORRIENTE CORRESPONDIENTES AL CRÉDITO Y DEBITO
			DELETE
				FROM CuentaCorriente
				WHERE IDMovimiento = @IDMovimiento_Credito OR IDMovimiento = @IDMovimiento_Debito
			
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
-- Author: Tomás A. Cardoner
-- Created:	30/04/2014 22:34
-- Updated:	07/05/2014 12:59 - Funciona el contador de Reservas
-- Description: Devuelve los datos del Prepago de una Persona tomando en cuenta las fechas
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_PersonaPrepago_CheckIfApply') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_PersonaPrepago_CheckIfApply
GO

CREATE PROCEDURE dbo.usp_PersonaPrepago_CheckIfApply
	@IDPersona int,
	@IDRuta char(20),
	@Fecha smalldatetime,
	@IDListaPrecio int OUTPUT
AS

BEGIN
	DECLARE @IDRutaGrupo int
	DECLARE @FechaInicio smalldatetime
	DECLARE @FechaFin smalldatetime
	DECLARE @ReservasCantidadLimite smallint
	DECLARE @ReservasCantidad smallint
	
	SET NOCOUNT ON;
	
	SET @IDRutaGrupo = (SELECT IDRutaGrupo FROM Ruta WHERE IDRuta = @IDRuta)

	SELECT @IDListaPrecio = IDListaPrecio, @FechaInicio = FechaInicio, @FechaFin = FechaFin FROM PersonaPrepago WHERE IDPersona = @IDPersona AND IDRutaGrupo = @IDRutaGrupo AND @Fecha BETWEEN FechaInicio AND FechaFin
	SET @IDListaPrecio = ISNULL(@IDListaPrecio, 0)
	
	IF @IDListaPrecio <> 0
		BEGIN
		IF ISNULL((SELECT Activo FROM ListaPrecio WHERE IDListaPrecio = @IDListaPrecio), 0) = 0
			BEGIN
			--LA LISTA DE PRECIOS ESTÁ DESACTIVADA, ASÍ QUE DEVUELVO CERO
			SET @IDListaPrecio = 0
			END
		ELSE
			BEGIN
			--LA LISTA DE PRECIOS ESTÁ OK, VERIFICO LA CANTIDAD DE RESERVAS DEL PREPAGO
			SET @ReservasCantidadLimite = ISNULL((SELECT PrepagoReservasCantidad FROM ListaPrecio WHERE IDListaPrecio = @IDListaPrecio), 0)
			IF @ReservasCantidadLimite > 0
				BEGIN
				--EL PREPAGO ESPECIFICA CANTIDAD DE RESERVAS, ASÍ QUE CUENTO LAS RESERVAS REALIZADAS PARA VERIFICAR QUE NO SE HAYA EXCEDIDO
				SET @ReservasCantidad = ISNULL((SELECT COUNT(*) FROM (Viaje INNER JOIN Ruta ON Viaje.IDRuta = Ruta.IDRuta) INNER JOIN ViajeDetalle ON Viaje.FechaHora = ViajeDetalle.FechaHora AND Viaje.IDRuta = ViajeDetalle.IDRuta WHERE Viaje.Estado <> 'CA' AND ViajeDetalle.Estado <> '3CA' AND ViajeDetalle.OcupanteTipo = 'PA' AND ViajeDetalle.IDPersona = @IDPersona AND Ruta.IDRutaGrupo = @IDRutaGrupo AND (Viaje.FechaHora BETWEEN @FechaInicio AND @FechaFin) AND ViajeDetalle.IDListaPrecio = @IDListaPrecio), 0)
				IF @ReservasCantidad >= @ReservasCantidadLimite
					BEGIN
					--YA TIENE LA CANTIDAD DE RESERVAS PERMITIDAS, ASÍ QUE NO APLICA EL PREPAGO
					SET @IDListaPrecio = 0
					END
				END
			END
		END
END
GO



-- =============================================
-- Author: Tomás A. Cardoner
-- Created:	17/05/2014 16:27
-- Updated:	
-- Description: Devuelve los viajes de una Persona entre las fechas especificadas
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_PersonaPrepago_GetViajeDetalleListForRecount') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_PersonaPrepago_GetViajeDetalleListForRecount
GO

CREATE PROCEDURE dbo.usp_PersonaPrepago_GetViajeDetalleListForRecount
	@IDPersona int,
	@IDRuta char(20),
	@Fecha smalldatetime,
	@IDListaPrecio int
AS

BEGIN
	DECLARE @IDRutaGrupo int
	DECLARE @FechaInicio smalldatetime
	DECLARE @FechaFin smalldatetime
	
	SET NOCOUNT ON;
	
	SET @IDRutaGrupo = (SELECT IDRutaGrupo FROM Ruta WHERE IDRuta = @IDRuta)
	
	SELECT @FechaInicio = FechaInicio, @FechaFin = FechaFin FROM PersonaPrepago WHERE IDPersona = @IDPersona AND IDRutaGrupo = @IDRutaGrupo AND @Fecha BETWEEN FechaInicio AND FechaFin
	
	--OBTENGO LAS RESERVAS QUE NO TIENEN APLICADA LA MISMA LISTA DE PRECIOS Y QUE ESTAN DENTRO DE LAS FECHAS
	SELECT ViajeDetalle.FechaHora, ViajeDetalle.IDRuta, ViajeDetalle.Indice, ViajeDetalle.IDListaPrecio
		FROM (Viaje INNER JOIN Ruta ON Viaje.IDRuta = Ruta.IDRuta) INNER JOIN ViajeDetalle ON Viaje.FechaHora = ViajeDetalle.FechaHora AND Viaje.IDRuta = ViajeDetalle.IDRuta
		WHERE Viaje.Estado <> 'CA' AND ViajeDetalle.Estado <> '3CA' AND ViajeDetalle.OcupanteTipo = 'PA' AND ViajeDetalle.IDPersona = @IDPersona AND Ruta.IDRutaGrupo = @IDRutaGrupo AND (Viaje.FechaHora BETWEEN @FechaInicio AND @FechaFin) AND ViajeDetalle.IDListaPrecio <> @IDListaPrecio

END
GO



-- =============================================
-- Author: Tomás A. Cardoner
-- Created:	09/06/2014 22:08
-- Updated:	
-- Description: Devuelve la cantidad de viajes prepagos restantes de una Persona
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_PersonaPrepago_GetViajeDetalleRemain') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_PersonaPrepago_GetViajeDetalleRemain
GO

CREATE PROCEDURE dbo.usp_PersonaPrepago_GetViajeDetalleRemain
	@IDPersona int,
	@IDRuta char(20),
	@Fecha smalldatetime,
	@IDListaPrecio int,
	@ReservasRestantes smallint OUTPUT
AS

BEGIN
	DECLARE @IDRutaGrupo int
	DECLARE @FechaInicio smalldatetime
	DECLARE @FechaFin smalldatetime
	DECLARE @PrepagoReservasCantidad smallint
	DECLARE @ReservasAnterioresCantidad smallint

	SET NOCOUNT ON;
	
	SET @IDRutaGrupo = (SELECT IDRutaGrupo FROM Ruta WHERE IDRuta = @IDRuta)
	
	SELECT @FechaInicio = FechaInicio, @FechaFin = FechaFin FROM PersonaPrepago WHERE IDPersona = @IDPersona AND IDRutaGrupo = @IDRutaGrupo AND @Fecha BETWEEN FechaInicio AND FechaFin
	
	SELECT @PrepagoReservasCantidad = PrepagoReservasCantidad FROM ListaPrecio WHERE IDListaPrecio = @IDListaPrecio
	
	--CUENTO LAS RESERVAS QUE TIENEN APLICADA LA MISMA LISTA DE PRECIOS Y QUE ESTAN DENTRO DE LAS FECHAS DEL PREPAGO
	SET @ReservasAnterioresCantidad = (SELECT COUNT(*)
			FROM (Viaje INNER JOIN Ruta ON Viaje.IDRuta = Ruta.IDRuta) INNER JOIN ViajeDetalle ON Viaje.FechaHora = ViajeDetalle.FechaHora AND Viaje.IDRuta = ViajeDetalle.IDRuta
			WHERE Viaje.Estado <> 'CA' AND ViajeDetalle.Estado <> '3CA' AND ViajeDetalle.OcupanteTipo = 'PA' AND ViajeDetalle.IDPersona = @IDPersona AND Ruta.IDRutaGrupo = @IDRutaGrupo AND (Viaje.FechaHora BETWEEN @FechaInicio AND @Fecha) AND ViajeDetalle.IDListaPrecio = @IDListaPrecio)

	SET @ReservasRestantes = (@PrepagoReservasCantidad - @ReservasAnterioresCantidad)

END
GO



-- =============================================
-- Author: Tomás A. Cardoner
-- Created:	17/05/2014 19:19
-- Updated:	
-- Description: Devuelve las Reservas con Lista de Precio Prepaga de un Viaje para verificar los prepagos de Otras Reservas
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_PersonaPrepago_GetViajeDetalleConListaPrepagoList') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_PersonaPrepago_GetViajeDetalleConListaPrepagoList
GO

CREATE PROCEDURE dbo.usp_PersonaPrepago_GetViajeDetalleConListaPrepagoList
	@FechaHora smalldatetime,
	@IDRuta char(20)
AS

BEGIN
	SET NOCOUNT ON;
	
	SELECT ViajeDetalle.FechaHora, ViajeDetalle.IDRuta, ViajeDetalle.Indice
		FROM ViajeDetalle INNER JOIN ListaPrecio ON ViajeDetalle.IDListaPrecio = ListaPrecio.IDListaPrecio
		WHERE ViajeDetalle.FechaHora = @FechaHora AND ViajeDetalle.IDRuta = @IDRuta AND ViajeDetalle.Estado <> '3CA' AND ViajeDetalle.OcupanteTipo = 'PA' AND ListaPrecio.PrepagoEs = 1 AND ListaPrecio.PrepagoReservasCantidad > 0

END
GO



-- =============================================
-- Author: Tomás A. Cardoner
-- Created:	17/05/2014 19:40
-- Updated:	
-- Description: Devuelve las Reservas de un Viaje para verificar los prepagos
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_PersonaPrepago_GetViajeDetalleList') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_PersonaPrepago_GetViajeDetalleList
GO

CREATE PROCEDURE dbo.usp_PersonaPrepago_GetViajeDetalleList
	@FechaHora smalldatetime,
	@IDRuta char(20)
AS

BEGIN
	SET NOCOUNT ON;
	
	SELECT ViajeDetalle.FechaHora, ViajeDetalle.IDRuta, ViajeDetalle.Indice
		FROM ViajeDetalle INNER JOIN ListaPrecio ON ViajeDetalle.IDListaPrecio = ListaPrecio.IDListaPrecio
		WHERE ViajeDetalle.FechaHora = @FechaHora AND ViajeDetalle.IDRuta = @IDRuta AND ViajeDetalle.Estado <> '3CA' AND ViajeDetalle.OcupanteTipo = 'PA' 

END
GO