
-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	26/04/2014 18:03:40
-- Updated:	
-- Description: Obtiene los datos de la Lista de Precios
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_ListaPrecio_Get') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_ListaPrecio_Get
GO

CREATE PROCEDURE dbo.usp_ListaPrecio_Get
	@IDListaPrecio int 
AS

BEGIN
	SET NOCOUNT ON;

	SELECT ListaPrecio.IDListaPrecio, ListaPrecio.Nombre, ListaPrecio.Leyenda, ListaPrecio.Descripcion, ListaPrecio.PrepagoEs, ListaPrecio.PrepagoVencimiento, ListaPrecio.PrepagoReservasCantidad, ListaPrecio.IDCuentaCorrienteGrupo_Credito, ListaPrecio.IDCuentaCorrienteGrupo_Debito, ListaPrecio.Notas, ListaPrecio.Activo, ListaPrecio.FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM ListaPrecio
		WHERE ListaPrecio.IDListaPrecio = @IDListaPrecio 

END
GO



-- =============================================
-- Author:		Tomás A. Cardoner
-- Created:	26/04/2014 18:03:40
-- Updated:	15/02/2018 19:31
-- Description: Agrega una Lista de Precios
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_ListaPrecio_Add') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_ListaPrecio_Add
GO

CREATE PROCEDURE dbo.usp_ListaPrecio_Add
	@IDListaPrecio int OUTPUT,
	@Nombre varchar(50),
	@Leyenda varchar(50),
	@Descripcion varchar(8000),
	@PrepagoEs bit,
	@PrepagoVencimiento char(3),
	@PrepagoReservasCantidad smallint,
	@IDCuentaCorrienteGrupo_Credito int,
	@IDCuentaCorrienteGrupo_Debito int,
	@Notas varchar(8000),
	@Activo bit,
	@IDUsuario smallint
AS

BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION

			SET @IDListaPrecio = (SELECT ISNULL(MAX(ListaPrecio.IDListaPrecio), 0) + 1 FROM ListaPrecio)
		
			INSERT INTO ListaPrecio
				(IDListaPrecio, Nombre, Leyenda, Descripcion, PrepagoEs, PrepagoVencimiento, PrepagoReservasCantidad, IDCuentaCorrienteGrupo_Credito, IDCuentaCorrienteGrupo_Debito, Notas, Activo, FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion)
				VALUES (@IDListaPrecio, @Nombre, @Leyenda, @Descripcion, @PrepagoEs, @PrepagoVencimiento, @PrepagoReservasCantidad, @IDCuentaCorrienteGrupo_Credito, @IDCuentaCorrienteGrupo_Debito, @Notas, @Activo, getdate(), @IDUsuario, getdate(), @IDUsuario)
	
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
-- Created:	26/04/2014 18:03:40
-- Updated:	15/02/2018 19:31
-- Description: Actualiza los datos de la Lista de Precios
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_ListaPrecio_Update') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_ListaPrecio_Update
GO

CREATE PROCEDURE dbo.usp_ListaPrecio_Update
	@IDListaPrecio int,
	@Nombre varchar(50),
	@Leyenda varchar(50),
	@Descripcion varchar(8000),
	@PrepagoEs bit,
	@PrepagoVencimiento char(3),
	@PrepagoReservasCantidad smallint,
	@IDCuentaCorrienteGrupo_Credito int,
	@IDCuentaCorrienteGrupo_Debito int,
	@Notas varchar(8000),
	@Activo bit,
	@IDUsuario smallint
AS

BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION	

				UPDATE ListaPrecio
					SET Nombre = @Nombre, Leyenda = @Leyenda, Descripcion = @Descripcion, PrepagoEs = @PrepagoEs, PrepagoVencimiento = @PrepagoVencimiento, PrepagoReservasCantidad = @PrepagoReservasCantidad, IDCuentaCorrienteGrupo_Credito = @IDCuentaCorrienteGrupo_Credito, IDCuentaCorrienteGrupo_Debito = @IDCuentaCorrienteGrupo_Debito, Notas = @Notas, Activo = @Activo, FechaHoraModificacion = getdate(), IDUsuarioModificacion = @IDUsuario
					WHERE IDListaPrecio = @IDListaPrecio 

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
-- Created:	26/04/2014 18:03:40
-- Updated:	
-- Description: Elimina una Lista de Precios
-- =============================================
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'usp_ListaPrecio_Delete') AND type in (N'P', N'PC'))
	 DROP PROCEDURE usp_ListaPrecio_Delete
GO

CREATE PROCEDURE dbo.usp_ListaPrecio_Delete 
	@IDListaPrecio int AS
	
BEGIN
	SET NOCOUNT ON;

	BEGIN TRY
	
		BEGIN TRANSACTION

			DELETE
				FROM ListaPrecio
				WHERE IDListaPrecio = @IDListaPrecio 

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



------------------------------------------------------------------------------------------
-- LISTAPRECIODETALLE_DATA
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ListaPrecioDetalle_Data' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ListaPrecioDetalle_Data
GO

CREATE PROCEDURE dbo.sp_ListaPrecioDetalle_Data 
	@IDListaPrecio_FILTER int, 
	@OcupanteTipo_FILTER char(2),
	@IDRuta_FILTER char(20),
	@IDLugarGrupoOrigen_FILTER int,
	@IDLugarGrupoDestino_FILTER int AS

	SELECT ListaPrecioDetalle.IDListaPrecio, ListaPrecioDetalle.OcupanteTipo, ListaPrecioDetalle.IDRuta, ListaPrecioDetalle.IDLugarGrupoOrigen, ListaPrecioDetalle.IDLugarGrupoDestino, ListaPrecioDetalle.Importe, ListaPrecioDetalle.ImporteWeb, ListaPrecioDetalle.FechaHoraCreacion, IDUsuarioCreacion, FechaHoraModificacion, IDUsuarioModificacion
		FROM ListaPrecioDetalle
		WHERE ListaPrecioDetalle.IDListaPrecio = @IDListaPrecio_FILTER AND ListaPrecioDetalle.OcupanteTipo = @OcupanteTipo_FILTER AND ListaPrecioDetalle.IDRuta = @IDRuta_FILTER AND ListaPrecioDetalle.IDLugarGrupoOrigen = @IDLugarGrupoOrigen_FILTER AND ListaPrecioDetalle.IDLugarGrupoDestino = @IDLugarGrupoDestino_FILTER

GO



------------------------------------------------------------------------------------------
-- LISTAPRECIODETALLE_IMPORTE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ListaPrecioDetalle_Importe' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ListaPrecioDetalle_Importe
GO

CREATE PROCEDURE dbo.sp_ListaPrecioDetalle_Importe 
	@IDListaPrecio_FILTER int,
	@OcupanteTipo_FILTER char(2),
	@IDRuta_FILTER char(20), 
	@IDOrigen_FILTER int,
	@IDDestino_FILTER int AS

	SELECT ListaPrecioDetalle.Importe, ListaPrecioDetalle.ImporteWeb
		FROM (ListaPrecioDetalle INNER JOIN RutaDetalle AS RutaDetalleOrigen ON ListaPrecioDetalle.IDRuta = RutaDetalleOrigen.IDRuta AND ListaPrecioDetalle.IDLugarGrupoOrigen = RutaDetalleOrigen.IDLugarGrupo) INNER JOIN RutaDetalle AS RutaDetalleDestino ON ListaPrecioDetalle.IDRuta = RutaDetalleDestino.IDRuta AND ListaPrecioDetalle.IDLugarGrupoDestino = RutaDetalleDestino.IDLugarGrupo
		WHERE ListaPrecioDetalle.IDListaPrecio = @IDListaPrecio_FILTER AND ListaPrecioDetalle.OcupanteTipo = @OcupanteTipo_FILTER AND ListaPrecioDetalle.IDRuta = @IDRuta_FILTER AND RutaDetalleOrigen.IDLugar = @IDOrigen_FILTER AND RutaDetalleDestino.IDLugar = @IDDestino_FILTER

GO



------------------------------------------------------------------------------------------
-- LISTAPRECIODETALLE_DATAGRID_COMPLETE
------------------------------------------------------------------------------------------
IF EXISTS (SELECT name 
	   FROM   sysobjects 
	   WHERE  name = N'sp_ListaPrecioDetalle_DataGrid_Complete' 
	   AND 	  type = 'P')
    DROP PROCEDURE sp_ListaPrecioDetalle_DataGrid_Complete
GO

CREATE PROCEDURE dbo.sp_ListaPrecioDetalle_DataGrid_Complete
	@IDListaPrecio_FILTER int,
	@OcupanteTipo_FILTER char(2),
	@IDRuta_FILTER char(20) AS

	SELECT ListaPrecioDetalle_DataGrid.IDLugarGrupoOrigen, ListaPrecioDetalle_DataGrid.LugarGrupoOrigen, ListaPrecioDetalle_DataGrid.IDLugarGrupoDestino, ListaPrecioDetalle_DataGrid.LugarGrupoDestino, ListaPrecioDetalle_List.Importe, ListaPrecioDetalle_List.ImporteWeb
		FROM
			(SELECT LugarGrupoOrigen.IDLugarGrupo AS IDLugarGrupoOrigen, LugarGrupoOrigen.Nombre AS LugarGrupoOrigen, LugarGrupoDestino.IDLugarGrupo AS IDLugarGrupoDestino, LugarGrupoDestino.Nombre AS LugarGrupoDestino, ListaPrecioDetalle_DataGrid_Order_Origen.IndiceMaximo AS IndiceMaximoOrigen, ListaPrecioDetalle_DataGrid_Order_Destino.IndiceMaximo AS IndiceMaximoDestino
				FROM LugarGrupo AS LugarGrupoOrigen INNER JOIN
					(SELECT IDLugarGrupo, Max(Indice) AS IndiceMaximo
						FROM RutaDetalle
						WHERE IDRuta = @IDRuta_FILTER
						GROUP BY IDLugarGrupo)
					AS ListaPrecioDetalle_DataGrid_Order_Origen
					ON LugarGrupoOrigen.IDLugarGrupo = ListaPrecioDetalle_DataGrid_Order_Origen.IDLugarGrupo, 
					(SELECT IDLugarGrupo, Max(Indice) AS IndiceMaximo
						FROM RutaDetalle
						WHERE IDRuta = @IDRuta_FILTER
						GROUP BY IDLugarGrupo)
					AS ListaPrecioDetalle_DataGrid_Order_Destino
					INNER JOIN LugarGrupo AS LugarGrupoDestino ON ListaPrecioDetalle_DataGrid_Order_Destino.IDLugarGrupo = LugarGrupoDestino.IDLugarGrupo
				WHERE ListaPrecioDetalle_DataGrid_Order_Origen.IndiceMaximo < ListaPrecioDetalle_DataGrid_Order_Destino.IndiceMaximo)			
			AS ListaPrecioDetalle_DataGrid LEFT JOIN
			(SELECT ListaPrecioDetalle.IDListaPrecio, ListaPrecioDetalle.OcupanteTipo, ListaPrecioDetalle.IDRuta, ListaPrecioDetalle.IDLugarGrupoOrigen, ListaPrecioDetalle.IDLugarGrupoDestino, ListaPrecioDetalle.Importe, ListaPrecioDetalle.ImporteWeb, ListaPrecioDetalle.FechaHoraCreacion
				FROM ListaPrecioDetalle
				WHERE ListaPrecioDetalle.IDListaPrecio = @IDListaPrecio_FILTER AND ListaPrecioDetalle.OcupanteTipo = @OcupanteTipo_FILTER AND ListaPrecioDetalle.IDRuta = @IDRuta_FILTER)
			AS ListaPrecioDetalle_List
			ON ListaPrecioDetalle_DataGrid.IDLugarGrupoDestino = ListaPrecioDetalle_List.IDLugarGrupoDestino AND ListaPrecioDetalle_DataGrid.IDLugarGrupoOrigen = ListaPrecioDetalle_List.IDLugarGrupoOrigen
		WHERE (ListaPrecioDetalle_List.OcupanteTipo = @OcupanteTipo_FILTER AND ListaPrecioDetalle_List.IDListaPrecio = @IDListaPrecio_FILTER AND ListaPrecioDetalle_List.IDRuta = @IDRuta_FILTER) OR ListaPrecioDetalle_List.IDListaPrecio Is Null
		ORDER BY ListaPrecioDetalle_DataGrid.IndiceMaximoOrigen, ListaPrecioDetalle_DataGrid.IndiceMaximoDestino, ListaPrecioDetalle_List.OcupanteTipo

GO