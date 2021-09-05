Attribute VB_Name = "ModuloKubalSoft"
Option Explicit

Private Function GetReservasOptimizacion(ByVal FechaHora As Date, ByVal IDRuta As String) As Collection
    Dim cmdViajeDetalleAsientoList As ADODB.command
    Dim recViajeDetalleAsientoList As ADODB.Recordset
    Dim reservasOptimizacion As New Collection
    Dim reservaOptimizacion As LobosBus_Server_Optimizador.reservaOptimizacion
    Dim paradaActualIndice As Long
    Dim paradaActualIDLugarGrupo As Long

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set cmdViajeDetalleAsientoList = New ADODB.command
    Set cmdViajeDetalleAsientoList.ActiveConnection = pDatabase.Connection
    cmdViajeDetalleAsientoList.CommandText = "sp_ViajeDetalle_Asiento_List"
    cmdViajeDetalleAsientoList.CommandType = adCmdStoredProc
    cmdViajeDetalleAsientoList.Parameters.Append cmdViajeDetalleAsientoList.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , FechaHora)
    cmdViajeDetalleAsientoList.Parameters.Append cmdViajeDetalleAsientoList.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, IDRuta)
    cmdViajeDetalleAsientoList.Parameters.Append cmdViajeDetalleAsientoList.CreateParameter("OcupanteTipo_FILTER", adChar, adParamInput, 2, OCUPANTE_TIPO_PASAJERO)
    cmdViajeDetalleAsientoList.Parameters.Append cmdViajeDetalleAsientoList.CreateParameter("EstadoExclude_FILTER", adChar, adParamInput, 3, VIAJE_DETALLE_ESTADO_CANCELADO)
    Set recViajeDetalleAsientoList = New ADODB.Recordset
    recViajeDetalleAsientoList.Open cmdViajeDetalleAsientoList, , adOpenForwardOnly, adLockReadOnly
    Set cmdViajeDetalleAsientoList = Nothing
    
    With recViajeDetalleAsientoList
        Do While Not .EOF
            If paradaActualIndice < .Fields("IndiceDestino").value Then
                reservaOptimizacion = New LobosBus_Server_Optimizador.reservaOptimizacion
                reservaOptimizacion.idReserva = .Fields("IDViajeDetalle").value
                reservaOptimizacion.Asiento = .Fields("AsientoIdentificacion").value
                reservaOptimizacion.ciudadDesde = IIf(paradaActualIndice > .Fields("IndiceOrigen").value, paradaActualIDLugarGrupo, .Fields("OrigenIDLugarGrupo").value)
                reservaOptimizacion.desde = IIf(paradaActualIndice > .Fields("IndiceOrigen").value, paradaActualIndice, .Fields("IndiceOrigen").value)
                reservaOptimizacion.hasta = .Fields("IndiceDestino").value
                
                reservasOptimizacion.Add (reservaOptimizacion)
            End If
            
            .MoveNext
        Loop
    End With
    
    recViajeDetalleAsientoList.Close
    Set recViajeDetalleAsientoList = Nothing
    Set GetReservasOptimizacion = reservasOptimizacion
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Modules.ModuloKubalSoft.GetReservasOptimizacion", "Error al cargar las reservas del viaje para asignar los asientos." & vbCr & vbCr & "Fecha/Hora: " & Format(FechaHora, "Short Date") & " " & Format(FechaHora, "Short Time") & vbCr & "IDRuta: " & IDRuta
End Function

Public Function GetAsientosPermitidos(ByVal IDViajeDetalle As Long) As Collection
    Dim reservasOptimizacion As Collection
    
    'reservasOptimizacion = GetReservasOptimizacion()
End Function

Public Function GetDisponibilidad(ByVal indiceDesde As Long, ByVal indiceHasta As Long)
    Dim listaDeReservaOptimizacion As Variant
    
    'listaDeReservaOptimizacion = new LobosBus_Server_Optimizador.ReservaOptimizacion()
End Function
