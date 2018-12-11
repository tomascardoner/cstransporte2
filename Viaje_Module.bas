Attribute VB_Name = "Viaje_Module"
Option Explicit

Public Sub Viaje_EstadoVencido_List()
    Dim Viaje As Viaje
    Dim ViajesVencidos As Long
    Dim Reporte As Reporte
    
    Set Viaje = New Viaje
    ViajesVencidos = Viaje.EstadoVencidoCheck
    Set Viaje = Nothing
    
    If ViajesVencidos > 0 Then
        If MsgBox(IIf(ViajesVencidos = 1, "Hay un Viaje con el Estado vencido.", "Hay " & ViajesVencidos & " Viajes con el Estado vencido.") & vbCr & vbCr & "¿Desea ver el Reporte?", vbExclamation + vbYesNo, App.Title) = vbYes Then
            If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Viaje_EstadoVencido_Listado") Then
                Set Reporte = New Reporte
                Reporte.IDReporte = "Viaje_EstadoVencido_Listado"
                If Reporte.Load() Then
                    If Reporte.OpenReport() Then
                        Reporte.PrintReport True
                    End If
                End If
                Set Reporte = Nothing
            End If
        End If
    End If
End Sub

Public Sub ViajeDetalle_ShowViajeVuelta(ByRef ViajeDetalle As ViajeDetalle, ByVal MessageText As String)
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    Dim Persona As Persona
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set recData = New ADODB.Recordset
    With recData
        Set cmdData = New ADODB.command
        Set cmdData.ActiveConnection = pDatabase.Connection
        cmdData.CommandType = adCmdStoredProc
        cmdData.CommandText = "sp_ViajeDetalle_BuscaVuelta"
        cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , ViajeDetalle.FechaHora)
        cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, ViajeDetalle.IDRuta)
        cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , ViajeDetalle.IDPersona)
        cmdData.Parameters.Append cmdData.CreateParameter("IDRutaEspecial_FILTER", adChar, adParamInput, 20, pParametro.Ruta_ID_Otra)
        Set recData = New ADODB.Recordset
        recData.Open cmdData, , adOpenStatic, adLockReadOnly
        
        If Not recData.EOF Then
            recData.MoveLast
            recData.MoveFirst
            MessageText = Replace(MessageText, "%1", recData.RecordCount) & vbCr
            Set Persona = New Persona
            Persona.IDPersona = ViajeDetalle.IDPersona
            If Persona.Load() Then
                MessageText = MessageText & vbCr & vbCr & "Pasajero: " & Persona.ApellidoNombre
            End If
            Set Persona = Nothing
            Do While Not recData.EOF
                MessageText = MessageText & vbCr & vbCr & "Hora: " & Format(recData("FechaHora").Value, "Short Time")
                MessageText = MessageText & vbCr & "Ruta: " & RTrim(recData("IDRuta").Value)
                recData.MoveNext
            Loop
            Screen.MousePointer = vbDefault
            MsgBox MessageText, vbExclamation + vbOKOnly, App.Title
        End If
        recData.Close
        Set recData = Nothing
        Set cmdData = Nothing
    End With
    Screen.MousePointer = vbDefault
    
ErrorHandler:
End Sub
