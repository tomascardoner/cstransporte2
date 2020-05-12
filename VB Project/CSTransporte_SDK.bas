Attribute VB_Name = "CSTransporte_SDK"
Option Explicit

Public Function EntidadTipo_GetNombre(ByVal EntidadTipo As String) As String
    Select Case EntidadTipo
        Case ENTIDAD_TIPO_PERSONA_CLIENTE
            EntidadTipo_GetNombre = ENTIDAD_TIPO_PERSONA_CLIENTE_NOMBRE
        Case ENTIDAD_TIPO_PERSONA_CONDUCTOR
            EntidadTipo_GetNombre = ENTIDAD_TIPO_PERSONA_CONDUCTOR_NOMBRE
        Case ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO
            EntidadTipo_GetNombre = ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO_NOMBRE
        Case ENTIDAD_TIPO_VEHICULO
            EntidadTipo_GetNombre = ENTIDAD_TIPO_VEHICULO_NOMBRE
        Case ENTIDAD_TIPO_VIAJE
            EntidadTipo_GetNombre = ENTIDAD_TIPO_VIAJE_NOMBRE
        Case ENTIDAD_TIPO_VIAJE_DETALLE
            EntidadTipo_GetNombre = ENTIDAD_TIPO_VIAJE_DETALLE_NOMBRE
        Case ENTIDAD_TIPO_CONTACTO
            EntidadTipo_GetNombre = ENTIDAD_TIPO_CONTACTO_NOMBRE
        Case Else
            EntidadTipo_GetNombre = ""
    End Select
End Function

Public Function AbrirViajeYDetalle(ByVal FechaHora As Date, ByVal IDRuta As String)
    Dim Viaje As Viaje
    
    If pCPermiso.GotPermission(PERMISO_VIAJE, False) Then
        If Not CSM_Forms.IsLoaded("frmViaje") Then
            Load frmViaje
        End If
        With frmViaje
            .mLoading = True
            .cboDiaSemana.ListIndex = 0
            .cboFecha.ListIndex = 1
            .dtpFechaDesde.Value = FechaHora
            .cboRuta.ListIndex = 0
            .mLoading = False
            .FillListView FechaHora, IDRuta
        End With
        frmViaje.Show
        If frmViaje.WindowState = vbMinimized Then
            frmViaje.WindowState = vbNormal
        End If
        frmViaje.SetFocus
        If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE, False) Then
            Set Viaje = New Viaje
            Viaje.FechaHora = FechaHora
            Viaje.IDRuta = IDRuta
            If Viaje.Load() Then
                frmViajeDetalle.LoadDataAndShow Viaje
            End If
            Set Viaje = Nothing
        End If
    End If
End Function

Public Sub BuscarYAbrirViajeActualPorVehiculo(ByVal IDVehiculo As Long, ByVal VehiculoNombre As String)
    Dim recViaje As ADODB.Recordset
    Dim ListItem As MSComctlLib.ListItem
    Dim Viaje As Viaje
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set Viaje = New Viaje
    
    Set recViaje = New ADODB.Recordset
    recViaje.Source = "SELECT Viaje.FechaHora, Viaje.IDRuta, Viaje.Estado FROM Viaje INNER JOIN Ruta ON Viaje.IDRuta = Ruta.IDRuta WHERE Viaje.IDVehiculo = " & IDVehiculo & " AND Viaje.FechaHora >= DateAdd(minute, -" & pParametro.Viaje_Actual_Rango_Minutos & ", getdate()) AND Viaje.FechaHora <= DateAdd(minute, " & pParametro.Viaje_Actual_Rango_Minutos & " + ISNULL(Viaje.Duracion, 0), getdate()) ORDER BY Viaje.FechaHora"
    recViaje.Open , pDatabase.Connection, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not recViaje.EOF Then
        recViaje.MoveLast
        recViaje.MoveFirst
        If recViaje.RecordCount = 1 Then
            AbrirViajeYDetalle recViaje("FechaHora").Value, recViaje("IDRuta").Value
        Else
            frmViajeVehiculoSelect.LoadDataAndShow Now
            frmViajeVehiculoSelect.txtFechaHora.Text = Format(Now, "Short Date") & " " & Format(Now, "Short Time")
            frmViajeVehiculoSelect.txtVehiculo.Text = VehiculoNombre
            frmViajeVehiculoSelect.lvwData.ListItems.Clear
            With recViaje
                Do While Not .EOF
                    Set ListItem = frmViajeVehiculoSelect.lvwData.ListItems.Add(, KEY_STRINGER & .Fields("FechaHora").Value & KEY_DELIMITER & .Fields("IDRuta").Value, WeekdayName(Weekday(.Fields("FechaHora").Value)))
                    ListItem.SubItems(1) = Format(.Fields("FechaHora").Value, "Short Date")
                    ListItem.SubItems(2) = Format(.Fields("FechaHora").Value, "Short Time")
                    ListItem.SubItems(3) = .Fields("IDRuta").Value
                    Viaje.Estado = .Fields("Estado").Value
                    ListItem.SubItems(4) = Viaje.Estado_ToString
                    ListItem.ForeColor = Viaje.Estado_ToColor
                    ListItem.Bold = Viaje.Estado_ToBold
                    .MoveNext
                Loop
                If frmViajeVehiculoSelect.WindowState = vbMinimized Then
                    frmViajeVehiculoSelect.WindowState = vbNormal
                End If
                frmViajeVehiculoSelect.Show
                frmViajeVehiculoSelect.SetFocus
            End With
        End If
    Else
        Screen.MousePointer = vbDefault
        MsgBox "No hay Viajes dentro de los últimos o siguientes " & pParametro.Viaje_Actual_Rango_Minutos & " minutos.", vbExclamation, App.Title
    End If
    
    recViaje.Close
    Set recViaje = Nothing
    
    Set Viaje = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Modules.CSTransporte_SDK.BuscarYAbrirViajeActualPorVehiculo", "Error al Obtener el Viaje Actual."
End Sub

Public Sub SetLastPersona(ByVal IDPersona As Long, Optional ByVal ApellidoNombre As String = "")
    Dim PersonaIndex As Long
    Dim SaveCallerID As Boolean
    Dim Persona As Persona
    
    With frmMDI.cboPersona
        For PersonaIndex = 0 To .ListCount - 1
            If .ItemData(PersonaIndex) = IDPersona Then
                SaveCallerID = True
                Exit For
            End If
        Next PersonaIndex
        .Clear
        .Tag = 0
        If IDPersona > 0 Then
            If ApellidoNombre = "" Then
                Set Persona = New Persona
                Persona.IDPersona = IDPersona
                Persona.Load
                ApellidoNombre = Persona.ApellidoNombre
                Set Persona = Nothing
            End If
            .AddItem ApellidoNombre
            .ItemData(.NewIndex) = IDPersona
            .ListIndex = 0
        End If
    End With
    If Not SaveCallerID Then
        frmMDI.txtCallerIDTipo.Text = ""
        frmMDI.txtCallerID.Text = ""
    End If
End Sub

Public Function LogAccionAdd(ByVal EntidadTipo As String, ByVal Descripcion As String) As Boolean
    Dim LogAccion As LogAccion
    
    If pParametro.LogAccion_Enabled Then
        Set LogAccion = New LogAccion
        LogAccion.EntidadTipo = EntidadTipo
        LogAccion.Descripcion = Descripcion
        LogAccionAdd = LogAccion.Update()
        Set LogAccion = Nothing
    End If
End Function

Public Function TextReplaceSystemVariables(ByVal Value As String, ByRef Persona As Persona) As String
    Value = Replace(Value, "|@NL|", vbCrLf)
    Value = Replace(Value, "|@CompanyName|", pParametro.CompanyName)
    Value = Replace(Value, "|@WebSite_HyperLink_Main|", "http://" & pParametro.WebSite_HyperLink_Main)
    Value = Replace(Value, "|@WebSite_HyperLink_Login|", "http://" & pParametro.WebSite_HyperLink_Login)
    Value = Replace(Value, "|@WebSite_HyperLink_ChangePassword|", "http://" & pParametro.WebSite_HyperLink_ChangePassword)
    If Not Persona Is Nothing Then
        Value = Replace(Value, "|@DocumentoTipo|", Trim(Persona.IDDocumentoTipo))
        Value = Replace(Value, "|@DocumentoNumero|", Persona.DocumentoNumero)
        Value = Replace(Value, "|@Password|", Persona.Password)
    End If
    
    TextReplaceSystemVariables = Value
End Function
