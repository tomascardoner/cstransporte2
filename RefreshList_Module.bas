Attribute VB_Name = "RefreshList_Module"
Option Explicit

Public Const MODULE_USUARIO_GRUPO = "UG"
Public Const MODULE_USUARIO_GRUPO_PERMISO = "UP"
Public Const MODULE_USUARIO = "US"
Public Const MODULE_LUGAR = "LU"
Public Const MODULE_LUGAR_GRUPO = "LG"
Public Const MODULE_RUTA = "RU"
Public Const MODULE_RUTA_DETALLE = "RD"
Public Const MODULE_RUTA_LUGARGRUPO = "RL"
Public Const MODULE_LISTA_PRECIO = "LP"
Public Const MODULE_LISTA_PRECIO_DETALLE = "LD"
Public Const MODULE_VEHICULO = "VE"
Public Const MODULE_VEHICULO_UTILIZACION = "VU"
Public Const MODULE_HORARIO = "HO"
Public Const MODULE_PERSONA = "PE"
Public Const MODULE_PERSONA_HORARIO = "PH"
Public Const MODULE_PERSONA_RUTA = "PR"
Public Const MODULE_PERSONA_PREPAGO = "PP"
Public Const MODULE_PERSONA_RESPUESTA = "PS"
Public Const MODULE_CONDUCTOR_RUTA = "CU"
Public Const MODULE_VIAJE = "VI"
Public Const MODULE_VIAJE_DETALLE = "VD"
Public Const MODULE_CUENTA_CORRIENTE = "CC"
Public Const MODULE_CUENTA_CORRIENTE_CAJA = "CA"
Public Const MODULE_CUENTA_CORRIENTE_GRUPO = "CG"
Public Const MODULE_MEDIOPAGO = "MP"
Public Const MODULE_FERIADO = "FE"
Public Const MODULE_VEHICULO_MANTENIMIENTO = "VM"
Public Const MODULE_VEHICULO_MANTENIMIENTO_GRUPO = "VG"
Public Const MODULE_VEHICULO_MANTENIMIENTO_ACCION = "VC"
Public Const MODULE_DOCUMENTO_TIPO = "DT"
Public Const MODULE_TELEFONO_TIPO = "TT"
Public Const MODULE_PERSONA_ALARMA = "PL"
Public Const MODULE_PERSONA_ALARMA_GRUPO = "PG"
Public Const MODULE_ALARMA = "AL"
Public Const MODULE_CONTACTO = "CO"
Public Const MODULE_CONTACTO_GRUPO = "CR"
Public Const MODULE_REGISTROLLAMADA = "RM"
Public Const MODULE_FRANCO = "FR"
Public Const MODULE_MENSAJE = "ME"
Public Const MODULE_MESSENGER = "MS"

Public Const MODULE_PERSONAL = "PO"

Public Sub RefreshList_Fastest_CheckForRefreshs()
    On Error GoTo ErrorHandler
    
    precSemaforoGeneral.Requery
    precSemaforoGeneral.Filter = ""
    If Not (precSemaforoGeneral.BOF And precSemaforoGeneral.EOF) Then
        precSemaforoGeneral.MoveFirst
        Do While Not precSemaforoGeneral.EOF
            If pCSemaforoGeneral(precSemaforoGeneral("IDSemaforo").Value).ValorTimer = -1 Then
                pCSemaforoGeneral(precSemaforoGeneral("IDSemaforo").Value).ValorTimer = precSemaforoGeneral("ValorTimer").Value
            Else
                If precSemaforoGeneral("ValorTimer").Value <> pCSemaforoGeneral(precSemaforoGeneral("IDSemaforo").Value).ValorTimer Then
                    pCSemaforoGeneral(precSemaforoGeneral("IDSemaforo").Value).ValorTimer = precSemaforoGeneral("ValorTimer").Value
                    Select Case precSemaforoGeneral("IDSemaforo").Value
                        Case MODULE_VIAJE
                            RefreshList_RefreshViaje DATE_TIME_FIELD_NULL_VALUE, "", False
                        Case MODULE_VIAJE_DETALLE
                            RefreshList_RefreshViajeDetalle DATE_TIME_FIELD_NULL_VALUE, "", 0, True, False
                        Case MODULE_PERSONAL
                            If pParametro.Personal_Status_Global Then
                                RefreshList_RefreshPersonal False
                            End If
                        Case MODULE_MESSENGER
                            If pMessengerEnabled Then
                                Call Messenger
                            End If
                    End Select
                End If
            End If
            precSemaforoGeneral.MoveNext
        Loop
    End If
    Exit Sub
    
ErrorHandler:
End Sub

Public Sub RefreshList_Slowest_CheckForRefreshs()
    On Error GoTo ErrorHandler
    
    precSemaforoGeneral.Requery
    precSemaforoGeneral.Filter = ""
    If Not (precSemaforoGeneral.BOF And precSemaforoGeneral.EOF) Then
        precSemaforoGeneral.MoveFirst
        Do While Not precSemaforoGeneral.EOF
            If pCSemaforoGeneral(precSemaforoGeneral("IDSemaforo").Value).ValorTimer = -1 Then
                pCSemaforoGeneral(precSemaforoGeneral("IDSemaforo").Value).ValorTimer = precSemaforoGeneral("ValorTimer").Value
            Else
                If precSemaforoGeneral("ValorTimer").Value <> pCSemaforoGeneral(precSemaforoGeneral("IDSemaforo").Value).ValorTimer Then
                    pCSemaforoGeneral(precSemaforoGeneral("IDSemaforo").Value).ValorTimer = precSemaforoGeneral("ValorTimer").Value
                    Select Case precSemaforoGeneral("IDSemaforo").Value
                        Case MODULE_USUARIO
                            RefreshList_RefreshUsuario "", False
                        Case MODULE_USUARIO_GRUPO
                            RefreshList_RefreshUsuarioGrupo 0, False
                        Case MODULE_USUARIO_GRUPO_PERMISO
                            RefreshList_RefreshUsuarioGrupoPermiso 0, 0, True, False
                        Case MODULE_LUGAR
                            RefreshList_RefreshLugar 0, False
                        Case MODULE_LUGAR_GRUPO
                            RefreshList_RefreshLugarGrupo 0, False
                        Case MODULE_RUTA
                            RefreshList_RefreshRuta "", False
                        Case MODULE_RUTA_DETALLE
                            RefreshList_RefreshRutaDetalle "", 0, True, False
                        Case MODULE_LISTA_PRECIO
                            RefreshList_RefreshListaPrecio 0, False
                        Case MODULE_LISTA_PRECIO_DETALLE
                            
                        Case MODULE_VEHICULO
                            RefreshList_RefreshVehiculo 0, False
                        Case MODULE_VEHICULO_UTILIZACION
                            RefreshList_RefreshVehiculoUtilizacion False
                        Case MODULE_HORARIO
                            RefreshList_RefreshHorario 0, DATE_TIME_FIELD_NULL_VALUE, "", False
                        Case MODULE_PERSONA
                            RefreshList_RefreshPersona 0, False
                        Case MODULE_PERSONA_HORARIO
                            RefreshList_RefreshPersonaHorario 0, 0, DATE_TIME_FIELD_NULL_VALUE, "", True, False
                        Case MODULE_PERSONA_RUTA
                            RefreshList_RefreshPersonaRuta 0, "", True, False
                        Case MODULE_PERSONA_PREPAGO
                            RefreshList_RefreshPersonaPrepago 0, "", DATE_TIME_FIELD_NULL_VALUE, True, False
                        Case MODULE_PERSONA_RESPUESTA
                            RefreshList_RefreshPersonaRespuesta 0, DATE_TIME_FIELD_NULL_VALUE, False
                        Case MODULE_CONDUCTOR_RUTA
                            RefreshList_RefreshConductorRuta 0, "", True, False
                        Case MODULE_CUENTA_CORRIENTE
                            'RefreshList_RefreshCuentaCorriente 0, False
                        Case MODULE_CUENTA_CORRIENTE_CAJA
                            RefreshList_RefreshCuentaCorrienteCaja 0, False
                        Case MODULE_CUENTA_CORRIENTE_GRUPO
                            RefreshList_RefreshCuentaCorrienteGrupo 0, False
                        Case MODULE_MEDIOPAGO
                            RefreshList_RefreshMedioPago 0, False
                        Case MODULE_FERIADO
                            RefreshList_RefreshFeriado DATE_TIME_FIELD_NULL_VALUE, False
                        Case MODULE_VEHICULO_MANTENIMIENTO
                            RefreshList_RefreshVehiculoMantenimiento 0, 0, False
                        Case MODULE_VEHICULO_MANTENIMIENTO_GRUPO
                            RefreshList_RefreshVehiculoMantenimientoGrupo 0, False
                        Case MODULE_VEHICULO_MANTENIMIENTO_ACCION
                            RefreshList_RefreshVehiculoMantenimientoAccion 0, False
                        Case MODULE_DOCUMENTO_TIPO
                            RefreshList_RefreshDocumentoTipo "", False
                        Case MODULE_TELEFONO_TIPO
                            RefreshList_RefreshTelefonoTipo 0, False
                        Case MODULE_PERSONA_ALARMA
                            RefreshList_RefreshPersonaAlarma 0, 0, False
                        Case MODULE_PERSONA_ALARMA_GRUPO
                            RefreshList_RefreshPersonaAlarmaGrupo 0, False
                        Case MODULE_ALARMA
                            RefreshList_RefreshAlarma 0, False
                        Case MODULE_CONTACTO
                            RefreshList_RefreshContacto 0, False
                        Case MODULE_CONTACTO_GRUPO
                            RefreshList_RefreshContactoGrupo 0, False
                        Case MODULE_REGISTROLLAMADA
                            RefreshList_RefreshRegistroLlamada 0, False
                        Case MODULE_FRANCO
                            RefreshList_Module.Franco DATE_TIME_FIELD_NULL_VALUE, 0, False
                        Case MODULE_MENSAJE
                            RefreshList_Module.Mensaje 0, False
                    End Select
                End If
            End If
            precSemaforoGeneral.MoveNext
        Loop
    End If
    Exit Sub
    
ErrorHandler:
End Sub

Public Sub RefreshList_CheckForPersonal()
    On Error GoTo ErrorHandler
    
    precSemaforoGeneral.Resync
    precSemaforoGeneral.Filter = "IDSemaforo = '" & MODULE_PERSONAL & "'"
    If Not (precSemaforoGeneral.BOF And precSemaforoGeneral.EOF) Then
        If precSemaforoGeneral("ValorTimer").Value <> pCSemaforoGeneral(precSemaforoGeneral("IDSemaforo").Value).ValorTimer Then
            pCSemaforoGeneral(precSemaforoGeneral("IDSemaforo").Value).ValorTimer = precSemaforoGeneral("ValorTimer").Value
            RefreshList_RefreshPersonal False
        End If
    End If
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub RefreshList_UpdateValue(ByVal IDSemaforo As String, Optional ByVal ExtraTexto As String = "", Optional ByVal ExtraNumero As Long = 0, Optional ByVal ExtraFecha As Date = DATE_TIME_FIELD_NULL_VALUE, Optional ByVal ExtraSiNo As Integer = -2)
    
RESTART:
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    precSemaforoGeneral.Filter = "IDSemaforo = '" & IDSemaforo & "'"
    If precSemaforoGeneral.EOF Then
        precSemaforoGeneral.AddNew
        precSemaforoGeneral("IDSemaforo").Value = IDSemaforo
    End If
    pCSemaforoGeneral(IDSemaforo).ValorTimer = timeGetTime()
    precSemaforoGeneral("ValorTimer").Value = pCSemaforoGeneral(IDSemaforo).ValorTimer
    precSemaforoGeneral("ExtraTexto").Value = IIf(ExtraTexto = "", Null, ExtraTexto)
    precSemaforoGeneral("ExtraNumero").Value = IIf(ExtraNumero = 0, Null, ExtraNumero)
    precSemaforoGeneral("ExtraFecha").Value = IIf(ExtraFecha = DATE_TIME_FIELD_NULL_VALUE, Null, ExtraFecha)
    precSemaforoGeneral("ExtraSiNo").Value = IIf(ExtraSiNo = -2, Null, Abs(ExtraSiNo))
    precSemaforoGeneral.Update
    Exit Sub
    
ErrorHandler:
    If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_VALUE_CHANGED_SINCE_LAST_READ Then
        precSemaforoGeneral.CancelUpdate
        precSemaforoGeneral.Requery
        Resume RESTART
    End If
End Sub

Public Sub RefreshList_RefreshAll()
    RefreshList_RefreshUsuario "", False
    RefreshList_RefreshUsuarioGrupo 0, False
    RefreshList_RefreshUsuarioGrupoPermiso 0, 0, True, False
    RefreshList_RefreshLugar 0, False
    RefreshList_RefreshLugarGrupo 0, False
    RefreshList_RefreshRuta "", False
    RefreshList_RefreshRutaDetalle "", False
    RefreshList_RefreshListaPrecio 0, False
    RefreshList_RefreshVehiculo 0, False
    RefreshList_RefreshHorario 0, DATE_TIME_FIELD_NULL_VALUE, "", False
    RefreshList_RefreshPersona 0, False
    RefreshList_RefreshPersonaHorario 0, 0, DATE_TIME_FIELD_NULL_VALUE, "", True, False
    RefreshList_RefreshPersonaRuta 0, "", True, False
    RefreshList_RefreshPersonaPrepago 0, "", DATE_TIME_FIELD_NULL_VALUE, True, False
    RefreshList_RefreshPersonaRespuesta 0, DATE_TIME_FIELD_NULL_VALUE, False
    RefreshList_RefreshConductorRuta 0, "", False
    RefreshList_RefreshViaje DATE_TIME_FIELD_NULL_VALUE, "", False
    RefreshList_RefreshViajeDetalle DATE_TIME_FIELD_NULL_VALUE, "", 0, True, False
    RefreshList_RefreshCuentaCorriente 0, False
    RefreshList_RefreshCuentaCorrienteCaja 0, False
    RefreshList_RefreshCuentaCorrienteGrupo 0, False
    RefreshList_RefreshMedioPago 0, False
    RefreshList_RefreshFeriado DATE_TIME_FIELD_NULL_VALUE, False
    RefreshList_RefreshVehiculoMantenimiento 0, 0, False
    RefreshList_RefreshVehiculoMantenimientoGrupo 0, False
    RefreshList_RefreshVehiculoMantenimientoAccion 0, False
    RefreshList_RefreshDocumentoTipo "", False
    RefreshList_RefreshTelefonoTipo 0, False
    RefreshList_RefreshPersonaAlarma 0, 0, False
    RefreshList_RefreshPersonaAlarmaGrupo 0, False
    RefreshList_RefreshAlarma 0, False
    RefreshList_RefreshContacto 0, False
    RefreshList_RefreshContactoGrupo 0, False
    RefreshList_RefreshRegistroLlamada 0, False
    Franco DATE_TIME_FIELD_NULL_VALUE, 0, False
    Mensaje 0, False
    'Messenger
End Sub

Public Sub RefreshList_RefreshPersonal(Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_PERSONAL, , , , pPersonal
    Else
        If IsNull(precSemaforoGeneral("ExtraSiNo").Value) Then
            'pPersonal = pParametro.Personal_Status_OnStartup
        Else
            pPersonal = precSemaforoGeneral("ExtraSiNo").Value
        End If
    End If
    
    frmMDI.stbMain.Panels("PERSONAL").Text = IIf(pPersonal, "ON", "OFF")

    RefreshList_RefreshHorario 0, Date, "", False
    RefreshList_RefreshViaje Date, "", False
    RefreshList_RefreshViajeDetalle Date, "", 0, True, False
    RefreshList_RefreshCuentaCorriente 0, False
    RefreshList_RefreshPersonaHorario 0, 0, Date, "", True, False
    
    If CSM_Forms.IsLoaded("frmReporte") Then
        frmReporte.FillListViewReport
    End If
    
    Call CSM_Registry.SetValue_ToApplication_CurrentUser("Interface", "Personal", pPersonal)
End Sub

Public Sub RefreshList_RefreshUsuarioGrupo(ByVal IDUsuarioGrupo As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_USUARIO_GRUPO
    End If
    If CSM_Forms.IsLoaded("frmUsuarioGrupo") Then
        frmUsuarioGrupo.FillListView IDUsuarioGrupo
    End If
    If CSM_Forms.IsLoaded("frmUsuario") Then
        frmUsuario.FillComboBoxUsuarioGrupo
    End If
    If CSM_Forms.IsLoaded("frmUsuarioPropiedad") Then
        frmUsuarioPropiedad.FillComboBoxUsuarioGrupo
    End If
    If CSM_Forms.IsLoaded("frmMensajePropiedad") Then
        frmMensajePropiedad.FillComboBoxUsuarioGrupo
    End If
End Sub

Public Sub RefreshList_RefreshUsuarioGrupoPermiso(ByVal IDUsuarioGrupo As Long, ByVal IDPermiso As String, Optional ByVal Force As Boolean = False, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_USUARIO_GRUPO_PERMISO
    End If
    If CSM_Forms.IsLoaded("frmUsuarioGrupoPermiso") Then
        If Force Then
            frmUsuarioGrupoPermiso.ForceRefresh
        Else
            frmUsuarioGrupoPermiso.FillListView IDUsuarioGrupo, IDPermiso
        End If
    End If
End Sub

Public Sub RefreshList_RefreshUsuario(ByVal IDUsuario As String, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_USUARIO
    End If
    If CSM_Forms.IsLoaded("frmUsuario") Then
        frmUsuario.FillListView IDUsuario
    End If
End Sub

Public Sub RefreshList_RefreshLugar(ByVal IDLugar As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_LUGAR
    End If
    If CSM_Forms.IsLoaded("frmLugar") Then
        frmLugar.FillListView IDLugar
    End If
    If CSM_Forms.IsLoaded("frmRutaPropiedad") Then
        frmRutaPropiedad.FillComboBoxLugar
    End If
    If CSM_Forms.IsLoaded("frmRutaDetallePropiedad") Then
        frmRutaDetallePropiedad.FillComboBoxLugar
    End If
    If CSM_Forms.IsLoaded("frmPersonaRutaPropiedad") Then
        frmPersonaRutaPropiedad.FillComboBoxLugar
    End If
    If CSM_Forms.IsLoaded("frmViajeDetallePropiedad") Then
        frmViajeDetallePropiedad.FillComboBoxLugar
    End If
    If CSM_Forms.IsLoaded("frmRutaLugarGrupoPropiedad") Then
        frmRutaLugarGrupoPropiedad.FillComboBoxLugar
    End If
End Sub

Public Sub RefreshList_RefreshLugarGrupo(ByVal IDLugarGrupo As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_LUGAR_GRUPO
    End If
    If CSM_Forms.IsLoaded("frmLugarGrupo") Then
        frmLugarGrupo.FillListView IDLugarGrupo
    End If
    If CSM_Forms.IsLoaded("frmRutaDetallePropiedad") Then
        frmRutaDetallePropiedad.FillComboBoxLugarGrupo
    End If
    If CSM_Forms.IsLoaded("frmRutaLugarGrupoPropiedad") Then
        frmRutaLugarGrupoPropiedad.FillComboBoxLugarGrupo
    End If
End Sub

Public Sub RefreshList_RefreshRuta(ByVal IDRuta As String, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_RUTA
    End If
    If CSM_Forms.IsLoaded("frmRuta") Then
        frmRuta.FillListView IDRuta
    End If
    If CSM_Forms.IsLoaded("frmRutaDetalle") Then
        frmRutaDetalle.FillComboBoxRuta
    End If
    If CSM_Forms.IsLoaded("frmRutaDetallePropiedad") Then
        frmRutaDetallePropiedad.FillComboBoxRuta
    End If
    If CSM_Forms.IsLoaded("frmRutaLugarGrupoPropiedad") Then
        frmRutaLugarGrupoPropiedad.FillComboBoxRuta
    End If
    If CSM_Forms.IsLoaded("frmHorario") Then
        frmHorario.FillComboBoxRuta
    End If
    If CSM_Forms.IsLoaded("frmHorarioPropiedad") Then
        frmHorarioPropiedad.FillComboBoxRuta
    End If
    If CSM_Forms.IsLoaded("frmViaje") Then
        frmViaje.FillComboBoxRuta
    End If
    If CSM_Forms.IsLoaded("frmViajePropiedad") Then
        frmViajePropiedad.FillComboBoxRuta
    End If
    If CSM_Forms.IsLoaded("frmViajeDetallePropiedad") Then
        frmViajeDetallePropiedad.FillComboBoxRuta
    End If
    If CSM_Forms.IsLoaded("frmPersonaHorario") Then
        frmPersonaHorario.FillComboBoxRuta
    End If
    If CSM_Forms.IsLoaded("frmPersonaRutaPropiedad") Then
        frmPersonaRutaPropiedad.FillComboBoxRuta
    End If
    If CSM_Forms.IsLoaded("frmComision") Then
        frmComision.FillComboBoxRuta
    End If
    If CSM_Forms.IsLoaded("frmConductorRutaPropiedad") Then
        frmComision.FillComboBoxRuta
    End If
End Sub

Public Sub RefreshList_RefreshRutaDetalle(ByVal IDRuta As String, ByVal IDLugar As Long, Optional ByVal Force As Boolean = False, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_RUTA_DETALLE
    End If
    If CSM_Forms.IsLoaded("frmRutaDetalle") Then
        If Force Then
            frmRutaDetalle.ForceRefresh
        Else
            frmRutaDetalle.FillListView IDRuta, IDLugar
        End If
    End If
End Sub

Public Sub RefreshList_RefreshRutaLugarGrupo(ByVal IDRuta As String, ByVal IDLugarGrupo As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_RUTA_LUGARGRUPO
    End If
    If CSM_Forms.IsLoaded("frmRutaLugarGrupo") Then
        frmRutaLugarGrupo.FillListView IDRuta, IDLugarGrupo
    End If
End Sub

Public Sub RefreshList_RefreshListaPrecio(ByVal IDListaPrecio As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_LISTA_PRECIO
    End If
    If CSM_Forms.IsLoaded("frmListaPrecio") Then
        frmListaPrecio.FillListView IDListaPrecio
    End If
    If CSM_Forms.IsLoaded("frmViajeDetallePropiedad") Then
        frmViajeDetallePropiedad.FillComboBoxListaPrecio
    End If
    If CSM_Forms.IsLoaded("frmPersonaRutaPropiedad") Then
        frmPersonaRutaPropiedad.FillComboBoxListaPrecio
    End If
    If CSM_Forms.IsLoaded("frmPersonaPrepagoPropiedad") Then
        frmPersonaPrepagoPropiedad.FillComboBoxListaPrecio
    End If
    If CSM_Forms.IsLoaded("frmComision") Then
        frmComision.FillComboBoxListaPrecio
    End If
End Sub

Public Sub RefreshList_RefreshVehiculo(ByVal IDVehiculo As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_VEHICULO
    End If
    If CSM_Forms.IsLoaded("frmVehiculo") Then
        frmVehiculo.FillListView IDVehiculo
    End If
    If CSM_Forms.IsLoaded("frmHorarioPropiedad") Then
        frmHorarioPropiedad.FillComboBoxVehiculo
    End If
    If CSM_Forms.IsLoaded("frmHorarioPropiedadMultiple") Then
        frmHorarioPropiedadMultiple.FillComboBoxVehiculo
    End If
    If CSM_Forms.IsLoaded("frmViajePropiedad") Then
        frmViajePropiedad.FillComboBoxVehiculo
    End If
    If CSM_Forms.IsLoaded("frmViajePropiedadMultiple") Then
        frmViajePropiedadMultiple.FillComboBoxVehiculo
    End If
    If CSM_Forms.IsLoaded("frmVehiculoMantenimientoPropiedad") Then
        frmVehiculoMantenimientoPropiedad.FillComboBoxVehiculo
    End If
    If CSM_Forms.IsLoaded("frmVehiculoMantenimientoAccion") Then
        frmVehiculoMantenimientoAccion.FillComboBoxVehiculo
    End If
    If CSM_Forms.IsLoaded("frmVehiculoMantenimientoAccionPropiedad") Then
        frmVehiculoMantenimientoAccionPropiedad.FillComboBoxVehiculo
    End If
End Sub

Public Sub RefreshList_RefreshVehiculoUtilizacion(Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_VEHICULO_UTILIZACION
    End If
    If CSM_Forms.IsLoaded("frmVehiculoUtilizacion") Then
        frmVehiculoUtilizacion.FillGrid
    End If
End Sub

Public Sub RefreshList_RefreshHorario(ByVal DiaSemana As Byte, ByVal Hora As Date, ByVal IDRuta As String, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_HORARIO
    End If
    If CSM_Forms.IsLoaded("frmHorario") Then
        frmHorario.FillListView DiaSemana, Hora, IDRuta
    End If
    If CSM_Forms.IsLoaded("frmListaPasajero") Then
        frmListaPasajero.FillListView
    End If
End Sub

Public Sub RefreshList_RefreshPersona(ByVal IDPersona As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_PERSONA
    End If
    If CSM_Forms.IsLoaded("frmPersona") Then
        frmPersona.FillData IDPersona
    End If
    If CSM_Forms.IsLoaded("frmHorarioPropiedad") Then
        frmHorarioPropiedad.FillComboBoxPersona
    End If
    If CSM_Forms.IsLoaded("frmViajePropiedad") Then
        frmViajePropiedad.FillComboBoxConductor
    End If
    If CSM_Forms.IsLoaded("frmVehiculoMantenimientoAccionPropiedad") Then
        frmVehiculoMantenimientoAccionPropiedad.FillComboBoxConductor
    End If
    If CSM_Forms.IsLoaded("frmConductorRuta") Then
        frmConductorRuta.FillComboBoxConductor
    End If
    If CSM_Forms.IsLoaded("frmFranco") Then
        frmFranco.FillComboBoxConductor
    End If
    If CSM_Forms.IsLoaded("frmFrancoPropiedad") Then
        frmFrancoPropiedad.FillComboBoxConductor
    End If
End Sub

Public Sub RefreshList_RefreshPersonaHorario(ByVal IDPersona As Long, ByVal DiaSemana As Byte, ByVal Hora As Date, ByVal IDRuta As String, Optional ByVal Force As Boolean = False, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_PERSONA_HORARIO
    End If
    If CSM_Forms.IsLoaded("frmPersonaHorario") Then
        If Force Then
            frmPersonaHorario.ForceRefresh
        Else
            frmPersonaHorario.FillListView IDPersona, DiaSemana, Hora, IDRuta
        End If
    End If
End Sub

Public Sub RefreshList_RefreshPersonaRuta(ByVal IDPersona As Long, ByVal IDRuta As String, Optional ByVal Force As Boolean = False, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_PERSONA_RUTA
    End If
    If CSM_Forms.IsLoaded("frmPersonaRuta") Then
        If Force Then
            frmPersonaRuta.ForceRefresh
        Else
            frmPersonaRuta.FillListView IDPersona, IDRuta
        End If
    End If
End Sub

Public Sub RefreshList_RefreshPersonaPrepago(ByVal IDPersona As Long, ByVal IDRutaGrupo As Long, ByVal FechaInicio As Date, Optional ByVal Force As Boolean = False, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_PERSONA_PREPAGO
    End If
    If CSM_Forms.IsLoaded("frmPersonaPrepago") Then
        If Force Then
            frmPersonaPrepago.ForceRefresh
        Else
            frmPersonaPrepago.FillListView IDPersona, IDRutaGrupo, FechaInicio
        End If
    End If
End Sub

Public Sub RefreshList_RefreshPersonaRespuesta(ByVal IDPersona As Long, ByVal FechaHora As Date, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_PERSONA_RESPUESTA
    End If
    If CSM_Forms.IsLoaded("frmPersonaRespuesta") Then
        frmPersonaRespuesta.FillListView IDPersona, FechaHora
    End If
End Sub

Public Sub RefreshList_RefreshConductorRuta(ByVal IDPersona As Long, ByVal IDRuta As String, Optional ByVal Force As Boolean = False, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_CONDUCTOR_RUTA
    End If
    If CSM_Forms.IsLoaded("frmConductorRuta") Then
        If Force Then
            frmConductorRuta.ForceRefresh
        Else
            frmConductorRuta.FillListView IDPersona, IDRuta
        End If
    End If
End Sub

Public Sub RefreshList_RefreshViaje(ByVal FechaHora As Date, ByVal IDRuta As String, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_VIAJE
    End If
    If CSM_Forms.IsLoaded("frmViaje") Then
        frmViaje.FillListView FechaHora, IDRuta
    End If
    If CSM_Forms.IsLoaded("frmPersonaInfo") Then
        frmPersonaInfo.ForceRefresh
    End If
'    If CSM_Forms.IsLoaded("frmViajeConductor") Then
'        frmViajeConductor.FillListView
'    End If
    If CSM_Forms.IsLoaded("frmViajeAsistencia") Then
        frmViajeAsistencia.FillListViewViaje
    End If
End Sub

Public Sub RefreshList_RefreshViajeDetalle(ByVal FechaHora As Date, ByVal IDRuta As String, ByVal Indice As Long, Optional ByVal Force As Boolean = False, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_VIAJE_DETALLE
    End If
    If CSM_Forms.IsLoaded("frmViajeDetalle") Then
        If Force Then
            frmViajeDetalle.ForceRefresh
        Else
            frmViajeDetalle.FillListView FechaHora, IDRuta, Indice
        End If
    End If
    If CSM_Forms.IsLoaded("frmViajeDetalleTransferencia") Then
        frmViajeDetalleTransferencia.FillListViewLeft FechaHora, IDRuta, Indice
        frmViajeDetalleTransferencia.FillListViewRight FechaHora, IDRuta, Indice
    End If
    If CSM_Forms.IsLoaded("frmPersonaInfo") Then
        frmPersonaInfo.ForceRefresh
    End If
    If CSM_Forms.IsLoaded("frmComision") Then
        frmComision.FillListView FechaHora, IDRuta, Indice
    End If
    If CSM_Forms.IsLoaded("frmViajeAsistencia") Then
        frmViajeAsistencia.FillListViewViajeDetalle
    End If
End Sub

Public Sub RefreshList_RefreshCuentaCorriente(ByVal IDMovimiento As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_CUENTA_CORRIENTE
    End If
    If CSM_Forms.IsLoaded("frmCuentaCorriente") Then
        frmCuentaCorriente.FillListView IDMovimiento
    End If
End Sub

Public Sub RefreshList_RefreshCuentaCorrienteCaja(ByVal IDCuentaCorrienteCaja As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_CUENTA_CORRIENTE_CAJA
    End If
    If CSM_Forms.IsLoaded("frmCuentaCorrienteCaja") Then
        frmCuentaCorrienteCaja.FillListView IDCuentaCorrienteCaja
    End If
    If CSM_Forms.IsLoaded("frmViajeDetallePropiedad") Then
        frmViajeDetallePropiedad.FillComboBoxCuentaCorrienteCaja
    End If
    If CSM_Forms.IsLoaded("frmViajeDetalleAsistencia") Then
        frmViajeDetalleAsistencia.FillComboBoxCuentaCorrienteCaja
    End If
    If CSM_Forms.IsLoaded("frmViajeDetalleAsistenciaMultiple") Then
        frmViajeDetalleAsistenciaMultiple.FillComboBoxCuentaCorrienteCaja
    End If
    If CSM_Forms.IsLoaded("frmComisionAsistenciaMultiple") Then
        frmComisionAsistenciaMultiple.FillComboBoxCuentaCorrienteCaja
    End If
    If CSM_Forms.IsLoaded("frmCuentaCorriente") Then
        frmCuentaCorriente.FillComboBoxCuentaCorrienteCaja
    End If
    If CSM_Forms.IsLoaded("frmCuentaCorrientePropiedad") Then
        frmCuentaCorrientePropiedad.FillComboBoxCuentaCorrienteCaja
    End If
    If CSM_Forms.IsLoaded("frmCuentaCorrienteTransferencia") Then
        frmCuentaCorrienteTransferencia.FillComboBoxCuentaCorrienteCaja
    End If
End Sub

Public Sub RefreshList_RefreshCuentaCorrienteGrupo(ByVal IDCuentaCorrienteGrupo As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_CUENTA_CORRIENTE_GRUPO
    End If
    If CSM_Forms.IsLoaded("frmCuentaCorrienteGrupo") Then
        frmCuentaCorrienteGrupo.FillListView IDCuentaCorrienteGrupo
    End If
    If CSM_Forms.IsLoaded("frmCuentaCorriente") Then
        frmCuentaCorriente.FillComboBoxCuentaCorrienteGrupo
    End If
    If CSM_Forms.IsLoaded("frmCuentaCorrientePropiedad") Then
        frmCuentaCorrientePropiedad.FillComboBoxCuentaCorrienteGrupo
    End If
    If CSM_Forms.IsLoaded("frmCuentaCorrienteTransferencia") Then
        frmCuentaCorrienteTransferencia.FillComboBoxCuentaCorrienteGrupo
    End If
    If CSM_Forms.IsLoaded("frmListaPrecioPropiedad") Then
        frmListaPrecioPropiedad.FillComboBoxCuentaCorrienteGrupo
    End If
End Sub

Public Sub RefreshList_RefreshMedioPago(ByVal IDMedioPago As Byte, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_MEDIOPAGO
    End If
    If CSM_Forms.IsLoaded("frmMedioPago") Then
        frmMedioPago.FillListView IDMedioPago
    End If
    If CSM_Forms.IsLoaded("frmCuentaCorrientePropiedad") Then
        frmCuentaCorrientePropiedad.FillComboBoxMedioPago
    End If
    If CSM_Forms.IsLoaded("frmViajeDetallePropiedad") Then
        frmViajeDetallePropiedad.FillComboBoxMedioPago
    End If
    If CSM_Forms.IsLoaded("frmViajeDetalleAsistencia") Then
        frmViajeDetalleAsistencia.FillComboBoxMedioPago
    End If
End Sub

Public Sub RefreshList_RefreshFeriado(ByVal Fecha As Date, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_FERIADO
    End If
    If CSM_Forms.IsLoaded("frmFeriado") Then
        frmFeriado.FillListView Fecha
    End If
End Sub

Public Sub RefreshList_RefreshVehiculoMantenimiento(ByVal IDVehiculo As Long, ByVal IDVehiculoMantenimientoGrupo As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_VEHICULO_MANTENIMIENTO
    End If
    If CSM_Forms.IsLoaded("frmVehiculoMantenimiento") Then
        frmVehiculoMantenimiento.FillListView IDVehiculo, IDVehiculoMantenimientoGrupo
    End If
End Sub

Public Sub RefreshList_RefreshVehiculoMantenimientoGrupo(ByVal IDVehiculoMantenimientoGrupo As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_VEHICULO_MANTENIMIENTO_GRUPO
    End If
    If CSM_Forms.IsLoaded("frmVehiculoMantenimientoGrupo") Then
        frmVehiculoMantenimientoGrupo.FillListView IDVehiculoMantenimientoGrupo
    End If
    If CSM_Forms.IsLoaded("frmVehiculoMantenimientoPropiedad") Then
        frmVehiculoMantenimientoPropiedad.FillComboBoxVehiculoMantenimientoGrupo
    End If
    If CSM_Forms.IsLoaded("frmVehiculoMantenimientoAccion") Then
        frmVehiculoMantenimientoAccion.FillComboBoxVehiculoMantenimientoGrupo
    End If
    If CSM_Forms.IsLoaded("frmVehiculoMantenimientoAccionPropiedad") Then
        frmVehiculoMantenimientoAccionPropiedad.FillComboBoxVehiculoMantenimientoGrupo
    End If
End Sub

Public Sub RefreshList_RefreshVehiculoMantenimientoAccion(ByVal IDVehiculoMantenimientoAccion As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_VEHICULO_MANTENIMIENTO_ACCION
    End If
    If CSM_Forms.IsLoaded("frmVehiculoMantenimientoAccion") Then
        frmVehiculoMantenimientoAccion.FillListView IDVehiculoMantenimientoAccion
    End If
End Sub

Public Sub RefreshList_RefreshDocumentoTipo(ByVal IDDocumentoTipo As String, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_DOCUMENTO_TIPO
    End If
    If CSM_Forms.IsLoaded("frmDocumentoTipo") Then
        frmDocumentoTipo.FillListView IDDocumentoTipo
    End If
    If CSM_Forms.IsLoaded("frmPersonaPropiedad") Then
        frmPersonaPropiedad.FillComboBoxDocumentoTipo
    End If
End Sub

Public Sub RefreshList_RefreshTelefonoTipo(ByVal IDTelefonoTipo As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_TELEFONO_TIPO
    End If
    If CSM_Forms.IsLoaded("frmTelefonoTipo") Then
        frmTelefonoTipo.FillListView IDTelefonoTipo
    End If
    If CSM_Forms.IsLoaded("frmPersonaPropiedad") Then
        frmPersonaPropiedad.FillComboBoxTelefonoTipo
    End If
    If CSM_Forms.IsLoaded("frmContactoPropiedad") Then
        frmContactoPropiedad.FillComboBoxTelefonoTipo
    End If
End Sub

Public Sub RefreshList_RefreshPersonaAlarma(ByVal IDPersona As Long, ByVal IDPersonaAlarmaGrupo As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_PERSONA_ALARMA
    End If
    If CSM_Forms.IsLoaded("frmPersonaAlarma") Then
        frmPersonaAlarma.FillListView IDPersona, IDPersonaAlarmaGrupo
    End If
End Sub

Public Sub RefreshList_RefreshPersonaAlarmaGrupo(ByVal IDPersonaAlarmaGrupo As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_PERSONA_ALARMA_GRUPO
    End If
    If CSM_Forms.IsLoaded("frmPersonaAlarmaGrupo") Then
        frmPersonaAlarmaGrupo.FillListView IDPersonaAlarmaGrupo
    End If
    If CSM_Forms.IsLoaded("frmPersonaAlarmaPropiedad") Then
        frmPersonaAlarmaPropiedad.FillComboBoxPersonaAlarmaGrupo
    End If
End Sub

Public Sub RefreshList_RefreshAlarma(ByVal IDAlarma As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_ALARMA
    End If
    If CSM_Forms.IsLoaded("frmAlarma") Then
        frmAlarma.FillListView IDAlarma
    End If
End Sub

Public Sub RefreshList_RefreshContacto(ByVal IDContacto As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_CONTACTO
    End If
    If CSM_Forms.IsLoaded("frmContacto") Then
        frmContacto.FillListView IDContacto
    End If
End Sub

Public Sub RefreshList_RefreshContactoGrupo(ByVal IDContactoGrupo As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_CONTACTO_GRUPO
    End If
    If CSM_Forms.IsLoaded("frmContactoGrupo") Then
        frmContactoGrupo.FillListView IDContactoGrupo
    End If
    If CSM_Forms.IsLoaded("frmContacto") Then
        frmContacto.FillComboBoxContactoGrupo
    End If
    If CSM_Forms.IsLoaded("frmContactoPropiedad") Then
        frmContactoPropiedad.FillComboBoxContactoGrupo
    End If
End Sub

Public Sub RefreshList_RefreshRegistroLlamada(ByVal IDRegistroLlamada As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_REGISTROLLAMADA
    End If
    If CSM_Forms.IsLoaded("frmRegistroLlamada") Then
        frmRegistroLlamada.FillListView
    End If
End Sub

Public Sub Franco(ByVal Fecha As Date, ByVal IDPersona As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_FRANCO
    End If
    If CSM_Forms.IsLoaded("frmFranco") Then
        frmFranco.FillListView Fecha, IDPersona
    End If
End Sub

Public Sub Mensaje(ByVal IDMensaje As Long, Optional ByVal UpdateRefreshValue As Boolean = True)
    If UpdateRefreshValue Then
        RefreshList_UpdateValue MODULE_MENSAJE
    End If
    If CSM_Forms.IsLoaded("frmMensajeLista") Then
        frmMensajeLista.FillListView IDMensaje
    End If
End Sub

Public Sub Messenger()
    If pUsuario.CambioSemaforo() Then
        pMessengerBlinking = True
    End If
End Sub
