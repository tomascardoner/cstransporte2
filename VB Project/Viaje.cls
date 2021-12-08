VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Viaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDViaje As Long
Private mFechaHora As Date
Private mIDRuta As String
Private mRutaOtra As String
Private mIDPersona As Long
Private mKilometro As Integer
Private mDuracion As Integer
Private mImporte As Currency
Private mIDMedioPago As Byte
Private mCuotas As Byte
Private mOperacion As String
Private mImporteContado As Currency
Private mIDCuentaCorrienteCaja As Long
Private mCharter As Boolean
Private mIDVehiculo As Long
Private mIDConductor As Long
Private mAcreditaSueldo As Boolean
Private mIDConductor2 As Long
Private mAcreditaSueldo2 As Boolean
Private mDiaSemanaBase As Integer
Private mEstado As String
Private mAsientoOcupado As Integer
Private mNotas As String
Private mPersonal As Boolean
Private mFechaHoraCreacion As Date
Private mIDUsuarioCreacion As Integer
Private mFechaHoraModificacion As Date
Private mIDUsuarioModificacion As Integer
Private mFechaHoraEnProgreso As Date
Private mIDUsuarioEnProgreso As String
Private mFechaHoraFinalizado As Date
Private mIDUsuarioFinalizado As String
Private mFechaHoraCancelado As Date
Private mIDUsuarioCancelado As String

'STATE VARIABLES
Private mIsNew As Boolean
Private mIsCopy As Boolean
Private mIsDirty As Boolean
Private mNoMatch As Boolean

'BEHAVIOR VARIABLES
Private mNoMatchRaiseError As Boolean
Private mRefreshListSkip As Boolean

'MISC VARIABLES
Private mIDOrigen As Long
Private mIDDestino As Long
Private mCuentaCorrienteAsignar As Boolean
Private mAsientoExcedido As Boolean
Private mFechaHoraOriginal As Date
Private mIDRutaOriginal As String
Private mKilometroOriginal As Integer
Private mIDVehiculoOriginal As Long
Private mIDConductorOriginal As Long
Private mIDConductor2Original As Long
Private mEstadoOriginal As String

'INTERNAL VARIABLES
Private mrecData As ADODB.Recordset

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDViaje() As Long
    IDViaje = mIDViaje
End Property

Public Property Get IDViaje_Formatted() As String
    IDViaje_Formatted = Format(mIDViaje, "#,###")
End Property

Public Property Let IDViaje(ByVal value As Long)
    If value <> mIDViaje Then
        mIsDirty = True
    End If
    mIDViaje = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHora() As Date
    FechaHora = mFechaHora
End Property

Public Property Get FechaHora_Weekday() As Integer
    FechaHora_Weekday = Weekday(mFechaHora)
End Property

Public Property Get FechaHora_WeekdayName() As String
    FechaHora_WeekdayName = WeekdayName(Weekday(mFechaHora))
End Property

Public Property Get FechaHora_Formatted() As String
    FechaHora_Formatted = Format(mFechaHora, "Short Date") & " " & Format(mFechaHora, "Short Time")
End Property

Public Property Get FechaHora_FormattedAsDate() As String
    FechaHora_FormattedAsDate = Format(mFechaHora, "Short Date")
End Property

Public Property Get FechaHora_FormattedAsTime() As String
    FechaHora_FormattedAsTime = Format(mFechaHora, "Short Time")
End Property

Public Property Let FechaHora(ByVal value As Date)
    If value <> mFechaHora Then
        mIsDirty = True
    End If
    mFechaHora = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDRuta() As String
    IDRuta = mIDRuta
End Property

Public Property Let IDRuta(ByVal value As String)
    If value <> mIDRuta Then
        mIsDirty = True
    End If
    mIDRuta = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RutaOtra() As String
    RutaOtra = mRutaOtra
End Property

Public Property Let RutaOtra(ByVal value As String)
    If value <> mRutaOtra Then
        mIsDirty = True
    End If
    mRutaOtra = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Ruta_DisplayName() As String
    Ruta_DisplayName = mIDRuta & IIf(mIDRuta = pParametro.Ruta_ID_Otra Or mIDRuta = pParametro.Ruta_Paquete_ID, ": " & mRutaOtra, "")
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDPersona() As Long
    IDPersona = mIDPersona
End Property

Public Property Let IDPersona(ByVal value As Long)
    If value <> mIDPersona Then
        mIsDirty = True
    End If
    mIDPersona = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Kilometro() As Integer
    Kilometro = mKilometro
End Property

Public Property Let Kilometro(ByVal value As Integer)
    If value <> mKilometro Then
        mIsDirty = True
    End If
    mKilometro = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Duracion() As Integer
    Duracion = mDuracion
End Property

Public Property Let Duracion(ByVal value As Integer)
    If value <> mDuracion Then
        mIsDirty = True
    End If
    mDuracion = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Importe() As Currency
    Importe = mImporte
End Property

Public Property Get Importe_Formatted() As String
    Importe_Formatted = Format(mImporte, "Currency")
End Property

Public Property Let Importe(ByVal value As Currency)
    If value <> mImporte Then
        mIsDirty = True
    End If
    mImporte = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ImporteContado() As Currency
    ImporteContado = mImporteContado
End Property

Public Property Get ImporteContado_Formatted() As String
    ImporteContado_Formatted = Format(mImporteContado, "Currency")
End Property

Public Property Let ImporteContado(ByVal value As Currency)
    If value <> mImporteContado Then
        mIsDirty = True
    End If
    mImporteContado = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDMedioPago() As Byte
    IDMedioPago = mIDMedioPago
End Property

Public Property Let IDMedioPago(ByVal value As Byte)
    If value <> mIDMedioPago Then
        mIsDirty = True
    End If
    mIDMedioPago = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Cuotas() As Byte
    Cuotas = mCuotas
End Property

Public Property Get Cuotas_Formatted() As String
    Cuotas_Formatted = IIf(mCuotas = 0, "", mCuotas)
End Property

Public Property Let Cuotas(ByVal value As Byte)
    If value <> mCuotas Then
        mIsDirty = True
    End If
    mCuotas = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Operacion() As String
    Operacion = mOperacion
End Property

Public Property Let Operacion(ByVal value As String)
    If value <> mOperacion Then
        mIsDirty = True
    End If
    mOperacion = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDCuentaCorrienteCaja() As Long
    IDCuentaCorrienteCaja = mIDCuentaCorrienteCaja
End Property

Public Property Let IDCuentaCorrienteCaja(ByVal value As Long)
    If value <> mIDCuentaCorrienteCaja Then
        mIsDirty = True
    End If
    mIDCuentaCorrienteCaja = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Charter() As Boolean
    Charter = mCharter
End Property

Public Property Let Charter(ByVal value As Boolean)
    If value <> mCharter Then
        mIsDirty = True
    End If
    mCharter = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDVehiculo() As Long
    IDVehiculo = mIDVehiculo
End Property

Public Property Let IDVehiculo(ByVal value As Long)
    If value <> mIDVehiculo Then
        mIsDirty = True
    End If
    mIDVehiculo = value
End Property

Public Property Get Vehiculo() As Vehiculo
    Static oVehiculo As Vehiculo
    
    If mIDVehiculo = 0 Then
        Set Vehiculo = Nothing
    Else
        If oVehiculo Is Nothing Then
            Set oVehiculo = New Vehiculo
        End If
        If oVehiculo.IDVehiculo <> mIDVehiculo Then
            oVehiculo.IDVehiculo = mIDVehiculo
            Call oVehiculo.Load
        End If
        Set Vehiculo = oVehiculo
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDConductor() As Long
    IDConductor = mIDConductor
End Property

Public Property Let IDConductor(ByVal value As Long)
    If value <> mIDConductor Then
        mIsDirty = True
    End If
    mIDConductor = value
End Property

Public Property Get Conductor() As Persona
    Static oConductor As Persona
    
    If mIDConductor = 0 Then
        Set Conductor = Nothing
    Else
        If oConductor Is Nothing Then
            Set oConductor = New Persona
        End If
        If oConductor.IDPersona <> mIDConductor Then
            oConductor.IDPersona = mIDConductor
            Call oConductor.Load
        End If
        Set Conductor = oConductor
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get AcreditaSueldo() As Boolean
    AcreditaSueldo = mAcreditaSueldo
End Property

Public Property Let AcreditaSueldo(ByVal value As Boolean)
    If value <> mAcreditaSueldo Then
        mIsDirty = True
    End If
    mAcreditaSueldo = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDConductor2() As Long
    IDConductor2 = mIDConductor2
End Property

Public Property Let IDConductor2(ByVal value As Long)
    If value <> mIDConductor2 Then
        mIsDirty = True
    End If
    mIDConductor2 = value
End Property

Public Property Get Conductor2() As Persona
    Static oConductor As Persona
    
    If mIDConductor2 = 0 Then
        Set Conductor2 = Nothing
    Else
        If oConductor Is Nothing Then
            Set oConductor = New Persona
        End If
        If oConductor.IDPersona <> mIDConductor2 Then
            oConductor.IDPersona = mIDConductor2
            Call oConductor.Load
        End If
        Set Conductor2 = oConductor
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get AcreditaSueldo2() As Boolean
    AcreditaSueldo2 = mAcreditaSueldo2
End Property

Public Property Let AcreditaSueldo2(ByVal value As Boolean)
    If value <> mAcreditaSueldo2 Then
        mIsDirty = True
    End If
    mAcreditaSueldo2 = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DiaSemanaBase() As Integer
    DiaSemanaBase = mDiaSemanaBase
End Property

Public Property Let DiaSemanaBase(ByVal value As Integer)
    If value <> mDiaSemanaBase Then
        mIsDirty = True
    End If
    mDiaSemanaBase = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Estado() As String
    Estado = mEstado
End Property

Public Property Get Estado_ToString() As String
    Select Case mEstado
        Case VIAJE_ESTADO_ACTIVO
            Estado_ToString = VIAJE_ESTADO_ACTIVO_NOMBRE
        Case VIAJE_ESTADO_EN_PROGRESO
            Estado_ToString = VIAJE_ESTADO_EN_PROGRESO_NOMBRE
        Case VIAJE_ESTADO_FINALIZADO
            Estado_ToString = VIAJE_ESTADO_FINALIZADO_NOMBRE
        Case VIAJE_ESTADO_CANCELADO
            Estado_ToString = VIAJE_ESTADO_CANCELADO_NOMBRE
        Case Else
            Estado_ToString = ""
    End Select
End Property

Public Property Get Estado_ToColor() As Long
    Select Case mEstado
        Case VIAJE_ESTADO_ACTIVO
            Estado_ToColor = pParametro.Viaje_Estado_Activo_Color
        Case VIAJE_ESTADO_EN_PROGRESO
            Estado_ToColor = pParametro.Viaje_Estado_EnProgreso_Color
        Case VIAJE_ESTADO_FINALIZADO
            Estado_ToColor = pParametro.Viaje_Estado_Finalizado_Color
        Case VIAJE_ESTADO_CANCELADO
            Estado_ToColor = pParametro.Viaje_Estado_Cancelado_Color
        Case Else
            Estado_ToColor = vbWindowText
    End Select
End Property

Public Property Get Estado_ToBold() As Boolean
    Select Case mEstado
        Case VIAJE_ESTADO_ACTIVO
            Estado_ToBold = pParametro.Viaje_Estado_Activo_Bold
        Case VIAJE_ESTADO_EN_PROGRESO
            Estado_ToBold = pParametro.Viaje_Estado_EnProgreso_Bold
        Case VIAJE_ESTADO_FINALIZADO
            Estado_ToBold = pParametro.Viaje_Estado_Finalizado_Bold
        Case VIAJE_ESTADO_CANCELADO
            Estado_ToBold = pParametro.Viaje_Estado_Cancelado_Bold
        Case Else
            Estado_ToBold = False
    End Select
End Property

Public Property Let Estado(ByVal value As String)
    If value <> mEstado Then
        mIsDirty = True
    End If
    mEstado = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get AsientoOcupado() As Integer
    AsientoOcupado = mAsientoOcupado
End Property

Public Property Let AsientoOcupado(ByVal value As Integer)
    If value <> mAsientoOcupado Then
        mIsDirty = True
    End If
    mAsientoOcupado = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get AsientoLibre() As Integer
    If IDVehiculo = 0 Then
        AsientoLibre = 0
    Else
        AsientoLibre = Me.Vehiculo.Asiento - mAsientoOcupado
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Notas() As String
    Notas = mNotas
End Property

Public Property Let Notas(ByVal value As String)
    If value <> mNotas Then
        mIsDirty = True
    End If
    mNotas = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Personal() As Boolean
    Personal = mPersonal
End Property

Public Property Let Personal(ByVal value As Boolean)
    If value <> mPersonal Then
        mIsDirty = True
    End If
    mPersonal = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHoraCreacion() As Date
    FechaHoraCreacion = mFechaHoraCreacion
End Property

Public Property Let FechaHoraCreacion(ByVal value As Date)
    If value <> mFechaHoraCreacion Then
        mIsDirty = True
    End If
    mFechaHoraCreacion = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDUsuarioCreacion() As String
    IDUsuarioCreacion = mIDUsuarioCreacion
End Property

Public Property Let IDUsuarioCreacion(ByVal value As String)
    If value <> mIDUsuarioCreacion Then
        mIsDirty = True
    End If
    mIDUsuarioCreacion = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHoraModificacion() As Date
    FechaHoraModificacion = mFechaHoraModificacion
End Property

Public Property Let FechaHoraModificacion(ByVal value As Date)
    If value <> mFechaHoraModificacion Then
        mIsDirty = True
    End If
    mFechaHoraModificacion = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDUsuarioModificacion() As String
    IDUsuarioModificacion = mIDUsuarioModificacion
End Property

Public Property Let IDUsuarioModificacion(ByVal value As String)
    If value <> mIDUsuarioModificacion Then
        mIsDirty = True
    End If
    mIDUsuarioModificacion = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHoraEnProgreso() As Date
    FechaHoraEnProgreso = mFechaHoraEnProgreso
End Property

Public Property Let FechaHoraEnProgreso(ByVal value As Date)
    If value <> mFechaHoraEnProgreso Then
        mIsDirty = True
    End If
    mFechaHoraEnProgreso = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDUsuarioEnProgreso() As String
    IDUsuarioEnProgreso = mIDUsuarioEnProgreso
End Property

Public Property Let IDUsuarioEnProgreso(ByVal value As String)
    If value <> mIDUsuarioEnProgreso Then
        mIsDirty = True
    End If
    mIDUsuarioEnProgreso = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHoraFinalizado() As Date
    FechaHoraFinalizado = mFechaHoraFinalizado
End Property

Public Property Let FechaHoraFinalizado(ByVal value As Date)
    If value <> mFechaHoraFinalizado Then
        mIsDirty = True
    End If
    mFechaHoraFinalizado = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDUsuarioFinalizado() As String
    IDUsuarioFinalizado = mIDUsuarioFinalizado
End Property

Public Property Let IDUsuarioFinalizado(ByVal value As String)
    If value <> mIDUsuarioFinalizado Then
        mIsDirty = True
    End If
    mIDUsuarioFinalizado = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHoraCancelado() As Date
    FechaHoraCancelado = mFechaHoraCancelado
End Property

Public Property Let FechaHoraCancelado(ByVal value As Date)
    If value <> mFechaHoraCancelado Then
        mIsDirty = True
    End If
    mFechaHoraCancelado = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDUsuarioCancelado() As String
    IDUsuarioCancelado = mIDUsuarioCancelado
End Property

Public Property Let IDUsuarioCancelado(ByVal value As String)
    If value <> mIDUsuarioCancelado Then
        mIsDirty = True
    End If
    mIDUsuarioCancelado = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IsNew() As Boolean
    IsNew = mIsNew
End Property

Public Property Get IsCopy() As Boolean
    IsCopy = mIsCopy
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = mIsDirty
End Property

Public Property Get NoMatch() As Boolean
    NoMatch = mNoMatch
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get NoMatchRaiseError() As Boolean
    NoMatchRaiseError = mNoMatchRaiseError
End Property

Public Property Let NoMatchRaiseError(ByVal value As Boolean)
    mNoMatchRaiseError = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RefreshListSkip() As Boolean
    RefreshListSkip = mRefreshListSkip
End Property

Public Property Let RefreshListSkip(ByVal value As Boolean)
    mRefreshListSkip = value
End Property

Public Sub RefreshList()
    RefreshList_Module.RefreshList_RefreshViaje FechaHora, IDRuta
    RefreshList_Module.RefreshList_RefreshVehiculoUtilizacion
End Sub

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDOrigen() As Long
    IDOrigen = mIDOrigen
End Property

Public Property Let IDOrigen(ByVal value As Long)
    mIDOrigen = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDDestino() As Long
    IDDestino = mIDDestino
End Property

Public Property Let IDDestino(ByVal value As Long)
    mIDDestino = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get CuentaCorrienteAsignar() As Boolean
    CuentaCorrienteAsignar = mCuentaCorrienteAsignar
End Property

Public Property Let CuentaCorrienteAsignar(ByVal value As Boolean)
    mCuentaCorrienteAsignar = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function OpenRecordset() As Boolean
    Dim cmdData As ADODB.command
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    Screen.MousePointer = vbHourglass
      
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Viaje_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, mIDRuta)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , pParametro.Database_CursorType, adLockOptimistic
    If pParametro.Database_CursorLocation = adUseClient Then
        mrecData.Properties("Update Criteria").value = adCriteriaKey
    End If
    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Viaje.OpenRecordset", "Error al abrir la tabla de Viajes."
    If Not mrecData Is Nothing Then
        Set mrecData = Nothing
    End If
End Function

Public Function Load() As Boolean
    If Not OpenRecordset() Then
        Exit Function
    End If

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Call InitializeValues

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDViaje = mrecData("IDViaje").value
        mRutaOtra = mrecData("RutaOtra").value & ""
        mIDPersona = Val(mrecData("IDPersona").value & "")
        mKilometro = Val(mrecData("Kilometro").value & "")
        mDuracion = Val(mrecData("Duracion").value & "")
        mImporte = IIf(IsNull(mrecData("Importe").value), 0, mrecData("Importe").value)
        mImporteContado = IIf(IsNull(mrecData("ImporteContado").value), 0, mrecData("ImporteContado").value)
        mIDMedioPago = Val(mrecData("IDMedioPago").value & "")
        mCuotas = Val(mrecData("Cuotas").value & "")
        mOperacion = mrecData("Operacion").value & ""
        mIDCuentaCorrienteCaja = Val(mrecData("IDCuentaCorrienteCaja").value & "")
        mCharter = mrecData("Charter").value
        mIDVehiculo = Val(mrecData("IDVehiculo").value & "")
        mIDConductor = Val(mrecData("IDConductor").value & "")
        mAcreditaSueldo = mrecData("AcreditaSueldo").value
        mIDConductor2 = Val(mrecData("IDConductor2").value & "")
        mAcreditaSueldo2 = mrecData("AcreditaSueldo2").value
        mDiaSemanaBase = mrecData("DiaSemanaBase").value
        mEstado = mrecData("Estado").value
        mAsientoOcupado = IIf(IsNull(mrecData("AsientoOcupado").value), -1, mrecData("AsientoOcupado").value)
        mNotas = mrecData("Notas").value & ""
        mPersonal = mrecData("Personal").value
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").value
        mFechaHoraEnProgreso = IIf(IsNull(mrecData("FechaHoraEnProgreso").value), DATE_TIME_FIELD_NULL_VALUE, mrecData("FechaHoraEnProgreso").value)
        mIDUsuarioEnProgreso = mrecData("IDUsuarioEnProgreso").value & ""
        mFechaHoraFinalizado = IIf(IsNull(mrecData("FechaHoraFinalizado").value), DATE_TIME_FIELD_NULL_VALUE, mrecData("FechaHoraFinalizado").value)
        mIDUsuarioFinalizado = mrecData("IDUsuarioFinalizado").value & ""
        mFechaHoraCancelado = IIf(IsNull(mrecData("FechaHoraCancelado").value), DATE_TIME_FIELD_NULL_VALUE, mrecData("FechaHoraCancelado").value)
        mIDUsuarioCancelado = mrecData("IDUsuarioCancelado").value & ""
    End If
    mFechaHoraOriginal = mFechaHora
    mIDRutaOriginal = mIDRuta
    mKilometroOriginal = mKilometro
    mIDVehiculoOriginal = mIDVehiculo
    mIDConductorOriginal = mIDConductor
    mIDConductor2Original = mIDConductor2
    mEstadoOriginal = mEstado
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Viaje.Load", "Error al obtener los datos del Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        'mFechaHora = DATE_TIME_FIELD_NULL_VALUE
        'mIDRuta = ""

        mIsNew = True
        mIsCopy = True
        mIsDirty = True
    End If
End Function

Public Function Update() As Boolean
    Dim Ruta As Ruta
    Dim Condicionales As Long
    Dim ConductorOriginal As Persona
    
    Dim VerificarOtrosPrepagos As Boolean
    Dim VerificarEstePrepago As Boolean
        
    Dim cmdData As ADODB.command
    
    If Not mIsDirty Then
        Update = True
        Exit Function
    End If

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Viaje_" & IIf(mIsNew, "Insert", "Update")
    cmdData.CommandType = adCmdStoredProc
    cmdData.NamedParameters = True
    If mIsNew Then
        cmdData.Parameters.Append cmdData.CreateParameter("@IDViaje", adInteger, adParamOutput)
    Else
        cmdData.Parameters.Append cmdData.CreateParameter("@FechaHora_Original", adDate, adParamInput, , mFechaHoraOriginal)
        cmdData.Parameters.Append cmdData.CreateParameter("@IDRuta_Original", adChar, adParamInput, 20, mIDRutaOriginal)
    End If
    cmdData.Parameters.Append cmdData.CreateParameter("@FechaHora", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDRuta", adChar, adParamInput, 20, mIDRuta)
    Select Case mIDRuta
        Case pParametro.Ruta_ID_Otra
            cmdData.Parameters.Append cmdData.CreateParameter("@RutaOtra", adVarChar, adParamInput, 50, IIf(mRutaOtra = "", Null, mRutaOtra))
            cmdData.Parameters.Append cmdData.CreateParameter("@IDPersona", adInteger, adParamInput, , mIDPersona)
            cmdData.Parameters.Append cmdData.CreateParameter("@Kilometro", adSmallInt, adParamInput, , IIf(mKilometro = 0, Null, mKilometro))
            cmdData.Parameters.Append cmdData.CreateParameter("@Duracion", adSmallInt, adParamInput, , IIf(mDuracion = 0, Null, mDuracion))
            cmdData.Parameters.Append cmdData.CreateParameter("@Importe", adCurrency, adParamInput, , IIf(mImporte = 0, Null, mImporte))
            cmdData.Parameters.Append cmdData.CreateParameter("@ImporteContado", adCurrency, adParamInput, , IIf(mImporteContado = 0, Null, mImporteContado))
            cmdData.Parameters.Append cmdData.CreateParameter("@IDMedioPago", adTinyInt, adParamInput, , IIf(mIDMedioPago = 0, Null, mIDMedioPago))
            cmdData.Parameters.Append cmdData.CreateParameter("@Cuotas", adTinyInt, adParamInput, , IIf(mCuotas = 0, Null, mCuotas))
            cmdData.Parameters.Append cmdData.CreateParameter("@Operacion", adVarChar, adParamInput, 20, IIf(Trim(mOperacion) = "", Null, mOperacion))
            cmdData.Parameters.Append cmdData.CreateParameter("@IDCuentaCorrienteCaja", adInteger, adParamInput, , IIf(mImporteContado = 0, Null, mIDCuentaCorrienteCaja))
        Case pParametro.Ruta_Paquete_ID
            cmdData.Parameters.Append cmdData.CreateParameter("@RutaOtra", adVarChar, adParamInput, 50, IIf(mRutaOtra = "", Null, mRutaOtra))
            cmdData.Parameters.Append cmdData.CreateParameter("@IDPersona", adInteger, adParamInput, , Null)
            cmdData.Parameters.Append cmdData.CreateParameter("@Kilometro", adSmallInt, adParamInput, , IIf(mKilometro = 0, Null, mKilometro))
            cmdData.Parameters.Append cmdData.CreateParameter("@Duracion", adSmallInt, adParamInput, , IIf(mDuracion = 0, Null, mDuracion))
            cmdData.Parameters.Append cmdData.CreateParameter("@Importe", adCurrency, adParamInput, , IIf(mImporte = 0, Null, mImporte))
            cmdData.Parameters.Append cmdData.CreateParameter("@ImporteContado", adCurrency, adParamInput, , Null)
            cmdData.Parameters.Append cmdData.CreateParameter("@IDMedioPago", adTinyInt, adParamInput, , Null)
            cmdData.Parameters.Append cmdData.CreateParameter("@Cuotas", adTinyInt, adParamInput, , Null)
            cmdData.Parameters.Append cmdData.CreateParameter("@Operacion", adVarChar, adParamInput, 20, Null)
            cmdData.Parameters.Append cmdData.CreateParameter("@IDCuentaCorrienteCaja", adInteger, adParamInput, , Null)
        Case Else
            cmdData.Parameters.Append cmdData.CreateParameter("@RutaOtra", adVarChar, adParamInput, 50, Null)
            cmdData.Parameters.Append cmdData.CreateParameter("@IDPersona", adInteger, adParamInput, , Null)
            cmdData.Parameters.Append cmdData.CreateParameter("@Kilometro", adSmallInt, adParamInput, , IIf(mKilometro = 0, Null, mKilometro))
            cmdData.Parameters.Append cmdData.CreateParameter("@Duracion", adSmallInt, adParamInput, , IIf(mDuracion = 0, Null, mDuracion))
            cmdData.Parameters.Append cmdData.CreateParameter("@Importe", adCurrency, adParamInput, , Null)
            cmdData.Parameters.Append cmdData.CreateParameter("@ImporteContado", adCurrency, adParamInput, , Null)
            cmdData.Parameters.Append cmdData.CreateParameter("@IDMedioPago", adTinyInt, adParamInput, , Null)
            cmdData.Parameters.Append cmdData.CreateParameter("@Cuotas", adTinyInt, adParamInput, , Null)
            cmdData.Parameters.Append cmdData.CreateParameter("@Operacion", adVarChar, adParamInput, 20, Null)
            cmdData.Parameters.Append cmdData.CreateParameter("@IDCuentaCorrienteCaja", adInteger, adParamInput, , Null)
    End Select
    cmdData.Parameters.Append cmdData.CreateParameter("@Charter", adBoolean, adParamInput, , mCharter)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDVehiculo", adInteger, adParamInput, , IIf(mIDVehiculo = 0, Null, mIDVehiculo))
    cmdData.Parameters.Append cmdData.CreateParameter("@IDConductor", adInteger, adParamInput, , IIf(mIDConductor = 0, Null, mIDConductor))
    cmdData.Parameters.Append cmdData.CreateParameter("@AcreditaSueldo", adBoolean, adParamInput, , mAcreditaSueldo)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDConductor2", adInteger, adParamInput, , IIf(mIDConductor2 = 0, Null, mIDConductor2))
    cmdData.Parameters.Append cmdData.CreateParameter("@AcreditaSueldo2", adBoolean, adParamInput, , mAcreditaSueldo2)
    cmdData.Parameters.Append cmdData.CreateParameter("@DiaSemanaBase", adTinyInt, adParamInput, , mDiaSemanaBase)
    cmdData.Parameters.Append cmdData.CreateParameter("@Estado", adChar, adParamInput, 2, IIf(mIsNew, VIAJE_ESTADO_ACTIVO, mEstado))
    cmdData.Parameters.Append cmdData.CreateParameter("@Notas", adVarChar, adParamInput, 8000, IIf(Trim(mNotas) = "", Null, mNotas))
    cmdData.Parameters.Append cmdData.CreateParameter("@Personal", adBoolean, adParamInput, , mPersonal)
    If mIsNew Then
        cmdData.Parameters.Append cmdData.CreateParameter("@IDUsuarioCreacion", adChar, adParamInput, 30, pUsuario.IDUsuario)
    Else
        cmdData.Parameters.Append cmdData.CreateParameter("@IDUsuarioModificacion", adChar, adParamInput, 30, pUsuario.IDUsuario)
        If mEstado <> mEstadoOriginal Then
            Select Case mEstado
                Case VIAJE_ESTADO_ACTIVO
                Case VIAJE_ESTADO_EN_PROGRESO
                    mFechaHoraEnProgreso = Now
                    mIDUsuarioEnProgreso = pUsuario.IDUsuario
                Case VIAJE_ESTADO_FINALIZADO
                    mFechaHoraFinalizado = Now
                    mIDUsuarioFinalizado = pUsuario.IDUsuario
                Case VIAJE_ESTADO_CANCELADO
                    VerificarOtrosPrepagos = True
                    mFechaHoraCancelado = Now
                    mIDUsuarioCancelado = pUsuario.IDUsuario
            End Select
            If mEstadoOriginal = VIAJE_ESTADO_CANCELADO Then
                VerificarEstePrepago = True
            End If
        End If
        cmdData.Parameters.Append cmdData.CreateParameter("@FechaHoraEnProgreso", adDate, adParamInput, , IIf(mFechaHoraEnProgreso = DATE_TIME_FIELD_NULL_VALUE, Null, mFechaHoraEnProgreso))
        cmdData.Parameters.Append cmdData.CreateParameter("@IDUsuarioEnProgreso", adChar, adParamInput, 30, IIf(mIDUsuarioEnProgreso = "", Null, mIDUsuarioEnProgreso))
        cmdData.Parameters.Append cmdData.CreateParameter("@FechaHoraFinalizado", adDate, adParamInput, , IIf(mFechaHoraFinalizado = DATE_TIME_FIELD_NULL_VALUE, Null, mFechaHoraFinalizado))
        cmdData.Parameters.Append cmdData.CreateParameter("@IDUsuarioFinalizado", adChar, adParamInput, 30, IIf(mIDUsuarioFinalizado = "", Null, mIDUsuarioFinalizado))
        cmdData.Parameters.Append cmdData.CreateParameter("@FechaHoraCancelado", adDate, adParamInput, , IIf(mFechaHoraCancelado = DATE_TIME_FIELD_NULL_VALUE, Null, mFechaHoraCancelado))
        cmdData.Parameters.Append cmdData.CreateParameter("@IDUsuarioCancelado", adChar, adParamInput, 30, IIf(mIDUsuarioCancelado = "", Null, mIDUsuarioCancelado))
    End If
    cmdData.Execute 0
    If mIsNew Then
        mIDViaje = cmdData.Parameters("@IDViaje").value
    End If
    Set cmdData = Nothing
    
    If Not mIsNew Then
        If mIDConductorOriginal <> 0 And mIDConductorOriginal <> mIDConductor Then
            Set ConductorOriginal = New Persona
            ConductorOriginal.IDPersona = mIDConductorOriginal
            Call ConductorOriginal.Load
            If Me.Conductor Is Nothing Then
                LogAccionAdd ENTIDAD_TIPO_VIAJE, "Cambio de Conductor: " & Me.FechaHora_Formatted & " - " & mIDRuta & " // De: " & ConductorOriginal.ApellidoNombre & " a: <<Ninguno>>"
            Else
                LogAccionAdd ENTIDAD_TIPO_VIAJE, "Cambio de Conductor: " & Me.FechaHora_Formatted & " - " & mIDRuta & " // De: " & ConductorOriginal.ApellidoNombre & " a: " & Me.Conductor.ApellidoNombre
            End If
            Set ConductorOriginal = Nothing
        End If
        If mIDConductor2Original <> 0 And mIDConductor2Original <> mIDConductor2 Then
            Set ConductorOriginal = New Persona
            ConductorOriginal.IDPersona = mIDConductor2Original
            Call ConductorOriginal.Load
            If Me.Conductor2 Is Nothing Then
                LogAccionAdd ENTIDAD_TIPO_VIAJE, "Cambio de Conductor 2: " & Me.FechaHora_Formatted & " - " & mIDRuta & " // De: " & ConductorOriginal.ApellidoNombre & " a: <<Ninguno>>"
            Else
                LogAccionAdd ENTIDAD_TIPO_VIAJE, "Cambio de Conductor 2: " & Me.FechaHora_Formatted & " - " & mIDRuta & " // De: " & ConductorOriginal.ApellidoNombre & " a: " & Me.Conductor2.ApellidoNombre
            End If
            Set ConductorOriginal = Nothing
        End If
    End If

    mIsCopy = False
    mIsDirty = False
    
    If mIsNew Then
        If DateDiff("n", Now, mFechaHora) >= 0 Then
            Call GenerarReservasFijas
        End If
        
        If mIDRuta = pParametro.Ruta_ID_Otra Or mIDRuta = pParametro.Ruta_Paquete_ID Then
            Call CuentaCorriente_Asignar
        End If
    Else
        If mEstadoOriginal = VIAJE_ESTADO_FINALIZADO And mIDVehiculoOriginal <> 0 Then
            If mIDVehiculoOriginal <> mIDVehiculo Then
                'CAMBIO EL VEHICULO
                'RESTO LOS DATOS AL VEHICULO ORIGINAL
                VehiculoCambiarDatos mIDVehiculoOriginal, False
    
                'ACTUALIZO LOS DATOS DEL VEHICULO NUEVO
                If Estado = VIAJE_ESTADO_FINALIZADO And IDVehiculo <> 0 Then
                    VehiculoCambiarDatos mIDVehiculo, True
                End If
            Else
                If mIDRutaOriginal <> mIDRuta Then
                    'CAMBIO LA RUTA
                    'RESTO LOS DATOS DE LA RUTA ORIGINAL
                    VehiculoCambiarDatos mIDVehiculoOriginal, False
                    
                    'ACTUALIZO LOS DATOS DE LA RUTA NUEVA
                    If mEstado = VIAJE_ESTADO_FINALIZADO And mIDVehiculo <> 0 Then
                        VehiculoCambiarDatos mIDVehiculo, True
                    End If
                Else
                    If mKilometro <> mKilometroOriginal Then
                        'CAMBIARON LOS KILOMETROS
                        'RESTO EL KILOMETRAJE ANTERIOR
                        VehiculoCambiarDatos mIDVehiculoOriginal, False
                        
                        'AGREGO EL KILOMETRAJE NUEVO
                        If mEstado = VIAJE_ESTADO_FINALIZADO And mIDVehiculo <> 0 Then
                            VehiculoCambiarDatos mIDVehiculo, True
                        End If
                    Else
                        If mEstado <> VIAJE_ESTADO_FINALIZADO Or mIDVehiculo = 0 Then
                            VehiculoCambiarDatos mIDVehiculoOriginal, False
                        End If
                    End If
                End If
            End If
        Else
            If mEstado = VIAJE_ESTADO_FINALIZADO And mIDVehiculo <> 0 Then
                VehiculoCambiarDatos mIDVehiculo, True
            End If
        End If
        
        'VERIFICO EL TEMA PREPAGOS
        If VerificarOtrosPrepagos Then
            Call VerificarPrepagosOtrasReservas
        End If
        If VerificarEstePrepago Then
            Call VerificarPrepagosEstasReservas
        End If
        
        If mIDRuta = pParametro.Ruta_ID_Otra Then
            Call CuentaCorriente_Asignar
        Else
            Call CuentaCorriente_Asignar_Sueldo
        End If
        
        mAsientoExcedido = False
        
        Call Asiento_Asignar
            
        'VERIFICO SI EL NUEVO VEHICULO TIENE CAPACIDAD SUFICIENTE
        If mIDVehiculoOriginal <> mIDVehiculo And mAsientoExcedido Then
            MsgBox "El Vehículo asignado al Viaje no tiene Asientos suficientes para todas las Reservas.", vbExclamation, App.Title
        End If
    End If
    
    If Not mRefreshListSkip Then
        Call RefreshList
    End If
    
    If mIsNew Then
        'Chequeo las Reservas Condicionales
        Condicionales = GetDetalleCountByTipoEstado(OCUPANTE_TIPO_PASAJERO, VIAJE_DETALLE_ESTADO_CONDICIONAL)
        If Condicionales > 0 Then
            MsgBox "Hay " & IIf(Condicionales = 1, "una Reserva", Condicionales & " Reservas") & " con Estado Condicional en este Viaje." & vbCr & vbCr & "Fecha/Hora: " & FechaHora_Formatted & vbCr & "Ruta: " & Ruta_DisplayName, vbExclamation, App.Title
        End If
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe un Viaje con el mismo Día, Hora y Ruta.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Viaje.Update", "Error al actualizar el Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta
        End If
    Else
        ShowErrorMessage "Classes.Viaje.Update", "Error al actualizar el Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta
    End If
End Function

Public Function UpdateAsientoOcupado() As Boolean
    If mrecData Is Nothing Then
        If Not OpenRecordset() Then
            Exit Function
        End If
    Else
        If mrecData.State <> adStateOpen Then
            If Not OpenRecordset() Then
                Exit Function
            End If
        End If
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    mrecData("AsientoOcupado").value = mAsientoOcupado
    mrecData.Update

    mIsCopy = False
    mIsDirty = False
    
    If Not mRefreshListSkip Then
        Call RefreshList
    End If
    
    Screen.MousePointer = vbDefault
    UpdateAsientoOcupado = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        ShowErrorMessage "Classes.Viaje.UpdateAsientoOcupado", "Error al actualizar la Cantidad de Asientos Libres del Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta
        If mrecData.EditMode = adEditAdd Or mrecData.EditMode = adEditInProgress Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.Viaje.UpdateAsientoOcupado", "Error al actualizar la Cantidad de Asientos Libres del Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta
    End If
End Function

Public Function Delete() As Boolean
    Dim cmdData As ADODB.command
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Viaje_Delete"
    cmdData.CommandType = adCmdStoredProc
    cmdData.NamedParameters = True
    cmdData.Parameters.Append cmdData.CreateParameter("@FechaHora", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDRuta", adChar, adParamInput, 20, mIDRuta)
    cmdData.Execute 0
    Set cmdData = Nothing
    
    '/////////////////////////////////////////////////////////////////
    'RESTO LOS KILOMETROS Y LOS VIAJES AL VEHICULO
    If mEstadoOriginal = VIAJE_ESTADO_FINALIZADO And mIDVehiculoOriginal > 0 Then
        VehiculoCambiarDatos mIDVehiculoOriginal, False
    End If
    
    If Not mRefreshListSkip Then
        RefreshList_RefreshCuentaCorriente 0
        RefreshList_RefreshViaje mFechaHoraOriginal, mIDRutaOriginal
    End If
        
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede eliminar el Viaje debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Viaje.Delete", "Error al eliminar el Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora & vbCr & "IDRuta: " & IDRuta
        End If
    Else
        ShowErrorMessage "Classes.Viaje.Delete", "Error al eliminar el Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora & vbCr & "IDRuta: " & IDRuta
    End If
End Function

Private Sub InitializeValues()
    mIDViaje = 0
    mRutaOtra = ""
    mIDPersona = 0
    mKilometro = 0
    mDuracion = 0
    mImporte = 0
    mImporteContado = 0
    mIDMedioPago = pParametro.MedioPago_Predeterminado_ID
    mCuotas = 1
    mOperacion = ""
    mIDCuentaCorrienteCaja = 0
    mCharter = False
    mIDVehiculo = 0
    mIDConductor = 0
    mAcreditaSueldo = True
    mIDConductor2 = 0
    mAcreditaSueldo2 = True
    mDiaSemanaBase = 0
    mEstado = VIAJE_ESTADO_ACTIVO
    mAsientoOcupado = 0
    mNotas = ""
    mPersonal = False
    mFechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioCreacion = 0
    mFechaHoraModificacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioModificacion = 0
    mFechaHoraEnProgreso = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioEnProgreso = ""
    mFechaHoraFinalizado = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioFinalizado = ""
    mFechaHoraCancelado = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioCancelado = ""
    
    mCuentaCorrienteAsignar = True

    mIsNew = True
    mIsCopy = False
    mIsDirty = False
End Sub

Private Sub Class_Initialize()
    InitializeValues
    mNoMatchRaiseError = True
    mRefreshListSkip = False
End Sub

Private Sub Class_Terminate()
    If Not mrecData Is Nothing Then
        If mrecData.State = adStateOpen Then
            mrecData.Close
        End If
        Set mrecData = Nothing
    End If
End Sub

Private Sub VehiculoCambiarDatos(ByVal IDVehiculoActualizar As Long, ByVal Incrementar As Boolean)
    Dim Vehiculo As Vehiculo
    
    Set Vehiculo = New Vehiculo
    With Vehiculo
        .IDVehiculo = IDVehiculoActualizar
        .NoMatchRaiseError = True
        .NotUpdateModifiedData = True
        If .Load() Then
            If Incrementar Then
                .KilometrajeEstimado = .KilometrajeEstimado + mKilometro
            Else
                .KilometrajeEstimado = .KilometrajeEstimado - mKilometroOriginal
            End If
            .Update
        End If
    End With
    Set Vehiculo = Nothing
End Sub

Public Function Asiento_Asignar() As Boolean
    Dim RefreshListSkipSave As Boolean
    
    'GRILLA DE ASIENTOS / LUGARES
    Dim DataGrid() As Long

    Dim cmdViajeDetalle_Asiento_List As ADODB.command
    Dim recViajeDetalle_Asiento_List As ADODB.Recordset
    
    'CAPACIDAD DEL VEHICULO
    Dim Vehiculo As Vehiculo
    Dim Ruta As Ruta
    Dim RutaDetalleCantidad As Long
    Dim RutaDetalleIndiceMinimo As Long
    Dim RutaDetalleIndiceMaximo As Long
    
    'CALCULOS
    Dim AsientoCurrent As Byte
    Dim AsientoAsignado As Boolean
    Dim LugarCurrent As Long
    Dim PasajerosParados As Long
    Dim IndiceOrigen As Long
    Dim IndiceDestino As Long
    
    'ERRORES
    Dim errorMessage As String
    Dim StartTime As Long
    Dim RandomDelayMilliseconds As Long
    Dim ErrorCounter As Long
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
RESTART:
    '//////////////////////////////////////////////////////////////////
    'OBTENGO LOS DATOS DEL VEHICULO
    If Not Load() Then
        Exit Function
    End If
    Set Vehiculo = New Vehiculo
    If mIDVehiculo <> 0 Then
        Vehiculo.IDVehiculo = mIDVehiculo
        If Not Vehiculo.Load() Then
            Set Vehiculo = Nothing
            Exit Function
        End If
    End If
    
    mAsientoOcupado = 0
    If Vehiculo.Asiento = 0 Then
        'EL VEHICULO NO ESPECIFICA LA CANTIDAD DE ASIENTOS
        'POR LO TANTO, ASIGNO ESTADO CONDICIONAL A TODOS
        'Y ASIENTO NULL
        errorMessage = "Error al cambiar el Estado."
        
        Set cmdViajeDetalle_Asiento_List = New ADODB.command
        Set cmdViajeDetalle_Asiento_List.ActiveConnection = pDatabase.Connection
        cmdViajeDetalle_Asiento_List.CommandText = "sp_ViajeDetalle_Asiento_List"
        cmdViajeDetalle_Asiento_List.CommandType = adCmdStoredProc
        cmdViajeDetalle_Asiento_List.Parameters.Append cmdViajeDetalle_Asiento_List.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , mFechaHora)
        cmdViajeDetalle_Asiento_List.Parameters.Append cmdViajeDetalle_Asiento_List.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, mIDRuta)
        cmdViajeDetalle_Asiento_List.Parameters.Append cmdViajeDetalle_Asiento_List.CreateParameter("OcupanteTipo_FILTER", adChar, adParamInput, 2, OCUPANTE_TIPO_PASAJERO)
        cmdViajeDetalle_Asiento_List.Parameters.Append cmdViajeDetalle_Asiento_List.CreateParameter("EstadoExclude_FILTER", adChar, adParamInput, 3, VIAJE_DETALLE_ESTADO_CANCELADO)
        Set recViajeDetalle_Asiento_List = New ADODB.Recordset
        recViajeDetalle_Asiento_List.Open cmdViajeDetalle_Asiento_List, , adOpenKeyset, adLockOptimistic
        recViajeDetalle_Asiento_List.Properties("Update Criteria").value = adCriteriaKey
        Set cmdViajeDetalle_Asiento_List = Nothing
        
        With recViajeDetalle_Asiento_List
            Do While Not .EOF
                If Not IsNull(.Fields("Asiento").value) Then
                    .Fields("Asiento").value = Null
                End If
                If .Fields("Estado").value & "" <> VIAJE_DETALLE_ESTADO_CONDICIONAL Then
                    .Fields("Estado").value = VIAJE_DETALLE_ESTADO_CONDICIONAL
                End If
                .Update
                
                .MoveNext
            Loop
        End With
    Else
        '//////////////////////////////////////////////////////////////////
        'OBTENGO LOS DATOS DEL DETALLE DE LA RUTA
        If mIDRuta <> pParametro.Ruta_ID_Otra Then
            Set Ruta = New Ruta
            Ruta.IDRuta = mIDRuta
            If Not Ruta.GetStatistics(RutaDetalleCantidad, RutaDetalleIndiceMinimo, RutaDetalleIndiceMaximo) Then
                Set Ruta = Nothing
                Exit Function
            End If
            Set Ruta = Nothing
        End If
        
        Set cmdViajeDetalle_Asiento_List = New ADODB.command
        Set cmdViajeDetalle_Asiento_List.ActiveConnection = pDatabase.Connection
        cmdViajeDetalle_Asiento_List.CommandText = "sp_ViajeDetalle_Asiento_List"
        cmdViajeDetalle_Asiento_List.CommandType = adCmdStoredProc
        cmdViajeDetalle_Asiento_List.Parameters.Append cmdViajeDetalle_Asiento_List.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , mFechaHora)
        cmdViajeDetalle_Asiento_List.Parameters.Append cmdViajeDetalle_Asiento_List.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, mIDRuta)
        cmdViajeDetalle_Asiento_List.Parameters.Append cmdViajeDetalle_Asiento_List.CreateParameter("OcupanteTipo_FILTER", adChar, adParamInput, 2, OCUPANTE_TIPO_PASAJERO)
        cmdViajeDetalle_Asiento_List.Parameters.Append cmdViajeDetalle_Asiento_List.CreateParameter("EstadoExclude_FILTER", adChar, adParamInput, 3, VIAJE_DETALLE_ESTADO_CANCELADO)
        Set recViajeDetalle_Asiento_List = New ADODB.Recordset
        recViajeDetalle_Asiento_List.Open cmdViajeDetalle_Asiento_List, , adOpenKeyset, adLockOptimistic
        recViajeDetalle_Asiento_List.Properties("Update Criteria").value = adCriteriaKey
        Set cmdViajeDetalle_Asiento_List = Nothing
        
        If Not recViajeDetalle_Asiento_List.EOF Then
            errorMessage = "Error al asignar el Asiento."
            
            'PREPARO EL TAMAÑO DEL ARRAY
            If mIDRuta = pParametro.Ruta_ID_Otra Then
                AsientoCurrent = 1
            Else
                ReDim DataGrid(1 To Vehiculo.Asiento, 0 To ((RutaDetalleCantidad - 1) * 2)) As Long
            End If
            
            With recViajeDetalle_Asiento_List
                Do While Not .EOF
                    If .Fields("Realizado").value = False Then
                        If IsNull(.Fields("Estado").value) Then
                            .Fields("Estado").value = VIAJE_DETALLE_ESTADO_CONFIRMADO
                        End If
                        If Not IsNull(.Fields("Asiento").value) Then
                            .Fields("Asiento").value = Null
                        End If
                        If .EditMode = adEditInProgress Then
                            .Update
                        End If
                    Else
                        If mIDRuta = pParametro.Ruta_ID_Otra Or mIDRuta = pParametro.Ruta_Paquete_ID Then
                            'SI ES UNA RUTA ESPECIAL, ASIGNO LOS ASIENTOS UNO A UNO
                            If AsientoCurrent <= Vehiculo.Asiento Then
                                If Val(.Fields("Asiento").value & "") <> AsientoCurrent Then
                                    .Fields("Asiento").value = AsientoCurrent
                                End If
                                If .Fields("Estado").value & "" <> VIAJE_DETALLE_ESTADO_CONFIRMADO Then
                                    .Fields("Estado").value = VIAJE_DETALLE_ESTADO_CONFIRMADO
                                End If
                                If .EditMode = adEditInProgress Then
                                    .Update
                                End If
                                AsientoAsignado = True
                                If mAsientoOcupado < AsientoCurrent Then
                                    mAsientoOcupado = AsientoCurrent
                                End If
                                AsientoCurrent = AsientoCurrent + 1
                            Else
                                AsientoAsignado = False
                            End If
                        Else
                            'ASIGNO LOS ASIENTOS EN BASE A DONDE SUBEN Y DONDE BAJAN
                            IndiceOrigen = ((.Fields("IndiceOrigen").value / 10) * 2) - 1
                            IndiceDestino = ((.Fields("IndiceDestino").value / 10) * 2) - 2
                            AsientoAsignado = False
                            For AsientoCurrent = 1 To Vehiculo.Asiento
                                If DataGrid(AsientoCurrent, 0) <> 0 Then
                                    'EL ASIENTO ESTA OCUPADO, VERIFICO SI TIENE LUGAR PARA LA RESERVA
                                    If DataGrid(AsientoCurrent, IndiceOrigen) = 0 And DataGrid(AsientoCurrent, IndiceDestino) = 0 Then
                                        'NI EL ORIGEN NI EL DESTINO ESTAN OCUPADOS, DEBO VERIFICAR ENTRE MEDIO
                                        For LugarCurrent = IndiceOrigen To IndiceDestino
                                            If DataGrid(AsientoCurrent, LugarCurrent) <> 0 Then
                                                Exit For
                                            End If
                                        Next LugarCurrent
                                        If LugarCurrent - 1 = IndiceDestino Then
                                            'SI NO SALIO ANTES PORQUE EL LUGAR ESTABA OCUPADO, LO ASIGNO
                                            If Val(.Fields("Asiento").value & "") <> AsientoCurrent Then
                                                .Fields("Asiento").value = AsientoCurrent
                                            End If
                                            If .Fields("Estado").value & "" <> VIAJE_DETALLE_ESTADO_CONFIRMADO Then
                                                .Fields("Estado").value = VIAJE_DETALLE_ESTADO_CONFIRMADO
                                            End If
                                            If .EditMode = adEditInProgress Then
                                                .Update
                                            End If
                                            AsientoAsignado = True
                                            If AsientoOcupado < AsientoCurrent Then
                                                AsientoOcupado = AsientoCurrent
                                            End If
                                            For LugarCurrent = IndiceOrigen To IndiceDestino
                                                DataGrid(AsientoCurrent, LugarCurrent) = .Fields("Indice").value
                                            Next LugarCurrent
                                            Exit For
                                        End If
                                    End If
                                Else
                                    'EL ASIENTO ESTA LIBRE, POR LO TANTO, ASIGNO EL VIAJE
                                    If Val(.Fields("Asiento").value & "") <> AsientoCurrent Then
                                        .Fields("Asiento").value = AsientoCurrent
                                    End If
                                    If .Fields("Estado").value & "" <> VIAJE_DETALLE_ESTADO_CONFIRMADO Then
                                        .Fields("Estado").value = VIAJE_DETALLE_ESTADO_CONFIRMADO
                                    End If
                                    If .EditMode = adEditInProgress Then
                                        .Update
                                    End If
                                    AsientoAsignado = True
                                    If mAsientoOcupado < AsientoCurrent Then
                                        mAsientoOcupado = AsientoCurrent
                                    End If
                                    DataGrid(AsientoCurrent, 0) = 1
                                    For LugarCurrent = IndiceOrigen To IndiceDestino
                                        DataGrid(AsientoCurrent, LugarCurrent) = .Fields("Indice").value
                                    Next LugarCurrent
                                    Exit For
                                End If
                            Next AsientoCurrent
                        End If
                            
                        If Not AsientoAsignado Then
                            'NO HAY ASIENTOS LIBRES PARA ESTA RESERVA
                            mAsientoExcedido = True
                            If PasajerosParados < (Vehiculo.Pasajero - Vehiculo.Asiento) Then
                                'COMO QUEDAN LUGARES PARADOS, LE ASIGNO CERO
                                PasajerosParados = PasajerosParados + 1
                                If IsNull(.Fields("Asiento").value) Or .Fields("Asiento").value <> 0 Then
                                    .Fields("Asiento").value = 0
                                End If
                                If .Fields("Estado").value & "" <> VIAJE_DETALLE_ESTADO_CONFIRMADO Then
                                    .Fields("Estado").value = VIAJE_DETALLE_ESTADO_CONFIRMADO
                                End If
                                If .EditMode = adEditInProgress Then
                                    .Update
                                End If
                            Else
                                'COMO NO HAY MAS LUGAR, LE ASIGNO NULL
                                If Not IsNull(.Fields("Asiento").value) Then
                                    .Fields("Asiento").value = Null
                                End If
                                If .Fields("Estado").value & "" <> VIAJE_DETALLE_ESTADO_CONDICIONAL Then
                                    .Fields("Estado").value = VIAJE_DETALLE_ESTADO_CONDICIONAL
                                End If
                                If .EditMode = adEditInProgress Then
                                    .Update
                                End If
                            End If
                        End If
                    End If
                    
                    .MoveNext
                Loop
            End With
        Else
            mAsientoOcupado = 0
        End If
        
        recViajeDetalle_Asiento_List.Close
        Set recViajeDetalle_Asiento_List = Nothing
    End If
    
    Set Vehiculo = Nothing
    
    Call UpdateAsientoOcupado
    
    RefreshListSkipSave = mRefreshListSkip
    mRefreshListSkip = True
    
    Do While ErrorCounter < 3
        If OrdenAndCuentaCorriente_Asignar Then
            Exit Do
        Else
            ErrorCounter = ErrorCounter + 1
            'HUBO ALGUN ERROR, PERO VOY A REINTENTAR 3 VECES
            'POR LO TANTO, HAGO UNA PAUSA AL AZAR POR SI PASA EN DOS PC,
            'PARA QUE NO SE QUEDEN EN UN DEAD LOCK INFINITO Y
            'REINICIO LA RUTINA
            
            If mrecData.EditMode = adEditInProgress Then
                mrecData.CancelUpdate
            End If
            
            Randomize Timer
            RandomDelayMilliseconds = Int(Rnd(1) * 2000) + 500
            
            StartTime = timeGetTime
            Do While (timeGetTime() - StartTime) < RandomDelayMilliseconds
                
            Loop
            
            mrecData.Requery
        End If
    Loop
    
    mRefreshListSkip = RefreshListSkipSave
    
    If Not mRefreshListSkip Then
        RefreshList_RefreshViaje FechaHora, IDRuta
        RefreshList_RefreshViajeDetalle FechaHora, IDRuta, 0
        RefreshList_RefreshCuentaCorriente 0
    End If
    
    Asiento_Asignar = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_VALUE_CHANGED_SINCE_LAST_READ Then
            'OTRA PC MODIFICO LOS DATOS DESPUES DE HABERLOS LEIDO EN ESTA PC,
            'POR LO TANTO, HAGO UNA PAUSA AL AZAR POR SI PASA EN LAS DOS PC,
            'PARA QUE NO SE QUEDEN EN UN DEAD LOCK INFINITO Y
            'REINICIO LA RUTINA
            Randomize Timer
            RandomDelayMilliseconds = Int(Rnd(1) * 2000) + 500
            
            StartTime = timeGetTime
            Do While (timeGetTime() - StartTime) < RandomDelayMilliseconds
                
            Loop
            
            Resume RESTART
        Else
            ShowErrorMessage "Classes.Viaje.Asiento_Asignar", errorMessage & vbCr & vbCr & "Fecha/Hora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta
        End If
    Else
        ShowErrorMessage "Classes.Viaje.Asiento_Asignar", errorMessage & vbCr & vbCr & "Fecha/Hora: " & Format(FechaHora, "Short Date") & " " & Format(FechaHora, "Short Time") & vbCr & "IDRuta: " & IDRuta
    End If
End Function

Public Function OrdenAndCuentaCorriente_Asignar() As Boolean
    Dim ViajeDetalle As ViajeDetalle
    Dim cmdViajeDetalle_List As ADODB.command
    Dim recViajeDetalle_List As ADODB.Recordset
    
    'CONFLICTOS
    Dim RandomDelayMilliseconds As Long
    Dim StartTime As Long
    
    'ORDEN
    Dim OrdenPersonaConfirmado As Long
    Dim OrdenPersonaCondicional As Long
    Dim OrdenPersonaCancelado As Long
    Dim OrdenComisionConfirmado As Long
    Dim OrdenComisionCancelado As Long
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set cmdViajeDetalle_List = New ADODB.command
    Set cmdViajeDetalle_List.ActiveConnection = pDatabase.Connection
    cmdViajeDetalle_List.CommandText = "sp_ViajeDetalle_List"
    cmdViajeDetalle_List.CommandType = adCmdStoredProc
    cmdViajeDetalle_List.Parameters.Append cmdViajeDetalle_List.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , mFechaHora)
    cmdViajeDetalle_List.Parameters.Append cmdViajeDetalle_List.CreateParameter("IDRuta_FILTER", adVarChar, adParamInput, 20, mIDRuta)
    Set recViajeDetalle_List = New ADODB.Recordset
    recViajeDetalle_List.Open cmdViajeDetalle_List, , adOpenKeyset, adLockOptimistic
    recViajeDetalle_List.Properties("Update Criteria").value = adCriteriaKey
    recViajeDetalle_List.Sort = "Prioridad, Indice"
    
RESTART:
    With recViajeDetalle_List
        Set ViajeDetalle = New ViajeDetalle
        Do While Not .EOF
            ViajeDetalle.RefreshListSkip = mRefreshListSkip
            ViajeDetalle.FechaHora = mFechaHora
            ViajeDetalle.IDRuta = mIDRuta
            ViajeDetalle.Indice = recViajeDetalle_List("Indice").value
            If ViajeDetalle.Load() Then
                '///////////////////////////////////////////////////////////////
                'ASIGNO EL ORDEN
                Select Case .Fields("OcupanteTipo").value
                    Case OCUPANTE_TIPO_COMISION
                        Select Case .Fields("Estado").value
                            Case VIAJE_DETALLE_ESTADO_CONFIRMADO
                                OrdenComisionConfirmado = OrdenComisionConfirmado + 1
                                If Val(.Fields("Orden").value & "") <> OrdenComisionConfirmado Then
                                    .Fields("Orden").value = OrdenComisionConfirmado
                                End If
                            Case VIAJE_DETALLE_ESTADO_CANCELADO
                                OrdenComisionCancelado = OrdenComisionCancelado + 1
                                If Val(.Fields("Orden").value & "") <> OrdenComisionCancelado Then
                                    .Fields("Orden").value = OrdenComisionCancelado
                                End If
                        End Select
                    Case OCUPANTE_TIPO_PASAJERO
                        Select Case .Fields("Estado").value
                            Case VIAJE_DETALLE_ESTADO_CONFIRMADO
                                OrdenPersonaConfirmado = OrdenPersonaConfirmado + 1
                                If Val(.Fields("Orden").value & "") <> OrdenPersonaConfirmado Then
                                    .Fields("Orden").value = OrdenPersonaConfirmado
                                End If
                            Case VIAJE_DETALLE_ESTADO_CONDICIONAL
                                OrdenPersonaCondicional = OrdenPersonaCondicional + 1
                                If Val(.Fields("Orden").value & "") <> OrdenPersonaCondicional Then
                                    .Fields("Orden").value = OrdenPersonaCondicional
                                End If
                            Case VIAJE_DETALLE_ESTADO_CANCELADO
                                OrdenPersonaCancelado = OrdenPersonaCancelado + 1
                                If Val(.Fields("Orden").value & "") <> OrdenPersonaCancelado Then
                                    .Fields("Orden").value = OrdenPersonaCancelado
                                End If
                        End Select
                End Select
                If .EditMode = adEditInProgress Then
                    .Update
                End If
                
                '///////////////////////////////////////////////////////////////
                'ASIGNO LOS MOVIMIENTOS DE CUENTA CORRIENTE
                If mCuentaCorrienteAsignar Then
                    If Not ViajeDetalle.CuentaCorriente_Asignar(mEstado) Then
                        Exit Function
                    End If
                End If
            End If
            
            .MoveNext
        Loop
        Set ViajeDetalle = Nothing
    End With
    
    recViajeDetalle_List.Close
    Set recViajeDetalle_List = Nothing
    Set cmdViajeDetalle_List = Nothing
    Set ViajeDetalle = Nothing
    
    If Not mRefreshListSkip Then
        RefreshList_RefreshViajeDetalle mFechaHora, mIDRuta, 0
        RefreshList_RefreshCuentaCorriente 0
    End If
    
    OrdenAndCuentaCorriente_Asignar = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_VALUE_CHANGED_SINCE_LAST_READ Then
            'OTRA PC MODIFICO LOS DATOS DESPUES DE HABERLOS LEIDO EN ESTA,
            'POR LO TANTO, HAGO UNA PAUSA AL AZAR POR SI PASA EN LAS DOS PC,
            'PARA QUE NO SE QUEDEN EN UN DEAD LOCK INFINITO Y
            'REINICIO LA RUTINA
            Randomize Timer
            RandomDelayMilliseconds = Int(Rnd(1) * 2000) + 500
            
            StartTime = timeGetTime
            Do While (timeGetTime() - StartTime) < RandomDelayMilliseconds
                
            Loop
            
            If recViajeDetalle_List.EditMode = adEditInProgress Then
                recViajeDetalle_List.CancelUpdate
            End If
            recViajeDetalle_List.Requery
        
            Resume RESTART
        Else
            Set recViajeDetalle_List = Nothing
            Set cmdViajeDetalle_List = Nothing
            ShowErrorMessage "Classes.Viaje.OrdenAndCuentaCorriente_Asignar", "Error al asignar el Orden del Detalle del Viaje." & vbCr & vbCr & "Fecha/Hora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta
        End If
    Else
        Set recViajeDetalle_List = Nothing
        Set cmdViajeDetalle_List = Nothing
        ShowErrorMessage "Classes.Viaje.OrdenAndCuentaCorriente_Asignar", "Error al asignar el Orden del Detalle del Viaje." & vbCr & vbCr & "Fecha/Hora: " & FechaHora & vbCr & "IDRuta: " & IDRuta
    End If
End Function

Public Function CuentaCorriente_Asignar() As Boolean

    Screen.MousePointer = vbHourglass
    
    'CREDITO
    If mImporteContado = 0 Then
        If Not CuentaCorriente_Asignar_Movimiento_Eliminar(pParametro.CuentaCorrienteGrupo_ID_ViajeCredito) Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    Else
        If Not CuentaCorriente_Asignar_Movimiento_Crear(pParametro.CuentaCorrienteGrupo_ID_ViajeCredito, mIDCuentaCorrienteCaja, mIDPersona, "Viaje Especial: " & FechaHora_Formatted & " - " & RutaOtra, mImporteContado, mIDMedioPago, mCuotas, mOperacion) Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
        
    'DEBITO
    If mImporte = 0 Or mEstado = VIAJE_ESTADO_ACTIVO Or mEstado = VIAJE_ESTADO_CANCELADO Then
        If Not CuentaCorriente_Asignar_Movimiento_Eliminar(pParametro.CuentaCorrienteGrupo_ID_ViajeDebito) Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    Else
        If Not CuentaCorriente_Asignar_Movimiento_Crear(pParametro.CuentaCorrienteGrupo_ID_ViajeDebito, pParametro.CuentaCorrienteCaja_ID_ViajeDebito, mIDPersona, "Viaje Especial: " & FechaHora_Formatted & " - " & RutaOtra, mImporte * -1, pParametro.MedioPago_Predeterminado_ID, 1, "") Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    If Not mRefreshListSkip Then
        RefreshList
    End If
    
    Screen.MousePointer = vbDefault
    CuentaCorriente_Asignar = True
End Function

Public Function CuentaCorriente_Asignar_Sueldo() As Boolean
    Dim Ruta As Ruta
    Dim Horario As Horario
    Dim ConductorRuta As ConductorRuta
    
    Dim ConductorImporteTramoCompleto As Currency
    Dim ConductorImporteTramo1 As Currency
    Dim ConductorImporteTramo2 As Currency
    
    Dim CrearMovimiento As Boolean
    Dim Permite2Conductores As Boolean
    
    Screen.MousePointer = vbHourglass
    
    'ELIMINO EL SUELDO ACTUAL DEL CONDUCTOR 1 SI NO CORRESPONDE
    If (Not mAcreditaSueldo) Or mEstado = VIAJE_ESTADO_ACTIVO Or mEstado = VIAJE_ESTADO_CANCELADO Or mIDConductor = 0 Then
        If Not CuentaCorriente_Asignar_Movimiento_Eliminar(pParametro.CuentaCorrienteGrupo_ID_Sueldo, 1) Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    Else
        CrearMovimiento = True
    End If
    
    'CARGO LOS DATOS DE LA RUTA
    Set Ruta = New Ruta
    Ruta.IDRuta = IDRuta
    If Not Ruta.Load() Then
        Screen.MousePointer = vbDefault
        Set Ruta = Nothing
        Exit Function
    End If
    Permite2Conductores = pParametro.Viaje_Permite_2_Conductores And Ruta.Permite2Conductores
    
    'ELIMINO EL SUELDO ACTUAL DEL CONDUCTOR 2 SI NO CORRESPONDE
    If Permite2Conductores Then
        If (Not mAcreditaSueldo2) Or mEstado = VIAJE_ESTADO_ACTIVO Or mEstado = VIAJE_ESTADO_CANCELADO Then
            If Not CuentaCorriente_Asignar_Movimiento_Eliminar(pParametro.CuentaCorrienteGrupo_ID_Sueldo, 2) Then
                Screen.MousePointer = vbDefault
                Set Ruta = Nothing
                Exit Function
            End If
        Else
            CrearMovimiento = True
        End If
    End If
    
    
    If CrearMovimiento Then
        'PONGO TODO EN CERO
        ConductorImporteTramoCompleto = -1
        ConductorImporteTramo1 = -1
        ConductorImporteTramo2 = -1
        
        'VERIFICO EL IMPORTE QUE FIGURA EN LA RUTA
        If Not Ruta.NoMatch Then
            If Ruta.ConductorImporteTramoCompleto <> -1 Then
                ConductorImporteTramoCompleto = Ruta.ConductorImporteTramoCompleto
            End If
            If Permite2Conductores Then
                If Ruta.ConductorImporteTramo1 <> -1 Then
                    ConductorImporteTramo1 = Ruta.ConductorImporteTramo1
                End If
                If Ruta.ConductorImporteTramo2 <> -1 Then
                    ConductorImporteTramo2 = Ruta.ConductorImporteTramo2
                End If
            End If
        End If
        Set Ruta = Nothing
        
        'VERIFICO EL IMPORTE EN EL HORARIO
        Set Horario = New Horario
        Horario.DiaSemana = FechaHora_Weekday
        Horario.Hora = FechaHora_FormattedAsTime
        Horario.IDRuta = IDRuta
        Horario.NoMatchRaiseError = False
        If Not Horario.Load() Then
            Screen.MousePointer = vbDefault
            Set Horario = Nothing
            Exit Function
        End If
        If Not Horario.NoMatch Then
            If Horario.ConductorImporteTramoCompleto <> -1 Then
                ConductorImporteTramoCompleto = Horario.ConductorImporteTramoCompleto
            End If
            If Permite2Conductores Then
                If Horario.ConductorImporteTramo1 <> -1 Then
                    ConductorImporteTramo1 = Horario.ConductorImporteTramo1
                End If
                If Horario.ConductorImporteTramo2 <> -1 Then
                    ConductorImporteTramo2 = Horario.ConductorImporteTramo2
                End If
            End If
        End If
        Set Horario = Nothing
        
        'VERIFICO EL IMPORTE EN CONDUCTOR 1 / RUTA
        Set ConductorRuta = New ConductorRuta
        ConductorRuta.NoMatchRaiseError = False
        ConductorRuta.IDPersona = mIDConductor
        ConductorRuta.IDRuta = mIDRuta
        If Not ConductorRuta.Load() Then
            Screen.MousePointer = vbDefault
            Set ConductorRuta = Nothing
            Exit Function
        End If
        If Not ConductorRuta.NoMatch Then
            If ConductorRuta.ConductorImporteTramoCompleto <> -1 Then
                ConductorImporteTramoCompleto = ConductorRuta.ConductorImporteTramoCompleto
            End If
            If Permite2Conductores Then
                If ConductorRuta.ConductorImporteTramo1 <> -1 Then
                    ConductorImporteTramo1 = ConductorRuta.ConductorImporteTramo1
                End If
            End If
        End If
        Set ConductorRuta = Nothing
        
        'VERIFICO EL IMPORTE EN CONDUCTOR 2 / RUTA
        If Permite2Conductores Then
            Set ConductorRuta = New ConductorRuta
            ConductorRuta.NoMatchRaiseError = False
            ConductorRuta.IDPersona = mIDConductor2
            ConductorRuta.IDRuta = mIDRuta
            If Not ConductorRuta.Load() Then
                Screen.MousePointer = vbDefault
                Set ConductorRuta = Nothing
                Exit Function
            End If
            If Not ConductorRuta.NoMatch Then
                If ConductorRuta.ConductorImporteTramo2 <> -1 Then
                    ConductorImporteTramo2 = ConductorRuta.ConductorImporteTramo2
                End If
            End If
            Set ConductorRuta = Nothing
        End If
                
        'CREO LOS MOVIMIENTOS
        If Permite2Conductores And mIDConductor2 > 0 Then
            'HAY 2 CONDUCTORES, ASIGNO LOS IMPORTES POR CADA TRAMO
            If mAcreditaSueldo And ConductorImporteTramo1 <> -1 Then
                If Not CuentaCorriente_Asignar_Movimiento_Crear(pParametro.CuentaCorrienteGrupo_ID_Sueldo, pParametro.CuentaCorrienteCaja_ID_ViajeDebito, mIDConductor, "Sueldo por Viaje (Tramo 1): " & FechaHora_Formatted & " - " & IDRuta, ConductorImporteTramo1, pParametro.MedioPago_Predeterminado_ID, 1, "", 1) Then
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
            End If
            If mAcreditaSueldo2 And ConductorImporteTramo2 <> -1 Then
                If Not CuentaCorriente_Asignar_Movimiento_Crear(pParametro.CuentaCorrienteGrupo_ID_Sueldo, pParametro.CuentaCorrienteCaja_ID_ViajeDebito, mIDConductor2, "Sueldo por Viaje (Tramo 2): " & FechaHora_Formatted & " - " & IDRuta, ConductorImporteTramo2, pParametro.MedioPago_Predeterminado_ID, 1, "", 2) Then
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
            End If
        Else
            'HAY 1 CONDUCTOR, ASIGNO EL IMPORTE DEL TRAMO COMPLETO
            If ConductorImporteTramoCompleto <> -1 Then
                If Not CuentaCorriente_Asignar_Movimiento_Crear(pParametro.CuentaCorrienteGrupo_ID_Sueldo, pParametro.CuentaCorrienteCaja_ID_ViajeDebito, mIDConductor, "Sueldo por Viaje" & IIf(pParametro.Viaje_Permite_2_Conductores, " (Tramo Completo)", "") & ": " & FechaHora_Formatted & " - " & IDRuta, ConductorImporteTramoCompleto, pParametro.MedioPago_Predeterminado_ID, 1, "", 1) Then
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
            End If
        End If
    Else
        Set Ruta = Nothing
    End If
    
    If Not mRefreshListSkip Then
        RefreshList
    End If
    
    Screen.MousePointer = vbDefault
    CuentaCorriente_Asignar_Sueldo = True
End Function

Private Function CuentaCorriente_Asignar_Movimiento_Crear(ByVal IDGrupo As Long, ByVal IDCaja As Long, ByVal IDPersonaCC As Long, ByVal Descripcion As String, ByVal ImporteCC As Currency, ByVal IDMedioPagoMov As Byte, ByVal CuotasMov As Byte, ByVal OperacionMov As String, Optional ConductorNumero As Byte = 0) As Boolean
    Dim CuentaCorriente As CuentaCorriente
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set CuentaCorriente = New CuentaCorriente
    CuentaCorriente.IDCuentaCorrienteGrupo = IDGrupo
    CuentaCorriente.Viaje_FechaHora = mFechaHora
    CuentaCorriente.Viaje_IDRuta = mIDRuta
    CuentaCorriente.Viaje_ConductorNumero = ConductorNumero
    CuentaCorriente.NoMatchRaiseError = False
    CuentaCorriente.RefreshListSkip = mRefreshListSkip
    If CuentaCorriente.LoadByViaje() Then
        CuentaCorriente.IDCuentaCorrienteGrupo = IDGrupo
        CuentaCorriente.IDCuentaCorrienteCaja = IDCaja
        CuentaCorriente.IDPersona = IDPersonaCC
        CuentaCorriente.IDPersonaOrigen = 0
        If CuentaCorriente.NoMatch Then
            CuentaCorriente.FechaHora = Now
        End If
        CuentaCorriente.Descripcion = Descripcion
        CuentaCorriente.Importe = ImporteCC
        CuentaCorriente.IDMedioPago = IDMedioPagoMov
        CuentaCorriente.Cuotas = CuotasMov
        CuentaCorriente.Operacion = OperacionMov
        CuentaCorriente.Viaje_FechaHora = mFechaHora
        CuentaCorriente.Viaje_IDRuta = mIDRuta
        CuentaCorriente.Viaje_ConductorNumero = ConductorNumero
        CuentaCorriente.Update
    End If
    Set CuentaCorriente = Nothing
        
    Screen.MousePointer = vbDefault
    CuentaCorriente_Asignar_Movimiento_Crear = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Viaje.CuentaCorriente_Asignar_Movimiento_Crear", "Error al Crear el Movimiento en la Cuenta Corriente." & vbCr & vbCr & "Fecha/Hora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta
End Function

Private Function CuentaCorriente_Asignar_Movimiento_Eliminar(ByVal IDGrupo As Long, Optional ConductorNumero As Byte = 0) As Boolean
    Dim CuentaCorriente As CuentaCorriente
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set CuentaCorriente = New CuentaCorriente
    With CuentaCorriente
        .IDCuentaCorrienteGrupo = IDGrupo
        .Viaje_FechaHora = mFechaHora
        .Viaje_IDRuta = mIDRuta
        .Viaje_ConductorNumero = ConductorNumero
        .NoMatchRaiseError = False
        .RefreshListSkip = mRefreshListSkip
        If .LoadByViaje() Then
            If Not .NoMatch Then
                If Not .Delete() Then
                    Set CuentaCorriente = Nothing
                    Exit Function
                End If
            End If
        End If
    End With
    Set CuentaCorriente = Nothing
        
    Screen.MousePointer = vbDefault
    CuentaCorriente_Asignar_Movimiento_Eliminar = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Viaje.CuentaCorriente_Asignar_Movimiento_Eliminar", "Error al Eliminar el Movimiento en la Cuenta Corriente." & vbCr & vbCr & "Fecha/Hora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta
End Function

'============================================================================
' Genera todas las Reservas Fijas de los Pasajeros, dado un Viaje
'============================================================================
Public Function GenerarReservasFijas() As Boolean
    Dim ViajeDetalle As ViajeDetalle
    
    Dim cmdPersona As ADODB.command
    Dim recPersona As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdPersona = New ADODB.command
    Set cmdPersona.ActiveConnection = pDatabase.Connection
    cmdPersona.CommandText = "sp_PersonaHorario_ByHorario"
    cmdPersona.CommandType = adCmdStoredProc
    cmdPersona.Parameters.Append cmdPersona.CreateParameter("DiaSemanaBase_FILTER", adTinyInt, adParamInput, , mDiaSemanaBase)
    cmdPersona.Parameters.Append cmdPersona.CreateParameter("Hora_FILTER", adChar, adParamInput, 8, Format(mFechaHora, "hh:nn:ss"))
    cmdPersona.Parameters.Append cmdPersona.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, mIDRuta)
    cmdPersona.Parameters.Append cmdPersona.CreateParameter("Fecha_FILTER", adChar, adParamInput, 10, Format(mFechaHora, "yyyy/mm/dd"))
    Set recPersona = New ADODB.Recordset
    recPersona.Open cmdPersona, , adOpenForwardOnly, adLockReadOnly
    Set cmdPersona = Nothing
    
    If Not recPersona.EOF Then
        Set ViajeDetalle = New ViajeDetalle
        Do While Not recPersona.EOF
            'Si no existe la Reserva, la genero
            ViajeDetalle.RefreshListSkip = True
            ViajeDetalle.NoMatchRaiseError = False
            ViajeDetalle.CallAsientoAsignarOnAddNew = False
            ViajeDetalle.VerifyDuplicates = False
            ViajeDetalle.GenerandoReservasFijas = True
            
            ViajeDetalle.IDViaje = mIDViaje
            ViajeDetalle.FechaHora = mFechaHora
            ViajeDetalle.IDRuta = mIDRuta
            ViajeDetalle.OcupanteTipo = OCUPANTE_TIPO_PASAJERO
            ViajeDetalle.IDPersona = recPersona("IDPersona").value
            ViajeDetalle.ReservaTipo = VIAJE_DETALLE_RESERVA_TIPO_FIJA
            If ViajeDetalle.LoadByReservaTipo() Then
                If ViajeDetalle.NoMatch Then
                    If ViajeDetalle.GetNewValues(mDiaSemanaBase) Then
                        ViajeDetalle.CreadoEnProgreso = False
                        ViajeDetalle.ReservadoPor = "Reserva Fija Automática"
                        ViajeDetalle.Update
                    End If
                End If
            End If
            
            recPersona.MoveNext
        Loop
        Set ViajeDetalle = Nothing
        
        Call Asiento_Asignar
    End If
    recPersona.Close
    Set recPersona = Nothing
    
    Set ViajeDetalle = Nothing
    
    If Not mRefreshListSkip Then
        RefreshList_RefreshViaje mFechaHora, mIDRuta
        RefreshList
    End If
    
    Screen.MousePointer = vbDefault
    GenerarReservasFijas = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Viaje.GenerarReservasFijas", "Error al Generar las Reservas Fijas del Viaje." & vbCr & vbCr & "Fecha/Hora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta
End Function

Public Function Asiento_Asignar_GetAsiento(ByVal IndiceActual As Long) As Long
    'GRILLA DE ASIENTOS / LUGARES
    Dim DataGrid() As Long

    Dim cmdViajeDetalle_Asiento_List As ADODB.command
    Dim recViajeDetalle_Asiento_List As ADODB.Recordset

    'CAPACIDAD DEL VEHICULO
    Dim Vehiculo As Vehiculo
    Dim Ruta As Ruta
    Dim RutaDetalle As RutaDetalle
    Dim RutaDetalleCantidad As Long
    Dim RutaDetalleIndiceMinimo As Long
    Dim RutaDetalleIndiceMaximo As Long

    'CALCULOS
    Dim AsientoCurrent As Long
    Dim AsientoAsignado As Boolean
    Dim LugarCurrent As Long
    Dim PasajerosParados As Long
    Dim IndiceOrigen As Long
    Dim IndiceDestino As Long

    'ERRORES
    Dim errorMessage As String
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    '//////////////////////////////////////////////////////////////////
    'OBTENGO LOS DATOS DEL VEHICULO
    Set Vehiculo = New Vehiculo
    If Not Load() Then
        Set Vehiculo = Nothing
        Asiento_Asignar_GetAsiento = -3
        Exit Function
    End If
    If IDVehiculo <> 0 Then
        Vehiculo.IDVehiculo = mIDVehiculo
        If Not Vehiculo.Load() Then
            Set Vehiculo = Nothing
            Asiento_Asignar_GetAsiento = -3
            Exit Function
        End If
    End If
    
    If Vehiculo.Asiento = 0 Then
        'EL VEHICULO NO ESPECIFICA LA CANTIDAD DE ASIENTOS
        'POR LO TANTO, ASIGNO ESTADO CONDICIONAL A TODOS
        Asiento_Asignar_GetAsiento = -2
    Else
        '//////////////////////////////////////////////////////////////////
        'OBTENGO LOS DATOS DEL DETALLE DE LA RUTA
        If mIDRuta <> pParametro.Ruta_ID_Otra Then
            Set Ruta = New Ruta
            Ruta.IDRuta = mIDRuta
            If Not Ruta.GetStatistics(RutaDetalleCantidad, RutaDetalleIndiceMinimo, RutaDetalleIndiceMaximo) Then
                Set Ruta = Nothing
                Asiento_Asignar_GetAsiento = -3
                Exit Function
            End If
            Set Ruta = Nothing
        End If
        
        Set cmdViajeDetalle_Asiento_List = New ADODB.command
        Set cmdViajeDetalle_Asiento_List.ActiveConnection = pDatabase.Connection
        cmdViajeDetalle_Asiento_List.CommandText = "sp_ViajeDetalle_Asiento_List"
        cmdViajeDetalle_Asiento_List.CommandType = adCmdStoredProc
        cmdViajeDetalle_Asiento_List.Parameters.Append cmdViajeDetalle_Asiento_List.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , mFechaHora)
        cmdViajeDetalle_Asiento_List.Parameters.Append cmdViajeDetalle_Asiento_List.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, mIDRuta)
        cmdViajeDetalle_Asiento_List.Parameters.Append cmdViajeDetalle_Asiento_List.CreateParameter("OcupanteTipo_FILTER", adChar, adParamInput, 2, OCUPANTE_TIPO_PASAJERO)
        cmdViajeDetalle_Asiento_List.Parameters.Append cmdViajeDetalle_Asiento_List.CreateParameter("EstadoExclude_FILTER", adChar, adParamInput, 3, VIAJE_DETALLE_ESTADO_CANCELADO)
        Set recViajeDetalle_Asiento_List = New ADODB.Recordset
        recViajeDetalle_Asiento_List.Open cmdViajeDetalle_Asiento_List, , adOpenForwardOnly, adLockReadOnly
        Set cmdViajeDetalle_Asiento_List = Nothing
        
        If Not recViajeDetalle_Asiento_List.EOF Then
            errorMessage = "Error al asignar el Asiento."
                        
            'PREPARO EL TAMAÑO DEL ARRAY
            If mIDRuta = pParametro.Ruta_ID_Otra Then
                AsientoCurrent = 1
            Else
                ReDim DataGrid(1 To Vehiculo.Asiento, 0 To ((RutaDetalleCantidad - 1) * 2)) As Long
            End If
        
            With recViajeDetalle_Asiento_List
                Do While Not .EOF
                    If .Fields("Realizado").value = False Then
                    Else
                        If mIDRuta = pParametro.Ruta_ID_Otra Then
                            'SI ES UNA RUTA ESPECIAL, ASIGNO LOS ASIENTOS UNO A UNO
                            If AsientoCurrent <= Vehiculo.Asiento Then
                                AsientoAsignado = True
                                AsientoCurrent = AsientoCurrent + 1
                            Else
                                AsientoAsignado = False
                            End If
                        Else
                            'ASIGNO LOS ASIENTOS EN BASE A DONDE SUBEN Y DONDE BAJAN
                            If IndiceActual > 0 And IndiceActual = .Fields("Indice").value Then
                            Else
                                IndiceOrigen = ((.Fields("IndiceOrigen").value / 10) * 2) - 1
                                IndiceDestino = ((.Fields("IndiceDestino").value / 10) * 2) - 2
                                AsientoAsignado = False
                                For AsientoCurrent = 1 To Vehiculo.Asiento
                                    If DataGrid(AsientoCurrent, 0) <> 0 Then
                                        'EL ASIENTO ESTA OCUPADO, VERIFICO SI TIENE LUGAR PARA LA RESERVA
                                        If DataGrid(AsientoCurrent, IndiceOrigen) = 0 And DataGrid(AsientoCurrent, IndiceDestino) = 0 Then
                                            'NI EL ORIGEN NI EL DESTINO ESTAN OCUPADOS, DEBO VERIFICAR ENTRE MEDIO
                                            For LugarCurrent = IndiceOrigen To IndiceDestino
                                                If DataGrid(AsientoCurrent, LugarCurrent) <> 0 Then
                                                    Exit For
                                                End If
                                            Next LugarCurrent
                                            If LugarCurrent - 1 = IndiceDestino Then
                                                'SI NO SALIO ANTES PORQUE EL LUGAR ESTABA OCUPADO, LO ASIGNO
                                                AsientoAsignado = True
                                                For LugarCurrent = IndiceOrigen To IndiceDestino
                                                    DataGrid(AsientoCurrent, LugarCurrent) = .Fields("Indice").value
                                                Next LugarCurrent
                                                Exit For
                                            End If
                                        End If
                                    Else
                                        'EL SIENTO ESTA LIBRE, POR LO TANTO, ASIGNO EL VIAJE
                                        AsientoAsignado = True
                                        DataGrid(AsientoCurrent, 0) = 1
                                        For LugarCurrent = IndiceOrigen To IndiceDestino
                                            DataGrid(AsientoCurrent, LugarCurrent) = .Fields("Indice").value
                                        Next LugarCurrent
                                        Exit For
                                    End If
                                Next AsientoCurrent
                            End If
                        End If
                        
                        If Not AsientoAsignado Then
                            'NO HAY ASIENTOS LIBRES PARA ESTA RESERVA
                            If PasajerosParados < (Vehiculo.Pasajero - Vehiculo.Asiento) Then
                                'COMO QUEDAN LUGARES PARADOS, LE ASIGNO CERO
                                PasajerosParados = PasajerosParados + 1
                            Else
                                'COMO NO HAY MAS LUGAR, LE ASIGNO NULL
                            End If
                        End If
                    End If
                    
                    .MoveNext
                Loop
            End With
            
            'AHORA VERIFICO SI HAY LUGAR PARA ESTA NUEVA RESERVA
            If mIDRuta = pParametro.Ruta_ID_Otra Then
                'SI ES UNA RUTA ESPECIAL, ASIGNO LOS ASIENTOS UNO A UNO
                If AsientoCurrent <= Vehiculo.Asiento Then
                    AsientoAsignado = True
                    Asiento_Asignar_GetAsiento = AsientoCurrent
                Else
                    AsientoAsignado = False
                End If
            Else
                'ASIGNO LOS ASIENTOS EN BASE A DONDE SUBEN Y DONDE BAJAN
                Set RutaDetalle = New RutaDetalle
                RutaDetalle.IDRuta = mIDRuta
                RutaDetalle.IDLugar = mIDOrigen
                If Not RutaDetalle.Load() Then
                    Set RutaDetalle = Nothing
                    Set Vehiculo = Nothing
                    Asiento_Asignar_GetAsiento = -3
                    Exit Function
                End If
                IndiceOrigen = ((RutaDetalle.Indice / 10) * 2) - 1
                Set RutaDetalle = Nothing
                
                Set RutaDetalle = New RutaDetalle
                RutaDetalle.IDRuta = mIDRuta
                RutaDetalle.IDLugar = mIDDestino
                If Not RutaDetalle.Load() Then
                    Set RutaDetalle = Nothing
                    Set Vehiculo = Nothing
                    Asiento_Asignar_GetAsiento = -3
                    Exit Function
                End If
                IndiceDestino = ((RutaDetalle.Indice / 10) * 2) - 2
                Set RutaDetalle = Nothing
                
                AsientoAsignado = False
                For AsientoCurrent = 1 To Vehiculo.Asiento
                    If DataGrid(AsientoCurrent, 0) <> 0 Then
                        'EL ASIENTO ESTA OCUPADO, VERIFICO SI TIENE LUGAR PARA LA RESERVA
                        If DataGrid(AsientoCurrent, IndiceOrigen) = 0 And DataGrid(AsientoCurrent, IndiceDestino) = 0 Then
                            'NI EL ORIGEN NI EL DESTINO ESTAN OCUPADOS, DEBO VERIFICAR ENTRE MEDIO
                            For LugarCurrent = IndiceOrigen To IndiceDestino
                                If DataGrid(AsientoCurrent, LugarCurrent) <> 0 Then
                                    Exit For
                                End If
                            Next LugarCurrent
                            If LugarCurrent - 1 = IndiceDestino Then
                                'SI NO SALIO ANTES PORQUE EL LUGAR ESTABA OCUPADO, LO ASIGNO
                                Asiento_Asignar_GetAsiento = AsientoCurrent
                                AsientoAsignado = True
                                Exit For
                            End If
                        End If
                    Else
                        'EL SIENTO ESTA LIBRE, POR LO TANTO, ASIGNO EL VIAJE
                        AsientoAsignado = True
                        Asiento_Asignar_GetAsiento = AsientoCurrent
                        Exit For
                    End If
                Next AsientoCurrent
            End If
                        
            If Not AsientoAsignado Then
                'NO HAY ASIENTOS LIBRES PARA ESTA RESERVA
                If PasajerosParados < (Vehiculo.Pasajero - Vehiculo.Asiento) Then
                    'COMO QUEDAN LUGARES PARADOS, LE ASIGNO CERO
                    PasajerosParados = PasajerosParados + 1
                    Asiento_Asignar_GetAsiento = 0
                Else
                    'COMO NO HAY MAS LUGAR, LE ASIGNO -1
                    Asiento_Asignar_GetAsiento = -1
                End If
            End If
        Else
            Asiento_Asignar_GetAsiento = 1
        End If
        
        recViajeDetalle_Asiento_List.Close
        Set recViajeDetalle_Asiento_List = Nothing
    End If
    
    Set Vehiculo = Nothing
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Viaje.Asiento_Asignar_GetAsiento", errorMessage & vbCr & vbCr & "Fecha/Hora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta
End Function

Public Function GetDetalleCountByTipoEstado(ByVal OcupanteTipo As String, ByVal EstadoFilter As String) As Long
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ViajeDetalle_ListByTipoEstado"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, mIDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("OcupanteTipo_FILTER", adChar, adParamInput, 2, OcupanteTipo)
    cmdData.Parameters.Append cmdData.CreateParameter("Estado_FILTER", adChar, adParamInput, 3, EstadoFilter)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenKeyset, adLockOptimistic
    Set cmdData = Nothing
    
    If Not recData.EOF Then
        recData.MoveLast
        GetDetalleCountByTipoEstado = recData.RecordCount
    End If

    recData.Close
    Set recData = Nothing
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Viaje.GetDetalleCountByTipoEstado", "Error al obtener la Cantidad de Items de Tipo ." & IIf(OcupanteTipo = OCUPANTE_TIPO_PASAJERO, "Pasajero.", "Comisión.") & vbCr & vbCr & "Fecha/Hora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta
End Function

Public Function EstadoVencidoCheck() As Long
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandType = adCmdStoredProc
    cmdData.CommandText = "sp_Report_Viaje_EstadoVencido_List"
    cmdData.Parameters.Append cmdData.CreateParameter("Personal", adBoolean, adParamInput, , pPersonal)

    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenStatic, adLockReadOnly
    Set cmdData = Nothing
    
    If Not recData.EOF Then
        recData.MoveLast
        
        EstadoVencidoCheck = recData.RecordCount
    End If
    
    recData.Close
    Set recData = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Viaje.EstadoVencidoCheck", "Error al verificar los Viajes con Estado vencido."
End Function

Private Sub VerificarPrepagosOtrasReservas()
    Dim cmdViajeDetalle_List As ADODB.command
    Dim recViajeDetalle_List As ADODB.Recordset
    Dim ViajeDetalle As ViajeDetalle
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass
    
    Set cmdViajeDetalle_List = New ADODB.command
    Set cmdViajeDetalle_List.ActiveConnection = pDatabase.Connection
    cmdViajeDetalle_List.CommandText = "usp_PersonaPrepago_GetViajeDetalleConListaPrepagoList"
    cmdViajeDetalle_List.CommandType = adCmdStoredProc
    cmdViajeDetalle_List.Parameters.Append cmdViajeDetalle_List.CreateParameter("FechaHora", adDate, adParamInput, , mFechaHora)
    cmdViajeDetalle_List.Parameters.Append cmdViajeDetalle_List.CreateParameter("IDRuta", adChar, adParamInput, 20, mIDRuta)
    Set recViajeDetalle_List = New ADODB.Recordset
    recViajeDetalle_List.Open cmdViajeDetalle_List, , adOpenForwardOnly, adLockReadOnly
    Set cmdViajeDetalle_List = Nothing
    
    Do While Not recViajeDetalle_List.EOF
        Set ViajeDetalle = New ViajeDetalle
        ViajeDetalle.FechaHora = recViajeDetalle_List("FechaHora").value
        ViajeDetalle.IDRuta = recViajeDetalle_List("IDRuta").value
        ViajeDetalle.Indice = recViajeDetalle_List("Indice").value
        If ViajeDetalle.Load() Then
            Call ViajeDetalle.VerificarPrepagoOtrasReservas
        End If
        Set ViajeDetalle = Nothing
        
        recViajeDetalle_List.MoveNext
    Loop

    recViajeDetalle_List.Close
    Set recViajeDetalle_List = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.Viaje.VerificarPrepagosOtrasReservas", "Error al verificar los Prepagos de otras Reservas."
End Sub

Private Sub VerificarPrepagosEstasReservas()
    Dim cmdViajeDetalle_List As ADODB.command
    Dim recViajeDetalle_List As ADODB.Recordset
    Dim ViajeDetalle As ViajeDetalle
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass
    
    Set cmdViajeDetalle_List = New ADODB.command
    Set cmdViajeDetalle_List.ActiveConnection = pDatabase.Connection
    cmdViajeDetalle_List.CommandText = "usp_PersonaPrepago_GetViajeDetalleList"
    cmdViajeDetalle_List.CommandType = adCmdStoredProc
    cmdViajeDetalle_List.Parameters.Append cmdViajeDetalle_List.CreateParameter("FechaHora", adDate, adParamInput, , mFechaHora)
    cmdViajeDetalle_List.Parameters.Append cmdViajeDetalle_List.CreateParameter("IDRuta", adChar, adParamInput, 20, mIDRuta)
    Set recViajeDetalle_List = New ADODB.Recordset
    recViajeDetalle_List.Open cmdViajeDetalle_List, , adOpenForwardOnly, adLockReadOnly
    Set cmdViajeDetalle_List = Nothing
    
    Do While Not recViajeDetalle_List.EOF
        Set ViajeDetalle = New ViajeDetalle
        ViajeDetalle.FechaHora = recViajeDetalle_List("FechaHora").value
        ViajeDetalle.IDRuta = recViajeDetalle_List("IDRuta").value
        ViajeDetalle.Indice = recViajeDetalle_List("Indice").value
        If ViajeDetalle.Load() Then
            Call ViajeDetalle.VerificarPrepagoEstaReserva
        End If
        Set ViajeDetalle = Nothing
        
        recViajeDetalle_List.MoveNext
    Loop

    recViajeDetalle_List.Close
    Set recViajeDetalle_List = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.Viaje.VerificarPrepagosEstasReservas", "Error al verificar los Prepagos de estas Reservas."
End Sub
