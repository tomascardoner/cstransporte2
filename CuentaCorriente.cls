VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuentaCorriente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDMovimiento As Long
Private mIDCuentaCorrienteGrupo As Long
Private mIDCuentaCorrienteCaja As Long
Private mIDPersona As Long
Private mFechaHora As Date
Private mDescripcion As String
Private mImporte As Currency
Private mIDMedioPago As Byte
Private mCuotas As Byte
Private mOperacion As String
Private mIDPersonaOrigen As Long
Private mNotas As String
Private mSaldoAnterior As Boolean
Private mViaje_FechaHora As Date
Private mViaje_IDRuta As String
Private mViaje_Indice As Long
Private mViaje_ConductorNumero As Byte
Private mFechaHoraCreacion As Date
Private mIDUsuarioCreacion As Integer
Private mFechaHoraModificacion As Date
Private mIDUsuarioModificacion As Integer

'TRANSFER VARIABLES
Private mIDCuentaCorrienteGrupo_Destino As Long
Private mIDCuentaCorrienteCaja_Destino As Long
Private mDescripcion_Destino As String

'STATE VARIABLES
Private mIsNew As Boolean
Private mIsCopy As Boolean
Private mIsDirty As Boolean
Private mNoMatch As Boolean

'BEHAVIOR VARIABLES
Private mNoMatchRaiseError As Boolean
Private mRefreshListSkip As Boolean

'INTERNAL VARIABLES
Private mDBConnection As ADODB.Connection
Private mrecData As ADODB.Recordset

Public IsHistory As Boolean

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDMovimiento() As Long
    IDMovimiento = mIDMovimiento
End Property

Public Property Let IDMovimiento(ByVal Value As Long)
    If Value <> mIDMovimiento Then
        mIsDirty = True
    End If
    mIDMovimiento = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDCuentaCorrienteGrupo() As Long
    IDCuentaCorrienteGrupo = mIDCuentaCorrienteGrupo
End Property

Public Property Let IDCuentaCorrienteGrupo(ByVal Value As Long)
    If Value <> mIDCuentaCorrienteGrupo Then
        mIsDirty = True
    End If
    mIDCuentaCorrienteGrupo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get CuentaCorrienteGrupo() As CuentaCorrienteGrupo
    Static oCuentaCorrienteGrupo As CuentaCorrienteGrupo
    
    If mIDCuentaCorrienteGrupo = 0 Then
        Set CuentaCorrienteGrupo = Nothing
    Else
        If oCuentaCorrienteGrupo Is Nothing Then
            Set oCuentaCorrienteGrupo = New CuentaCorrienteGrupo
        End If
        If oCuentaCorrienteGrupo.IDCuentaCorrienteGrupo <> mIDCuentaCorrienteGrupo Then
            oCuentaCorrienteGrupo.IDCuentaCorrienteGrupo = mIDCuentaCorrienteGrupo
            Call oCuentaCorrienteGrupo.Load
        End If
        Set CuentaCorrienteGrupo = oCuentaCorrienteGrupo
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDCuentaCorrienteCaja() As Long
    IDCuentaCorrienteCaja = mIDCuentaCorrienteCaja
End Property

Public Property Let IDCuentaCorrienteCaja(ByVal Value As Long)
    If Value <> mIDCuentaCorrienteCaja Then
        mIsDirty = True
    End If
    mIDCuentaCorrienteCaja = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get CuentaCorrienteCaja() As CuentaCorrienteCaja
    Static oCuentaCorrienteCaja As CuentaCorrienteCaja
    
    If mIDCuentaCorrienteCaja = 0 Then
        Set CuentaCorrienteCaja = Nothing
    Else
        If oCuentaCorrienteCaja Is Nothing Then
            Set oCuentaCorrienteCaja = New CuentaCorrienteCaja
        End If
        If oCuentaCorrienteCaja.IDCuentaCorrienteCaja <> mIDCuentaCorrienteCaja Then
            oCuentaCorrienteCaja.IDCuentaCorrienteCaja = mIDCuentaCorrienteCaja
            Call oCuentaCorrienteCaja.Load
        End If
        Set CuentaCorrienteCaja = oCuentaCorrienteCaja
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDPersona() As Long
    IDPersona = mIDPersona
End Property

Public Property Let IDPersona(ByVal Value As Long)
    If Value <> mIDPersona Then
        mIsDirty = True
    End If
    mIDPersona = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Persona() As Persona
    Static oPersona As Persona
    
    If mIDPersona = 0 Then
        Set Persona = Nothing
    Else
        If oPersona Is Nothing Then
            Set oPersona = New Persona
        End If
        If oPersona.IDPersona <> mIDPersona Then
            oPersona.IDPersona = mIDPersona
            Call oPersona.Load
        End If
        Set Persona = oPersona
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHora() As Date
    FechaHora = mFechaHora
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

Public Property Let FechaHora(ByVal Value As Date)
    If Value <> mFechaHora Then
        mIsDirty = True
    End If
    mFechaHora = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property

Public Property Let Descripcion(ByVal Value As String)
    If Value <> mDescripcion Then
        mIsDirty = True
    End If
    mDescripcion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Importe() As Currency
    Importe = mImporte
End Property

Public Property Get Importe_Formatted() As String
    Importe_Formatted = Format(mImporte, "Currency")
End Property

Public Property Let Importe(ByVal Value As Currency)
    If Value <> mImporte Then
        mIsDirty = True
    End If
    mImporte = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDMedioPago() As Byte
    IDMedioPago = mIDMedioPago
End Property

Public Property Let IDMedioPago(ByVal Value As Byte)
    If Value <> mIDMedioPago Then
        mIsDirty = True
    End If
    mIDMedioPago = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get MedioPago() As MedioPago
    Static oMedioPago As MedioPago
    
    If mIDMedioPago = 0 Then
        Set MedioPago = Nothing
    Else
        If oMedioPago Is Nothing Then
            Set oMedioPago = New MedioPago
        End If
        If oMedioPago.IDMedioPago <> mIDMedioPago Then
            oMedioPago.IDMedioPago = mIDMedioPago
            Call oMedioPago.Load
        End If
        Set MedioPago = oMedioPago
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Cuotas() As Byte
    Cuotas = mCuotas
End Property

Public Property Get Cuotas_Formatted() As String
    Cuotas_Formatted = IIf(mCuotas = 0, "", mCuotas)
End Property

Public Property Let Cuotas(ByVal Value As Byte)
    If Value <> mCuotas Then
        mIsDirty = True
    End If
    mCuotas = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Operacion() As String
    Operacion = mOperacion
End Property

Public Property Let Operacion(ByVal Value As String)
    If Value <> mOperacion Then
        mIsDirty = True
    End If
    mOperacion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDPersonaOrigen() As Long
    IDPersonaOrigen = mIDPersonaOrigen
End Property

Public Property Let IDPersonaOrigen(ByVal Value As Long)
    If Value <> mIDPersonaOrigen Then
        mIsDirty = True
    End If
    mIDPersonaOrigen = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get PersonaOrigen() As Persona
    Static oPersonaOrigen As Persona
    
    If mIDPersonaOrigen = 0 Then
        Set PersonaOrigen = Nothing
    Else
        If oPersonaOrigen Is Nothing Then
            Set oPersonaOrigen = New Persona
        End If
        If oPersonaOrigen.IDPersona <> mIDPersonaOrigen Then
            oPersonaOrigen.IDPersona = mIDPersonaOrigen
            Call oPersonaOrigen.Load
        End If
        Set PersonaOrigen = oPersonaOrigen
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Notas() As String
    Notas = mNotas
End Property

Public Property Let Notas(ByVal Value As String)
    If Value <> mNotas Then
        mIsDirty = True
    End If
    mNotas = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get SaldoAnterior() As Boolean
    SaldoAnterior = mSaldoAnterior
End Property

Public Property Let SaldoAnterior(ByVal Value As Boolean)
    If Value <> mSaldoAnterior Then
        mIsDirty = True
    End If
    mSaldoAnterior = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Viaje_FechaHora() As Date
    Viaje_FechaHora = mViaje_FechaHora
End Property

Public Property Get Viaje_FechaHora_Formatted() As String
    Viaje_FechaHora_Formatted = Format(mViaje_FechaHora, "Short Date") & " " & Format(mViaje_FechaHora, "Short Time")
End Property

Public Property Get Viaje_FechaHora_FormattedAsDate() As String
    Viaje_FechaHora_FormattedAsDate = Format(mViaje_FechaHora, "Short Date")
End Property

Public Property Get Viaje_FechaHora_FormattedAsTime() As String
    Viaje_FechaHora_FormattedAsTime = Format(mViaje_FechaHora, "Short Time")
End Property

Public Property Let Viaje_FechaHora(ByVal Value As Date)
    If Value <> mViaje_FechaHora Then
        mIsDirty = True
    End If
    mViaje_FechaHora = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Viaje_IDRuta() As String
    Viaje_IDRuta = mViaje_IDRuta
End Property

Public Property Let Viaje_IDRuta(ByVal Value As String)
    If Value <> mViaje_IDRuta Then
        mIsDirty = True
    End If
    mViaje_IDRuta = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Viaje_Indice() As Long
    Viaje_Indice = mViaje_Indice
End Property

Public Property Let Viaje_Indice(ByVal Value As Long)
    If Value <> mViaje_Indice Then
        mIsDirty = True
    End If
    mViaje_Indice = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ViajeDetalle() As ViajeDetalle
    Static oViajeDetalle As ViajeDetalle
    
    If oViajeDetalle Is Nothing Then
        Set oViajeDetalle = New ViajeDetalle
    End If
    If oViajeDetalle.FechaHora <> mViaje_FechaHora Or oViajeDetalle.IDRuta <> mViaje_IDRuta Or oViajeDetalle.Indice <> mViaje_Indice Then
        If mViaje_IDRuta = "" Then
            Set oViajeDetalle = New ViajeDetalle
        Else
            oViajeDetalle.FechaHora = mViaje_FechaHora
            oViajeDetalle.IDRuta = mViaje_IDRuta
            oViajeDetalle.Indice = mViaje_Indice
            Call oViajeDetalle.Load
        End If
    End If
    Set ViajeDetalle = oViajeDetalle
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Viaje_ConductorNumero() As Long
    Viaje_ConductorNumero = mViaje_ConductorNumero
End Property

Public Property Let Viaje_ConductorNumero(ByVal Value As Long)
    If Value <> mViaje_ConductorNumero Then
        mIsDirty = True
    End If
    mViaje_ConductorNumero = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHoraCreacion() As Date
    FechaHoraCreacion = mFechaHoraCreacion
End Property

Public Property Get FechaHoraCreacion_Formatted() As String
    FechaHoraCreacion_Formatted = Format(mFechaHoraCreacion, "Short Date") & " " & Format(mFechaHoraCreacion, "Short Time")
End Property

Public Property Get FechaHoraCreacion_FormattedAsDate() As String
    FechaHoraCreacion_FormattedAsDate = Format(mFechaHoraCreacion, "Short Date")
End Property

Public Property Get FechaHoraCreacion_FormattedAsTime() As String
    FechaHoraCreacion_FormattedAsTime = Format(mFechaHoraCreacion, "Short Time")
End Property

Public Property Let FechaHoraCreacion(ByVal Value As Date)
    If Value <> mFechaHoraCreacion Then
        mIsDirty = True
    End If
    mFechaHoraCreacion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDUsuarioCreacion() As String
    IDUsuarioCreacion = mIDUsuarioCreacion
End Property

Public Property Let IDUsuarioCreacion(ByVal Value As String)
    If Value <> mIDUsuarioCreacion Then
        mIsDirty = True
    End If
    mIDUsuarioCreacion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHoraModificacion() As Date
    FechaHoraModificacion = mFechaHoraModificacion
End Property

Public Property Get FechaHoraModificacion_Formatted() As String
    FechaHoraModificacion_Formatted = Format(mFechaHoraModificacion, "Short Date") & " " & Format(mFechaHoraModificacion, "Short Time")
End Property

Public Property Get FechaHoraModificacion_FormattedAsDate() As String
    FechaHoraModificacion_FormattedAsDate = Format(mFechaHoraModificacion, "Short Date")
End Property

Public Property Get FechaHoraModificacion_FormattedAsTime() As String
    FechaHoraModificacion_FormattedAsTime = Format(mFechaHoraModificacion, "Short Time")
End Property

Public Property Let FechaHoraModificacion(ByVal Value As Date)
    If Value <> mFechaHoraModificacion Then
        mIsDirty = True
    End If
    mFechaHoraModificacion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDUsuarioModificacion() As String
    IDUsuarioModificacion = mIDUsuarioModificacion
End Property

Public Property Let IDUsuarioModificacion(ByVal Value As String)
    If Value <> mIDUsuarioModificacion Then
        mIsDirty = True
    End If
    mIDUsuarioModificacion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDCuentaCorrienteGrupo_Destino() As Long
    IDCuentaCorrienteGrupo_Destino = mIDCuentaCorrienteGrupo_Destino
End Property

Public Property Let IDCuentaCorrienteGrupo_Destino(ByVal Value As Long)
    If Value <> mIDCuentaCorrienteGrupo_Destino Then
        mIsDirty = True
    End If
    mIDCuentaCorrienteGrupo_Destino = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDCuentaCorrienteCaja_Destino() As Long
    IDCuentaCorrienteCaja_Destino = mIDCuentaCorrienteCaja_Destino
End Property

Public Property Let IDCuentaCorrienteCaja_Destino(ByVal Value As Long)
    If Value <> mIDCuentaCorrienteCaja_Destino Then
        mIsDirty = True
    End If
    mIDCuentaCorrienteCaja_Destino = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Descripcion_Destino() As String
    Descripcion_Destino = mDescripcion_Destino
End Property

Public Property Let Descripcion_Destino(ByVal Value As String)
    If Value <> mDescripcion_Destino Then
        mIsDirty = True
    End If
    mDescripcion_Destino = Value
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

Public Property Let NoMatchRaiseError(ByVal Value As Boolean)
    mNoMatchRaiseError = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RefreshListSkip() As Boolean
    RefreshListSkip = mRefreshListSkip
End Property

Public Property Let RefreshListSkip(ByVal Value As Boolean)
    mRefreshListSkip = Value
End Property

Public Sub RefreshList()
    RefreshList_RefreshCuentaCorriente IDMovimiento
End Sub

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DatabaseConnection() As ADODB.Connection
    Set DatabaseConnection = mDBConnection
End Property

Public Property Set DatabaseConnection(ByVal Value As ADODB.Connection)
    Set mDBConnection = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function Load() As Boolean
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Call InitializeValues


    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = mDBConnection
    cmdData.CommandText = IIf(IsHistory, "[" & pParametro.Database_DatabaseHistory & "]..", "") & "sp_CuentaCorriente_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDMovimiento", adInteger, adParamInput, , mIDMovimiento)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDCuentaCorrienteGrupo = mrecData("IDCuentaCorrienteGrupo").Value
        mIDCuentaCorrienteCaja = mrecData("IDCuentaCorrienteCaja").Value
        mIDPersona = Val(mrecData("IDPersona").Value & "")
        mFechaHora = mrecData("FechaHora").Value
        mDescripcion = mrecData("Descripcion").Value
        mImporte = mrecData("Importe").Value
        mIDMedioPago = Val(mrecData("IDMedioPago").Value & "")
        mCuotas = Val(mrecData("Cuotas").Value & "")
        mOperacion = mrecData("Operacion").Value & ""
        mIDPersonaOrigen = Val(mrecData("IDPersonaOrigen").Value & "")
        mNotas = mrecData("Notas").Value & ""
        mSaldoAnterior = mrecData("SaldoAnterior").Value
        mViaje_FechaHora = IIf(IsNull(mrecData("Viaje_FechaHora").Value), DATE_TIME_FIELD_NULL_VALUE, mrecData("Viaje_FechaHora").Value)
        mViaje_IDRuta = mrecData("Viaje_IDRuta").Value & ""
        mViaje_Indice = IIf(IsNull(mrecData("Viaje_Indice").Value), -1, mrecData("Viaje_Indice").Value)
        mViaje_ConductorNumero = Val(mrecData("Viaje_ConductorNumero").Value & "")
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.CuentaCorriente.Load", "Error al obtener los datos de el Movimiento de CuentaCorriente." & vbCr & vbCr & "IDMovimiento: " & IDMovimiento
End Function

Public Function LoadByViaje() As Boolean
    Dim Save_IDCuentaCorrienteGrupo As Long
    Dim Save_Viaje_FechaHora As Date
    Dim Save_Viaje_IDRuta As String
    Dim Save_Viaje_ConductorNumero As Byte
    Dim cmdData As ADODB.command
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass
    
    Save_IDCuentaCorrienteGrupo = mIDCuentaCorrienteGrupo
    Save_Viaje_FechaHora = mViaje_FechaHora
    Save_Viaje_IDRuta = mViaje_IDRuta
    Save_Viaje_ConductorNumero = mViaje_ConductorNumero
    mIDMovimiento = 0
    Call InitializeValues
    mIDCuentaCorrienteGrupo = Save_IDCuentaCorrienteGrupo
    mViaje_FechaHora = Save_Viaje_FechaHora
    mViaje_IDRuta = Save_Viaje_IDRuta
    mViaje_Indice = 0
    mViaje_ConductorNumero = Save_Viaje_ConductorNumero
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_CuentaCorriente_DataByViaje" & IIf(mViaje_ConductorNumero = 0, "", "ConductorNumero")
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDCuentaCorrienteGrupo_FILTER", adInteger, adParamInput, , mIDCuentaCorrienteGrupo)
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , mViaje_FechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, mViaje_IDRuta)
    If mViaje_ConductorNumero > 0 Then
        cmdData.Parameters.Append cmdData.CreateParameter("ConductorNumero_FILTER", adTinyInt, adParamInput, , mViaje_ConductorNumero)
    End If
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDMovimiento = mrecData("IDMovimiento").Value
        mIDCuentaCorrienteCaja = mrecData("IDCuentaCorrienteCaja").Value
        mIDPersona = Val(mrecData("IDPersona").Value & "")
        mFechaHora = mrecData("FechaHora").Value
        mDescripcion = mrecData("Descripcion").Value
        mImporte = mrecData("Importe").Value
        mIDMedioPago = Val(mrecData("IDMedioPago").Value & "")
        mCuotas = Val(mrecData("Cuotas").Value & "")
        mOperacion = mrecData("Operacion").Value & ""
        mIDPersonaOrigen = Val(mrecData("IDPersonaOrigen").Value & "")
        mNotas = mrecData("Notas").Value & ""
        mSaldoAnterior = mrecData("SaldoAnterior").Value
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    
    Screen.MousePointer = vbDefault
    LoadByViaje = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.CuentaCorriente.LoadByViaje", "Error al obtener los datos de el Movimiento de CuentaCorriente." & vbCr & vbCr & "IDCuentaCorrienteGrupo: " & IDCuentaCorrienteGrupo & vbCr & "Viaje_FechaHora: " & Viaje_FechaHora_Formatted & vbCr & "Viaje_IDRuta: " & Viaje_IDRuta
End Function

Public Function LoadByViajeDetalle() As Boolean
    Dim Save_IDCuentaCorrienteGrupo As Long
    Dim Save_Viaje_FechaHora As Date
    Dim Save_Viaje_IDRuta As String
    Dim Save_Viaje_Indice As Long
    Dim cmdData As ADODB.command
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass
    
    Save_IDCuentaCorrienteGrupo = mIDCuentaCorrienteGrupo
    Save_Viaje_FechaHora = mViaje_FechaHora
    Save_Viaje_IDRuta = mViaje_IDRuta
    Save_Viaje_Indice = mViaje_Indice
    mIDMovimiento = 0
    Call InitializeValues
    mIDCuentaCorrienteGrupo = Save_IDCuentaCorrienteGrupo
    mViaje_FechaHora = Save_Viaje_FechaHora
    mViaje_IDRuta = Save_Viaje_IDRuta
    mViaje_Indice = Save_Viaje_Indice
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_CuentaCorriente_DataByViajeDetalle"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDCuentaCorrienteGrupo_FILTER", adInteger, adParamInput, , mIDCuentaCorrienteGrupo)
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , mViaje_FechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, mViaje_IDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("Indice_FILTER", adInteger, adParamInput, , mViaje_Indice)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDMovimiento = mrecData("IDMovimiento").Value
        mIDCuentaCorrienteCaja = mrecData("IDCuentaCorrienteCaja").Value
        mIDPersona = Val(mrecData("IDPersona").Value & "")
        mFechaHora = mrecData("FechaHora").Value
        mDescripcion = mrecData("Descripcion").Value
        mImporte = mrecData("Importe").Value
        mIDMedioPago = Val(mrecData("IDMedioPago").Value & "")
        mCuotas = Val(mrecData("Cuotas").Value & "")
        mOperacion = mrecData("Operacion").Value & ""
        mIDPersonaOrigen = Val(mrecData("IDPersonaOrigen").Value & "")
        mNotas = mrecData("Notas").Value & ""
        mSaldoAnterior = mrecData("SaldoAnterior").Value
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    
    Screen.MousePointer = vbDefault
    LoadByViajeDetalle = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.CuentaCorriente.LoadByViajeDetalle", "Error al obtener los datos de el Movimiento de CuentaCorriente." & vbCr & vbCr & "IDCuentaCorrienteGrupo: " & IDCuentaCorrienteGrupo & vbCr & "Viaje_FechaHora: " & Viaje_FechaHora_Formatted & vbCr & "Viaje_IDRuta: " & Viaje_IDRuta & vbCr & "Viaje_Indice: " & Viaje_Indice
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIDMovimiento = 0

        mIsNew = True
        mIsCopy = True
        mIsDirty = True
    End If
End Function

Public Function Update(Optional PermiteMultiples As Boolean = False) As Boolean
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
    cmdData.CommandText = "sp_CuentaCorriente_Update"
    cmdData.CommandType = adCmdStoredProc
    cmdData.NamedParameters = True
    cmdData.Parameters.Append cmdData.CreateParameter("@IDMovimiento", adInteger, adParamInputOutput, , mIDMovimiento)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDCuentaCorrienteGrupo", adInteger, adParamInput, , mIDCuentaCorrienteGrupo)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDCuentaCorrienteCaja", adInteger, adParamInput, , mIDCuentaCorrienteCaja)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDPersona", adInteger, adParamInput, , IIf(mIDPersona = 0, Null, mIDPersona))
    cmdData.Parameters.Append cmdData.CreateParameter("@FechaHora", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("@Descripcion", adVarChar, adParamInput, 255, mDescripcion)
    cmdData.Parameters.Append cmdData.CreateParameter("@Importe", adCurrency, adParamInput, , mImporte)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDMedioPago", adTinyInt, adParamInput, , IIf(mIDMedioPago = 0, pParametro.MedioPago_Predeterminado_ID, mIDMedioPago))
    cmdData.Parameters.Append cmdData.CreateParameter("@Cuotas", adTinyInt, adParamInput, , IIf(mCuotas = 0, Null, mCuotas))
    cmdData.Parameters.Append cmdData.CreateParameter("@Operacion", adVarChar, adParamInput, 20, IIf(Trim(mOperacion) = "", Null, mOperacion))
    cmdData.Parameters.Append cmdData.CreateParameter("@IDPersonaOrigen", adInteger, adParamInput, , IIf(mIDPersonaOrigen = 0, Null, mIDPersonaOrigen))
    cmdData.Parameters.Append cmdData.CreateParameter("@Notas", adVarChar, adParamInput, 8000, IIf(Trim(mNotas) = "", Null, mNotas))
    cmdData.Parameters.Append cmdData.CreateParameter("@SaldoAnterior", adBoolean, adParamInput, , mSaldoAnterior)
    cmdData.Parameters.Append cmdData.CreateParameter("@Viaje_FechaHora", adDate, adParamInput, , IIf(mViaje_FechaHora = DATE_TIME_FIELD_NULL_VALUE, Null, mViaje_FechaHora))
    cmdData.Parameters.Append cmdData.CreateParameter("@Viaje_IDRuta", adChar, adParamInput, 20, IIf(mViaje_IDRuta = "", Null, mViaje_IDRuta))
    cmdData.Parameters.Append cmdData.CreateParameter("@Viaje_Indice", adInteger, adParamInput, , IIf(mViaje_Indice = -1, Null, mViaje_Indice))
    cmdData.Parameters.Append cmdData.CreateParameter("@Viaje_ConductorNumero", adTinyInt, adParamInput, , IIf(mViaje_ConductorNumero = 0, Null, mViaje_ConductorNumero))
    cmdData.Parameters.Append cmdData.CreateParameter("@IDUsuario", adChar, adParamInput, 30, pUsuario.IDUsuario)
    cmdData.Parameters.Append cmdData.CreateParameter("@PermiteMultiples", adBoolean, adParamInput, , PermiteMultiples)
    cmdData.Execute 0
    mIDMovimiento = cmdData.Parameters("@IDMovimiento").Value
    Set cmdData = Nothing
    
    mIsCopy = False
    mIsDirty = False
    
    If Not mRefreshListSkip Then
        Call RefreshList
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            ShowErrorMessage "Classes.CuentaCorriente.Update", "Error al actualizar el Movimiento de Cuenta Corriente." & vbCr & vbCr & "IDMovimiento: " & IDMovimiento
        Else
            ShowErrorMessage "Classes.CuentaCorriente.Update", "Error al actualizar el Movimiento de Cuenta Corriente." & vbCr & vbCr & "IDMovimiento: " & IDMovimiento
        End If
    Else
        ShowErrorMessage "Classes.CuentaCorriente.Update", "Error al actualizar el Movimiento de Cuenta Corriente." & vbCr & vbCr & "IDMovimiento: " & IDMovimiento
    End If
End Function

Public Function Transferir() As Boolean
    Dim cmdData As ADODB.command
    
    If Not mIsDirty Then
        Transferir = True
        Exit Function
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_CuentaCorriente_Transferencia"
    cmdData.CommandType = adCmdStoredProc
    cmdData.NamedParameters = True
    cmdData.Parameters.Append cmdData.CreateParameter("@IDCuentaCorrienteGrupo_Origen", adInteger, adParamInput, , mIDCuentaCorrienteGrupo)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDCuentaCorrienteCaja_Origen", adInteger, adParamInput, , mIDCuentaCorrienteCaja)
    cmdData.Parameters.Append cmdData.CreateParameter("@Descripcion_Origen", adVarChar, adParamInput, 255, mDescripcion)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDCuentaCorrienteGrupo_Destino", adInteger, adParamInput, , mIDCuentaCorrienteGrupo_Destino)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDCuentaCorrienteCaja_Destino", adInteger, adParamInput, , mIDCuentaCorrienteCaja_Destino)
    cmdData.Parameters.Append cmdData.CreateParameter("@Descripcion_Destino", adVarChar, adParamInput, 255, mDescripcion_Destino)
    cmdData.Parameters.Append cmdData.CreateParameter("@FechaHora", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("@Importe", adCurrency, adParamInput, , mImporte)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDMedioPago", adTinyInt, adParamInput, , IIf(mIDMedioPago = 0, Null, mIDMedioPago))
    cmdData.Parameters.Append cmdData.CreateParameter("@IDUsuario", adChar, adParamInput, 30, pUsuario.IDUsuario)
    cmdData.Execute 0
    Set cmdData = Nothing
    
    mIsCopy = False
    mIsDirty = False
    
    If Not mRefreshListSkip Then
        Call RefreshList
    End If
    
    Screen.MousePointer = vbDefault
    Transferir = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            ShowErrorMessage "Classes.CuentaCorriente.Transferir", "Error al generar la Transferencia de Cuenta Corriente." & vbCr & vbCr & "IDMovimiento: " & IDMovimiento
            If mrecData.EditMode = adEditAdd Or mrecData.EditMode = adEditInProgress Then
                mrecData.CancelUpdate
            End If
        Else
            ShowErrorMessage "Classes.CuentaCorriente.Transferir", "Error al generar la Transferencia de Cuenta Corriente." & vbCr & vbCr & "IDMovimiento: " & IDMovimiento
            If mrecData.EditMode = adEditAdd Or mrecData.EditMode = adEditInProgress Then
                mrecData.CancelUpdate
            End If
        End If
    Else
        ShowErrorMessage "Classes.CuentaCorriente.Transferir", "Error al generar la Transferencia de Cuenta Corriente." & vbCr & vbCr & "IDMovimiento: " & IDMovimiento
    End If
End Function

Public Function Delete() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass
    
    mrecData.Delete
    
    If Not mRefreshListSkip Then
        Call RefreshList
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede eliminar el Movimiento de Cuenta Corriente debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.CuentaCorriente.Delete", "Error al eliminar el Movimiento de Cuenta Corriente." & vbCr & vbCr & "IDMovimiento: " & IDMovimiento
        End If
        If mrecData.EditMode = adEditDelete Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.CuentaCorriente.Delete", "Error al eliminar el Movimiento de Cuenta Corriente." & vbCr & vbCr & "IDMovimiento: " & IDMovimiento
    End If
End Function

Private Sub InitializeValues()
    mIDCuentaCorrienteGrupo = 0
    mIDCuentaCorrienteCaja = 0
    mIDPersona = 0
    mFechaHora = DATE_TIME_FIELD_NULL_VALUE
    mDescripcion = ""
    mImporte = 0
    mIDMedioPago = 0
    mCuotas = 0
    mOperacion = ""
    mIDPersonaOrigen = 0
    mNotas = ""
    mSaldoAnterior = False
    mViaje_FechaHora = DATE_TIME_FIELD_NULL_VALUE
    mViaje_IDRuta = ""
    mViaje_Indice = -1
    mFechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioCreacion = 0
    mFechaHoraModificacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioModificacion = 0

    mIDCuentaCorrienteGrupo_Destino = 0
    mIDCuentaCorrienteCaja_Destino = 0
    mDescripcion_Destino = ""

    mIsNew = True
    mIsCopy = False
    mIsDirty = False
End Sub

Private Sub Class_Initialize()
    InitializeValues
    mNoMatchRaiseError = True
    mRefreshListSkip = False
    Set mDBConnection = pDatabase.Connection
End Sub

Private Sub Class_Terminate()
    If Not mrecData Is Nothing Then
        If mrecData.State = adStateOpen Then
            mrecData.Close
        End If
        Set mrecData = Nothing
    End If
    Set mDBConnection = Nothing
End Sub
