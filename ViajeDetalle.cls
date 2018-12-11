VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViajeDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDViajeDetalle As Long
Private mIDViaje As Long
Private mFechaHora As Date
Private mIDRuta As String
Private mIndice As Long
Private mGrupoNumero As Byte
Private mReservaCodigo As String
Private mOcupanteTipo As String
Private mEstado As String
Private mPrioridad As Long
Private mOrden As Long
Private mAsiento As Integer
Private mRealizado As Long
Private mIDPersona As Long
Private mIDPersonaMenor As Long
Private mIDListaPrecio As Long
Private mIDOrigen As Long
Private mSube As String
Private mIDDestino As Long
Private mBaja As String
Private mValorDeclarado As Currency
Private mImporteSeguro As Currency
Private mImporte As Currency
Private mImporteContado As Currency
Private mIDMedioPago As Byte
Private mCuotas As Byte
Private mOperacion As String
Private mImporteCuentaCorriente As Currency
Private mImprimirSaldo As Boolean
Private mIDCuentaCorrienteCaja As Long
Private mForzarDebito As Boolean
Private mIDPersonaCuentaCorriente As Long
Private mFacturar As Boolean
Private mFacturarNotas As String
Private mFacturaNumero As String
Private mIDPersonaRecibe As Long
Private mPagaQuienRecibe As Boolean
Private mRecibe As String
Private mDescripcion As String
Private mDomicilio As String
Private mHorario As String
Private mTelefono As String
Private mDejarTraer As String
Private mEntregada As Boolean
Private mEntregadaFechaHora As Date
Private mRetira As String
Private mReservaTipo As String
Private mCreadoEnProgreso As Boolean
Private mModificadoEnProgreso As Boolean
Private mReservadoPor As String
Private mCanceladoPor As String
Private mCanceladoFechaHora As Date
Private mNotas As String
Private mFechaHoraCreacion As Date
Private mIDUsuarioCreacion As Integer
Private mFechaHoraModificacion As Date
Private mIDUsuarioModificacion As Integer
Private mIDUsuarioCancelacion As String

'STATE VARIABLES
Private mIsNew As Boolean
Private mIsCopy As Boolean
Private mIsDirty As Boolean
Private mNoMatch As Boolean

'BEHAVIOR VARIABLES
Private mNoMatchRaiseError As Boolean
Private mRefreshListSkip As Boolean

'WORK VARIABLES
Private mCallAsientoAsignarOnAddNew As Boolean
Private mVerifyDuplicates As Boolean
Private mVerifyEstadoCondicionalOnAddNew As Boolean
Private mFechaHoraOriginal As Date
Private mIDRutaOriginal As String
Private mIndiceOriginal As Long
Private mIDOrigenOriginal As Long
Private mSubeOriginal As String
Private mIDDestinoOriginal As Long
Private mBajaOriginal As String
Private mGenerandoReservasFijas As Boolean

'INTERNAL VARIABLES
Private mrecData As ADODB.Recordset

Private mCPagos As Collection

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDViajeDetalle() As Long
    IDViajeDetalle = mIDViajeDetalle
End Property

Public Property Get IDViajeDetalle_Formatted() As String
    IDViajeDetalle_Formatted = Format(mIDViaje, "#,###")
End Property

Public Property Let IDViajeDetalle(ByVal Value As Long)
    If Value <> mIDViajeDetalle Then
        mIsDirty = True
    End If
    mIDViajeDetalle = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDViaje() As Long
    IDViaje = mIDViaje
End Property

Public Property Let IDViaje(ByVal Value As Long)
    If Value <> mIDViaje Then
        mIsDirty = True
    End If
    mIDViaje = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHora() As Date
    FechaHora = mFechaHora
End Property

Public Property Get FechaHora_Formatted() As String
    FechaHora_Formatted = Format(FechaHora, "Short Date") & " " & Format(FechaHora, "Short Time")
End Property

Public Property Get FechaHora_WeekdayName() As String
    FechaHora_WeekdayName = WeekdayName(Weekday(FechaHora))
End Property

Public Property Get FechaHora_FormattedAsDate() As String
    FechaHora_FormattedAsDate = Format(FechaHora, "Short Date")
End Property

Public Property Get FechaHora_FormattedAsTime() As String
    FechaHora_FormattedAsTime = Format(FechaHora, "Short Time")
End Property

Public Property Let FechaHora(ByVal Value As Date)
    If Value <> mFechaHora Then
        mIsDirty = True
    End If
    mFechaHora = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDRuta() As String
    IDRuta = mIDRuta
End Property

Public Property Let IDRuta(ByVal Value As String)
    If Value <> mIDRuta Then
        mIsDirty = True
    End If
    mIDRuta = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Viaje() As Viaje
    Static oViaje As Viaje
    
    If oViaje Is Nothing Then
        Set oViaje = New Viaje
    End If
    If oViaje.FechaHora <> mFechaHora Or oViaje.IDRuta <> mIDRuta Then
        If mIDRuta = "" Then
            Set oViaje = New Viaje
        Else
            oViaje.FechaHora = mFechaHora
            oViaje.IDRuta = mIDRuta
            Call oViaje.Load
        End If
    End If
    Set Viaje = oViaje
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Indice() As Long
    Indice = mIndice
End Property

Public Property Let Indice(ByVal Value As Long)
    If Value <> mIndice Then
        mIsDirty = True
    End If
    mIndice = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get GrupoNumero() As Byte
    GrupoNumero = mGrupoNumero
End Property

Public Property Let GrupoNumero(ByVal Value As Byte)
    If Value <> mGrupoNumero Then
        mIsDirty = True
    End If
    mGrupoNumero = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ReservaCodigo() As String
    ReservaCodigo = mReservaCodigo
End Property

Public Property Let ReservaCodigo(ByVal Value As String)
    If Value <> mReservaCodigo Then
        mIsDirty = True
    End If
    mReservaCodigo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get OcupanteTipo() As String
    OcupanteTipo = mOcupanteTipo
End Property

Public Property Get OcupanteTipoNombre() As String
    Select Case mOcupanteTipo
        Case OCUPANTE_TIPO_COMISION
            OcupanteTipoNombre = OCUPANTE_TIPO_COMISION_NOMBRE
        Case OCUPANTE_TIPO_PASAJERO
            OcupanteTipoNombre = OCUPANTE_TIPO_PASAJERO_NOMBRE
    End Select
End Property

Public Property Let OcupanteTipo(ByVal Value As String)
    If Value <> mOcupanteTipo Then
        mIsDirty = True
    End If
    mOcupanteTipo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Estado() As String
    Estado = mEstado
End Property

Public Property Get Estado_ToString() As String
    Select Case Estado
        Case VIAJE_DETALLE_ESTADO_CONFIRMADO
            Estado_ToString = VIAJE_DETALLE_ESTADO_CONFIRMADO_NOMBRE
        Case VIAJE_DETALLE_ESTADO_CONDICIONAL
            Estado_ToString = VIAJE_DETALLE_ESTADO_CONDICIONAL_NOMBRE
        Case VIAJE_DETALLE_ESTADO_CANCELADO
            Estado_ToString = VIAJE_DETALLE_ESTADO_CANCELADO_NOMBRE
        Case Else
            Estado_ToString = ""
    End Select
End Property

Public Property Let Estado(ByVal Value As String)
    If Value <> mEstado Then
        mIsDirty = True
    End If
    mEstado = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Prioridad() As Long
    Prioridad = mPrioridad
End Property

Public Property Let Prioridad(ByVal Value As Long)
    If Value <> mPrioridad Then
        mIsDirty = True
    End If
    mPrioridad = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Orden() As Long
    Orden = mOrden
End Property

Public Property Let Orden(ByVal Value As Long)
    If Value <> mOrden Then
        mIsDirty = True
    End If
    mOrden = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Asiento() As Integer
    Asiento = mAsiento
End Property

Public Property Let Asiento(ByVal Value As Integer)
    If Value <> mAsiento Then
        mIsDirty = True
    End If
    mAsiento = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Realizado() As Long
    Realizado = mRealizado
End Property

Public Property Let Realizado(ByVal Value As Long)
    If Value <> mRealizado Then
        mIsDirty = True
    End If
    mRealizado = Value
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

Public Property Get IDPersonaMenor() As Long
    IDPersonaMenor = mIDPersonaMenor
End Property

Public Property Let IDPersonaMenor(ByVal Value As Long)
    If Value <> mIDPersonaMenor Then
        mIsDirty = True
    End If
    mIDPersonaMenor = Value
End Property

Public Property Get PersonaMenor() As Persona
    Static oPersona As Persona
    
    If mIDPersonaMenor = 0 Then
        Set PersonaMenor = Nothing
    Else
        If oPersona Is Nothing Then
            Set oPersona = New Persona
        End If
        If oPersona.IDPersona <> mIDPersonaMenor Then
            oPersona.IDPersona = mIDPersonaMenor
            Call oPersona.Load
        End If
        Set PersonaMenor = oPersona
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDListaPrecio() As Long
    IDListaPrecio = mIDListaPrecio
End Property

Public Property Let IDListaPrecio(ByVal Value As Long)
    If Value <> mIDListaPrecio Then
        mIsDirty = True
    End If
    mIDListaPrecio = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ListaPrecio() As ListaPrecio
    Static oListaPrecio As ListaPrecio
    
    If mIDListaPrecio = 0 Then
        Set ListaPrecio = Nothing
    Else
        If oListaPrecio Is Nothing Then
            Set oListaPrecio = New ListaPrecio
        End If
        If oListaPrecio.IDListaPrecio <> mIDListaPrecio Then
            oListaPrecio.IDListaPrecio = mIDListaPrecio
            Call oListaPrecio.Load
        End If
        Set ListaPrecio = oListaPrecio
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDOrigen() As Long
    IDOrigen = mIDOrigen
End Property

Public Property Let IDOrigen(ByVal Value As Long)
    If Value <> mIDOrigen Then
        mIsDirty = True
    End If
    mIDOrigen = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Sube() As String
    Sube = mSube
End Property

Public Property Let Sube(ByVal Value As String)
    If Value <> mSube Then
        mIsDirty = True
    End If
    mSube = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDDestino() As Long
    IDDestino = mIDDestino
End Property

Public Property Let IDDestino(ByVal Value As Long)
    If Value <> mIDDestino Then
        mIsDirty = True
    End If
    mIDDestino = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Baja() As String
    Baja = mBaja
End Property

Public Property Let Baja(ByVal Value As String)
    If Value <> mBaja Then
        mIsDirty = True
    End If
    mBaja = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ValorDeclarado() As Currency
    ValorDeclarado = mValorDeclarado
End Property

Public Property Get ValorDeclarado_Formatted() As String
    ValorDeclarado_Formatted = Format(ValorDeclarado, "Currency")
End Property

Public Property Let ValorDeclarado(ByVal Value As Currency)
    If Value <> mValorDeclarado Then
        mIsDirty = True
    End If
    mValorDeclarado = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ImporteSeguro() As Currency
    ImporteSeguro = mImporteSeguro
End Property

Public Property Get ImporteSeguro_Formatted() As String
    ImporteSeguro_Formatted = Format(ImporteSeguro, "Currency")
End Property

Public Property Let ImporteSeguro(ByVal Value As Currency)
    If Value <> mImporteSeguro Then
        mIsDirty = True
    End If
    mImporteSeguro = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Importe() As Currency
    Importe = mImporte
End Property

Public Property Get Importe_Formatted() As String
    Importe_Formatted = Format(Importe, "Currency")
End Property

Public Property Let Importe(ByVal Value As Currency)
    If Value <> mImporte Then
        mIsDirty = True
    End If
    mImporte = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ImporteContado() As Currency
    ImporteContado = mImporteContado
End Property

Public Property Get ImporteContado_Formatted() As String
    ImporteContado_Formatted = Format(ImporteContado, "Currency")
End Property

Public Property Let ImporteContado(ByVal Value As Currency)
    If Value <> mImporteContado Then
        mIsDirty = True
    End If
    mImporteContado = Value
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

Public Property Get ImporteCuentaCorriente() As Currency
    ImporteCuentaCorriente = mImporteCuentaCorriente
End Property

Public Property Get ImporteCuentaCorriente_Formatted() As String
    ImporteCuentaCorriente_Formatted = Format(ImporteCuentaCorriente, "Currency")
End Property

Public Property Let ImporteCuentaCorriente(ByVal Value As Currency)
    If Value <> mImporteCuentaCorriente Then
        mIsDirty = True
    End If
    mImporteCuentaCorriente = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ImprimirSaldo() As Boolean
    ImprimirSaldo = mImprimirSaldo
End Property

Public Property Let ImprimirSaldo(ByVal Value As Boolean)
    If Value <> mImprimirSaldo Then
        mIsDirty = True
    End If
    mImprimirSaldo = Value
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

Public Property Get ForzarDebito() As Boolean
    ForzarDebito = mForzarDebito
End Property

Public Property Let ForzarDebito(ByVal Value As Boolean)
    If Value <> mForzarDebito Then
        mIsDirty = True
    End If
    mForzarDebito = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDPersonaCuentaCorriente() As Long
    IDPersonaCuentaCorriente = mIDPersonaCuentaCorriente
End Property

Public Property Let IDPersonaCuentaCorriente(ByVal Value As Long)
    If Value <> mIDPersonaCuentaCorriente Then
        mIsDirty = True
    End If
    mIDPersonaCuentaCorriente = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Facturar() As Boolean
    Facturar = mFacturar
End Property

Public Property Let Facturar(ByVal Value As Boolean)
    If Value <> mFacturar Then
        mIsDirty = True
    End If
    mFacturar = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FacturarNotas() As String
    FacturarNotas = mFacturarNotas
End Property

Public Property Let FacturarNotas(ByVal Value As String)
    If Value <> mFacturarNotas Then
        mIsDirty = True
    End If
    mFacturarNotas = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FacturaNumero() As String
    FacturaNumero = mFacturaNumero
End Property

Public Property Let FacturaNumero(ByVal Value As String)
    If Value <> mFacturaNumero Then
        mIsDirty = True
    End If
    mFacturaNumero = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDPersonaRecibe() As Long
    IDPersonaRecibe = mIDPersonaRecibe
End Property

Public Property Let IDPersonaRecibe(ByVal Value As Long)
    If Value <> mIDPersonaRecibe Then
        mIsDirty = True
    End If
    mIDPersonaRecibe = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get PagaQuienRecibe() As Boolean
    PagaQuienRecibe = mPagaQuienRecibe
End Property

Public Property Let PagaQuienRecibe(ByVal Value As Boolean)
    If Value <> mPagaQuienRecibe Then
        mIsDirty = True
    End If
    mPagaQuienRecibe = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Recibe() As String
    Recibe = mRecibe
End Property

Public Property Let Recibe(ByVal Value As String)
    If Value <> mRecibe Then
        mIsDirty = True
    End If
    mRecibe = Value
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

Public Property Get Telefono() As String
    Telefono = mTelefono
End Property

Public Property Let Telefono(ByVal Value As String)
    If Value <> mTelefono Then
        mIsDirty = True
    End If
    mTelefono = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Domicilio() As String
    Domicilio = mDomicilio
End Property

Public Property Let Domicilio(ByVal Value As String)
    If Value <> mDomicilio Then
        mIsDirty = True
    End If
    mDomicilio = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Horario() As String
    Horario = mHorario
End Property

Public Property Let Horario(ByVal Value As String)
    If Value <> mHorario Then
        mIsDirty = True
    End If
    mHorario = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DejarTraer() As String
    DejarTraer = mDejarTraer
End Property

Public Property Let DejarTraer(ByVal Value As String)
    If Value <> mDejarTraer Then
        mIsDirty = True
    End If
    mDejarTraer = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Entregada() As Boolean
    Entregada = mEntregada
End Property

Public Property Let Entregada(ByVal Value As Boolean)
    If Value <> mEntregada Then
        mIsDirty = True
    End If
    mEntregada = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get EntregadaFechaHora() As Date
    EntregadaFechaHora = mEntregadaFechaHora
End Property

Public Property Let EntregadaFechaHora(ByVal Value As Date)
    If Value <> mEntregadaFechaHora Then
        mIsDirty = True
    End If
    mEntregadaFechaHora = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Retira() As String
    Retira = mRetira
End Property

Public Property Let Retira(ByVal Value As String)
    If Value <> mRetira Then
        mIsDirty = True
    End If
    mRetira = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ViajeDetalle_Comision() As ViajeDetalle_Comision
    Static oViajeDetalle_Comision As ViajeDetalle_Comision
    
    If mIndice = 0 Then
        Set ViajeDetalle_Comision = Nothing
    Else
        If oViajeDetalle_Comision Is Nothing Then
            Set oViajeDetalle_Comision = New ViajeDetalle_Comision
        End If
        If oViajeDetalle_Comision.FechaHora <> mFechaHora Or oViajeDetalle_Comision.IDRuta <> mIDRuta Or oViajeDetalle_Comision.Indice <> mIndice Then
            oViajeDetalle_Comision.FechaHora = mFechaHora
            oViajeDetalle_Comision.IDRuta = mIDRuta
            oViajeDetalle_Comision.Indice = mIndice
            Call oViajeDetalle_Comision.Load
        End If
        Set ViajeDetalle_Comision = oViajeDetalle_Comision
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ReservaTipo() As String
    ReservaTipo = mReservaTipo
End Property

Public Property Get ReservaTipo_ToString() As String
    Select Case ReservaTipo
        Case VIAJE_DETALLE_RESERVA_TIPO_STANDARD
            ReservaTipo_ToString = VIAJE_DETALLE_RESERVA_TIPO_STANDARD_NOMBRE
        Case VIAJE_DETALLE_RESERVA_TIPO_FIJA
            ReservaTipo_ToString = VIAJE_DETALLE_RESERVA_TIPO_FIJA_NOMBRE
        Case VIAJE_DETALLE_RESERVA_TIPO_INTERNET
            ReservaTipo_ToString = VIAJE_DETALLE_RESERVA_TIPO_INTERNET_NOMBRE
        Case Else
            ReservaTipo_ToString = ""
    End Select
End Property

Public Property Let ReservaTipo(ByVal Value As String)
    If Value <> mReservaTipo Then
        mIsDirty = True
    End If
    mReservaTipo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get CreadoEnProgreso() As Boolean
    CreadoEnProgreso = mCreadoEnProgreso
End Property

Public Property Let CreadoEnProgreso(ByVal Value As Boolean)
    If Value <> mCreadoEnProgreso Then
        mIsDirty = True
    End If
    mCreadoEnProgreso = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ModificadoEnProgreso() As Boolean
    ModificadoEnProgreso = mModificadoEnProgreso
End Property

Public Property Let ModificadoEnProgreso(ByVal Value As Boolean)
    If Value <> mModificadoEnProgreso Then
        mIsDirty = True
    End If
    mModificadoEnProgreso = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ReservadoPor() As String
    ReservadoPor = mReservadoPor
End Property

Public Property Let ReservadoPor(ByVal Value As String)
    If Value <> mReservadoPor Then
        mIsDirty = True
    End If
    mReservadoPor = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get CanceladoPor() As String
    CanceladoPor = mCanceladoPor
End Property

Public Property Let CanceladoPor(ByVal Value As String)
    If Value <> mCanceladoPor Then
        mIsDirty = True
    End If
    mCanceladoPor = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get CanceladoFechaHora() As Date
    CanceladoFechaHora = mCanceladoFechaHora
End Property

Public Property Get CanceladoFechaHora_Formatted() As String
    CanceladoFechaHora_Formatted = Format(mCanceladoFechaHora, "Short Date") & " " & Format(mCanceladoFechaHora, "Short Time")
End Property

Public Property Get CanceladoFechaHora_FormattedAsDate() As String
    CanceladoFechaHora_FormattedAsDate = Format(mCanceladoFechaHora, "Short Date")
End Property

Public Property Get CanceladoFechaHora_FormattedAsTime() As String
    CanceladoFechaHora_FormattedAsTime = Format(mCanceladoFechaHora, "Short Time")
End Property

Public Property Let CanceladoFechaHora(ByVal Value As Date)
    If Value <> mCanceladoFechaHora Then
        mIsDirty = True
    End If
    mCanceladoFechaHora = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Pagos() As Collection
    Set Pagos = mCPagos
End Property

Public Sub Pago_Add(ByRef Pago As CuentaCorriente)
    mCPagos.Add Pago
End Sub

Public Sub Pago_Delete(ByRef Pago As CuentaCorriente)
    If Pago.IDMovimiento > 0 Then
        mCPagos.Remove KEY_STRINGER & Pago.IDMovimiento
    End If
    Set Pago = Nothing
End Sub

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

Public Property Get IDUsuarioCancelacion() As String
    IDUsuarioCancelacion = mIDUsuarioCancelacion
End Property

Public Property Let IDUsuarioCancelacion(ByVal Value As String)
    If Value <> mIDUsuarioCancelacion Then
        mIsDirty = True
    End If
    mIDUsuarioCancelacion = Value
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

Public Sub MakeDirty()
    mIsDirty = True
End Sub

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
    RefreshList_Module.RefreshList_RefreshViajeDetalle FechaHora, IDRuta, Indice
End Sub

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get CallAsientoAsignarOnAddNew() As Boolean
    CallAsientoAsignarOnAddNew = mCallAsientoAsignarOnAddNew
End Property

Public Property Let CallAsientoAsignarOnAddNew(ByVal Value As Boolean)
    mCallAsientoAsignarOnAddNew = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get VerifyDuplicates() As Boolean
    VerifyDuplicates = mVerifyDuplicates
End Property

Public Property Let VerifyDuplicates(ByVal Value As Boolean)
    mVerifyDuplicates = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get VerifyEstadoCondicionalOnAddNew() As Boolean
    VerifyEstadoCondicionalOnAddNew = mVerifyEstadoCondicionalOnAddNew
End Property

Public Property Let VerifyEstadoCondicionalOnAddNew(ByVal Value As Boolean)
    mVerifyEstadoCondicionalOnAddNew = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get GenerandoReservasFijas() As Boolean
    GenerandoReservasFijas = mGenerandoReservasFijas
End Property

Public Property Let GenerandoReservasFijas(ByVal Value As Boolean)
    mGenerandoReservasFijas = Value
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
    cmdData.CommandText = "sp_ViajeDetalle_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, mIDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("Indice_FILTER", adInteger, adParamInput, , mIndice)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , pParametro.Database_CursorType, adLockOptimistic
    If pParametro.Database_CursorLocation = adUseClient Then
        mrecData.Properties("Update Criteria").Value = adCriteriaKey
    End If
    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.ViajeDetalle.OpenRecordset", "Error al abrir la tabla de Detalles de Viajes."
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
    
        mIDViajeDetalle = mrecData("IDViajeDetalle").Value
        mIDViaje = mrecData("IDViaje").Value
        mGrupoNumero = Val(mrecData("GrupoNumero").Value & "")
        mReservaCodigo = mrecData("ReservaCodigo").Value & ""
        mOcupanteTipo = mrecData("OcupanteTipo").Value
        mEstado = mrecData("Estado").Value & ""
        mPrioridad = Val(mrecData("Prioridad").Value & "")
        mOrden = Val(mrecData("Orden").Value & "")
        mAsiento = IIf(IsNull(mrecData("Asiento").Value), -1, mrecData("Asiento").Value)
        mRealizado = IIf(IsNull(mrecData("Realizado").Value), 0, IIf(mrecData("Realizado").Value, 1, 2))
        mIDPersona = mrecData("IDPersona").Value
        mIDPersonaMenor = Val(mrecData("IDPersonaMenor").Value & "")
        mIDListaPrecio = Val(mrecData("IDListaPrecio").Value & "")
        mIDOrigen = mrecData("IDOrigen").Value
        mSube = mrecData("Sube").Value & ""
        mIDDestino = mrecData("IDDestino").Value
        mBaja = mrecData("Baja").Value & ""
        mValorDeclarado = IIf(IsNull(mrecData("ValorDeclarado").Value), 0, mrecData("ValorDeclarado").Value)
        mImporteSeguro = IIf(IsNull(mrecData("ImporteSeguro").Value), 0, mrecData("ImporteSeguro").Value)
        mImporte = mrecData("Importe").Value
        mImporteContado = mrecData("ImporteContado").Value
        mIDMedioPago = Val(mrecData("IDMedioPago").Value & "")
        mCuotas = Val(mrecData("Cuotas").Value & "")
        mOperacion = mrecData("Operacion").Value & ""
        mImporteCuentaCorriente = mrecData("ImporteCuentaCorriente").Value
        mImprimirSaldo = mrecData("ImprimirSaldo").Value
        mIDCuentaCorrienteCaja = Val(mrecData("IDCuentaCorrienteCaja").Value & "")
        mForzarDebito = mrecData("ForzarDebito").Value
        mIDPersonaCuentaCorriente = Val(mrecData("IDPersonaCuentaCorriente").Value & "")
        mFacturar = mrecData("Facturar").Value
        mFacturarNotas = mrecData("FacturarNotas").Value & ""
        mFacturaNumero = mrecData("FacturaNumero").Value & ""
        mIDPersonaRecibe = Val(mrecData("IDPersonaRecibe").Value & "")
        mPagaQuienRecibe = mrecData("PagaQuienRecibe").Value
        mRecibe = mrecData("Recibe").Value & ""
        mDescripcion = mrecData("Descripcion").Value & ""
        mDomicilio = mrecData("Domicilio").Value & ""
        mHorario = mrecData("Horario").Value & ""
        mTelefono = mrecData("Telefono").Value & ""
        mDejarTraer = mrecData("DejarTraer").Value & ""
        mEntregada = mrecData("Entregada").Value
        mEntregadaFechaHora = IIf(IsNull(mrecData("EntregadaFechaHora").Value), DATE_TIME_FIELD_NULL_VALUE, mrecData("EntregadaFechaHora").Value)
        mRetira = mrecData("Retira").Value & ""
        mReservaTipo = mrecData("ReservaTipo").Value
        mCreadoEnProgreso = mrecData("CreadoEnProgreso").Value
        mModificadoEnProgreso = mrecData("ModificadoEnProgreso").Value
        mReservadoPor = mrecData("ReservadoPor").Value & ""
        mCanceladoPor = mrecData("CanceladoPor").Value & ""
        mCanceladoFechaHora = IIf(IsNull(mrecData("CanceladoFechaHora").Value), DATE_TIME_FIELD_NULL_VALUE, mrecData("CanceladoFechaHora").Value)
        mNotas = mrecData("Notas").Value & ""
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
        mIDUsuarioCancelacion = mrecData("IDUsuarioCancelacion").Value & ""
    End If
    mFechaHoraOriginal = FechaHora
    mIDRutaOriginal = IDRuta
    mIndiceOriginal = Indice
    mIDOrigenOriginal = IDOrigen
    mSubeOriginal = Sube
    mIDDestinoOriginal = IDDestino
    mBajaOriginal = Baja
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.ViajeDetalle.Load", "Error al obtener los datos del Detalle del Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta & vbCr & "Indice: " & Indice
End Function

Public Function LoadByIDViajeDetalle() As Boolean
    Dim cmdData As ADODB.command
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ViajeDetalle_LoadByIDViajeDetalle"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDViajeDetalle", adInteger, adParamInput, , mIDViajeDetalle)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , pParametro.Database_CursorType, adLockOptimistic
    If pParametro.Database_CursorLocation = adUseClient Then
        mrecData.Properties("Update Criteria").Value = adCriteriaKey
    End If
    Set cmdData = Nothing
    
    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDViajeDetalle = mrecData("IDViajeDetalle").Value
        mIDViaje = mrecData("IDViaje").Value
        mFechaHora = mrecData("FechaHora").Value
        mIDRuta = mrecData("IDRuta").Value
        mIndice = mrecData("Indice").Value
        mGrupoNumero = Val(mrecData("GrupoNumero").Value & "")
        mReservaCodigo = mrecData("ReservaCodigo").Value & ""
        mOcupanteTipo = mrecData("OcupanteTipo").Value
        mEstado = mrecData("Estado").Value & ""
        mPrioridad = Val(mrecData("Prioridad").Value & "")
        mOrden = Val(mrecData("Orden").Value & "")
        mAsiento = IIf(IsNull(mrecData("Asiento").Value), -1, mrecData("Asiento").Value)
        mRealizado = IIf(IsNull(mrecData("Realizado").Value), 0, IIf(mrecData("Realizado").Value, 1, 2))
        mIDPersona = mrecData("IDPersona").Value
        mIDPersonaMenor = Val(mrecData("IDPersonaMenor").Value & "")
        mIDListaPrecio = Val(mrecData("IDListaPrecio").Value & "")
        mIDOrigen = mrecData("IDOrigen").Value
        mSube = mrecData("Sube").Value & ""
        mIDDestino = mrecData("IDDestino").Value
        mBaja = mrecData("Baja").Value & ""
        mValorDeclarado = IIf(IsNull(mrecData("ValorDeclarado").Value), 0, mrecData("ValorDeclarado").Value)
        mImporteSeguro = IIf(IsNull(mrecData("ImporteSeguro").Value), 0, mrecData("ImporteSeguro").Value)
        mImporte = mrecData("Importe").Value
        mImporteContado = mrecData("ImporteContado").Value
        mIDMedioPago = Val(mrecData("IDMedioPago").Value & "")
        mCuotas = Val(mrecData("Cuotas").Value & "")
        mOperacion = mrecData("Operacion").Value & ""
        mImporteCuentaCorriente = mrecData("ImporteCuentaCorriente").Value
        mImprimirSaldo = mrecData("ImprimirSaldo").Value
        mIDCuentaCorrienteCaja = Val(mrecData("IDCuentaCorrienteCaja").Value & "")
        mForzarDebito = mrecData("ForzarDebito").Value
        mIDPersonaCuentaCorriente = Val(mrecData("IDPersonaCuentaCorriente").Value & "")
        mFacturar = mrecData("Facturar").Value
        mFacturarNotas = mrecData("FacturarNotas").Value & ""
        mFacturaNumero = mrecData("FacturaNumero").Value & ""
        mIDPersonaRecibe = Val(mrecData("IDPersonaRecibe").Value & "")
        mPagaQuienRecibe = mrecData("PagaQuienRecibe").Value
        mRecibe = mrecData("Recibe").Value & ""
        mDescripcion = mrecData("Descripcion").Value & ""
        mDomicilio = mrecData("Domicilio").Value & ""
        mHorario = mrecData("Horario").Value & ""
        mTelefono = mrecData("Telefono").Value & ""
        mDejarTraer = mrecData("DejarTraer").Value & ""
        mEntregada = mrecData("Entregada").Value
        mEntregadaFechaHora = IIf(IsNull(mrecData("EntregadaFechaHora").Value), DATE_TIME_FIELD_NULL_VALUE, mrecData("EntregadaFechaHora").Value)
        mRetira = mrecData("Retira").Value & ""
        mReservaTipo = mrecData("ReservaTipo").Value
        mCreadoEnProgreso = mrecData("CreadoEnProgreso").Value
        mModificadoEnProgreso = mrecData("ModificadoEnProgreso").Value
        mReservadoPor = mrecData("ReservadoPor").Value & ""
        mCanceladoPor = mrecData("CanceladoPor").Value & ""
        mCanceladoFechaHora = IIf(IsNull(mrecData("CanceladoFechaHora").Value), DATE_TIME_FIELD_NULL_VALUE, mrecData("CanceladoFechaHora").Value)
        mNotas = mrecData("Notas").Value & ""
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
        mIDUsuarioCancelacion = mrecData("IDUsuarioCancelacion").Value & ""
    End If
    mFechaHoraOriginal = FechaHora
    mIDRutaOriginal = IDRuta
    mIndiceOriginal = Indice
    mIDOrigenOriginal = IDOrigen
    mSubeOriginal = Sube
    mIDDestinoOriginal = IDDestino
    mBajaOriginal = Baja
    
    Screen.MousePointer = vbDefault
    LoadByIDViajeDetalle = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.ViajeDetalle.LoadByIDViajeDetalle", "Error al obtener los datos del Detalle del Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta & vbCr & "OcupanteTipo: " & OcupanteTipo & vbCr & "IDPersona: " & mIDPersona & vbCr & "ReservaTipo: " & mReservaTipo
End Function

Public Function LoadPagos() As Boolean
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    Dim Pago As CuentaCorriente
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Set mCPagos = New Collection
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandType = adCmdStoredProc
    cmdData.CommandText = "sp_ViajeDetalle_Pagos"
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta", adChar, adParamInput, 20, mIDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("Indice", adInteger, adParamInput, , mIndice)
    cmdData.Parameters.Append cmdData.CreateParameter("IDCuentaCorrienteGrupo_ViajeCredito", adInteger, adParamInput, , pParametro.CuentaCorrienteGrupo_ID_ViajeCredito)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    
    Do While Not recData.EOF
        Set Pago = New CuentaCorriente
        Pago.IDMovimiento = recData("IDMovimiento").Value
        If Pago.Load() Then
            mCPagos.Add Pago, KEY_STRINGER & Pago.IDMovimiento
        End If
        
        recData.MoveNext
    Loop
    
    If recData.State = adStateOpen Then
        recData.Close
    End If
    Set recData = Nothing
    
    Screen.MousePointer = vbDefault
    LoadPagos = True
    Exit Function
    
ErrorHandler:
    CSM_Error.ShowErrorMessage "Classes.ViajeDetalle.LoadPagos", "Error al cargar los pagos del Detalle del Viaje."
End Function

Public Function LoadByReservaTipo() As Boolean
    Dim cmdData As ADODB.command
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ViajeDetalle_ReservaTipoData"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, mIDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("OcupanteTipo_FILTER", adChar, adParamInput, 2, mOcupanteTipo)
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , mIDPersona)
    cmdData.Parameters.Append cmdData.CreateParameter("ReservaTipo_FILTER", adChar, adParamInput, 2, mReservaTipo)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , pParametro.Database_CursorType, adLockOptimistic
    If pParametro.Database_CursorLocation = adUseClient Then
        mrecData.Properties("Update Criteria").Value = adCriteriaKey
    End If
    Set cmdData = Nothing
    
    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
    End If
    
    Screen.MousePointer = vbDefault
    LoadByReservaTipo = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.ViajeDetalle.LoadByReservaTipo", "Error al obtener los datos del Detalle del Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta & vbCr & "OcupanteTipo: " & OcupanteTipo & vbCr & "IDPersona: " & mIDPersona & vbCr & "ReservaTipo: " & mReservaTipo
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        'mFechaHora =
        'mIDRuta =
        'mIndice =

        mIsNew = True
        mIsCopy = True
        mIsDirty = True
    End If
End Function

Public Function Update() As Boolean
    Dim Viaje As Viaje
    Dim ViajeOriginal As Viaje
    Dim Message As String
    Dim Persona As Persona
    Dim CambioDeViaje As Boolean
    
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset

    If Not mIsDirty Then
        Update = True
        Exit Function
    End If

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If mImporteContado > 0 And IDCuentaCorrienteCaja = 0 Then
        MsgBox "Debe especificar la Caja.", vbInformation, App.Title
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    CambioDeViaje = ((Not mIsNew) And (mFechaHora <> mFechaHoraOriginal Or mIDRuta <> mIDRutaOriginal))
    
    'CARGO EL VIAJE PARA TODO LO QUE SE VA A USAR
    '02/11/2006 - ANTES SE CARGABA 4 VECES EN LA MISMA FUNCION
    Set Viaje = New Viaje
    Viaje.FechaHora = mFechaHora
    Viaje.IDRuta = mIDRuta
    Call Viaje.Load
        
    'CREADO EN PROGRESO O CAMBIADO DE VIAJE EN PROGRESO
    If (Not mGenerandoReservasFijas) And (Not mCreadoEnProgreso) And (mIsNew Or CambioDeViaje) Then
        mCreadoEnProgreso = (Viaje.Estado = VIAJE_ESTADO_EN_PROGRESO)
    End If
    'MODIFICADO EN PROGRESO
    If (Not mIsNew) And (Not mCreadoEnProgreso) Then
        If mIDOrigenOriginal <> mIDOrigen Or mSubeOriginal <> mSube Or mIDDestinoOriginal <> mIDDestino Or mBajaOriginal <> mBaja Then
            mModificadoEnProgreso = (Viaje.Estado = VIAJE_ESTADO_EN_PROGRESO)
        End If
    End If

    'VERIFICO DUPLICADOS EN EL MISMO DIA Y RUTA
    If (mIsNew Or CambioDeViaje) And mVerifyDuplicates And mOcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
        Set cmdData = New ADODB.command
        Set cmdData.ActiveConnection = pDatabase.Connection
        cmdData.CommandType = adCmdStoredProc
        cmdData.CommandText = "sp_ViajeDetalle_FindDuplicate"
        cmdData.NamedParameters = True
        cmdData.Parameters.Append cmdData.CreateParameter("@FechaHora_FILTER", adDate, adParamInput, , mFechaHora)
        cmdData.Parameters.Append cmdData.CreateParameter("@IDRuta_FILTER", adChar, adParamInput, 20, mIDRuta)
        cmdData.Parameters.Append cmdData.CreateParameter("@IDPersona_FILTER", adInteger, adParamInput, , mIDPersona)
        cmdData.Parameters.Append cmdData.CreateParameter("@FechaHoraOriginal_FILTER", adDate, adParamInput, , IIf(mIsNew, mFechaHora, mFechaHoraOriginal))
        cmdData.Parameters.Append cmdData.CreateParameter("@IDRutaOriginal_FILTER", adChar, adParamInput, 20, IIf(mIsNew, mIDRuta, mIDRutaOriginal))
        cmdData.Parameters.Append cmdData.CreateParameter("@IndiceOriginal_FILTER", adInteger, adParamInput, , IIf(mIsNew, mIndice, mIndiceOriginal))
        Set recData = New ADODB.Recordset
        recData.Open cmdData, , adOpenStatic, adLockReadOnly
        If Not recData.EOF Then
            recData.MoveLast
            recData.MoveFirst
            If recData.RecordCount = 1 Then
                Message = "Ya existe 1 Reserva para este Pasajero en el mismo Día y Ruta." & vbCr
            Else
                Message = "Ya existen " & recData.RecordCount & "  Reservas para este Pasajero en el mismo Día y Ruta." & vbCr
            End If
            Set Persona = New Persona
            Persona.IDPersona = mIDPersona
            If Persona.Load() Then
                Message = Message & vbCr & "Pasajero: " & Persona.ApellidoNombre & vbCr
            End If
            Set Persona = Nothing
            Do While Not recData.EOF
                Message = Message & vbCr & "Hora: " & Format(recData("FechaHora").Value, "Short Time")
                recData.MoveNext
            Loop
            Message = Message & vbCr & vbCr & "¿Desea Generar esta Reserva de todos modos?"
            Screen.MousePointer = vbDefault
            If MsgBox(Message, vbExclamation + vbYesNo, App.Title) = vbNo Then
                recData.Close
                Set recData = Nothing
                Set cmdData = Nothing
                Screen.MousePointer = vbDefault
                Update = False
                Exit Function
            End If
            Screen.MousePointer = vbHourglass
        End If
        recData.Close
        Set recData = Nothing
        Set cmdData = Nothing
    End If
    
'    If mImporte = 0 Or Viaje.Estado = VIAJE_ESTADO_ACTIVO Or Viaje.Estado = VIAJE_ESTADO_CANCELADO Or (mEstado = VIAJE_DETALLE_ESTADO_CONFIRMADO And mRealizado = VIAJE_DETALLE_REALIZADO_NO And Not mForzarDebito) Or mEstado = VIAJE_DETALLE_ESTADO_CONDICIONAL Or (mEstado = VIAJE_DETALLE_ESTADO_CANCELADO And Not mForzarDebito) Then
'        mImporteCuentaCorriente = 0
'    Else
'        mImporteCuentaCorriente = CalcularImporteCuentaCorrienteExceptCurrent()
'    End If
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ViajeDetalle_" & IIf(mIsNew, "Insert", "Update")
    cmdData.CommandType = adCmdStoredProc
    cmdData.NamedParameters = True
    If mIsNew Then
        cmdData.Parameters.Append cmdData.CreateParameter("@IDViajeDetalle", adInteger, adParamOutput)
    Else
        cmdData.Parameters.Append cmdData.CreateParameter("@CambioDeViaje", adBoolean, adParamInput, , CambioDeViaje)
        cmdData.Parameters.Append cmdData.CreateParameter("@FechaHora_Original", adDate, adParamInput, , mFechaHoraOriginal)
        cmdData.Parameters.Append cmdData.CreateParameter("@IDRuta_Original", adChar, adParamInput, 20, mIDRutaOriginal)
        cmdData.Parameters.Append cmdData.CreateParameter("@Indice_Original", adInteger, adParamInput, , mIndiceOriginal)
    End If
    cmdData.Parameters.Append cmdData.CreateParameter("@IDViaje", adInteger, adParamInput, , mIDViaje)
    cmdData.Parameters.Append cmdData.CreateParameter("@FechaHora", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDRuta", adChar, adParamInput, 20, mIDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("@Indice", adInteger, adParamInputOutput, , mIndice)
    cmdData.Parameters.Append cmdData.CreateParameter("@GrupoNumero", adTinyInt, adParamInput, , mGrupoNumero)
    cmdData.Parameters.Append cmdData.CreateParameter("@ReservaCodigo", adChar, adParamInput, 8, mReservaCodigo)
    cmdData.Parameters.Append cmdData.CreateParameter("@OcupanteTipo", adChar, adParamInput, 2, mOcupanteTipo)
    If mIsNew Then
        If mOcupanteTipo = OCUPANTE_TIPO_COMISION Then
            mEstado = VIAJE_DETALLE_ESTADO_CONFIRMADO
        Else
            mEstado = IIf(mAsiento = -1, VIAJE_DETALLE_ESTADO_CONDICIONAL, VIAJE_DETALLE_ESTADO_CONFIRMADO)
        End If
    End If
    cmdData.Parameters.Append cmdData.CreateParameter("@Estado", adChar, adParamInput, 3, IIf(mEstado = "", Null, mEstado))
    If Not mIsNew Then
        cmdData.Parameters.Append cmdData.CreateParameter("@Prioridad", adInteger, adParamInput, , IIf(mPrioridad = 0, Null, mPrioridad))
        cmdData.Parameters.Append cmdData.CreateParameter("@Orden", adInteger, adParamInput, , IIf(mOrden = 0, Null, mOrden))
    End If
    cmdData.Parameters.Append cmdData.CreateParameter("@Asiento", adTinyInt, adParamInput, , IIf(mAsiento < 0, Null, mAsiento))
    cmdData.Parameters.Append cmdData.CreateParameter("@Realizado", adBoolean, adParamInput, , IIf(mRealizado = 0, Null, IIf(mRealizado = 1, True, False)))
    cmdData.Parameters.Append cmdData.CreateParameter("@IDPersona", adInteger, adParamInput, , mIDPersona)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDPersonaMenor", adInteger, adParamInput, , IIf(mIDPersonaMenor = 0, Null, mIDPersonaMenor))
    cmdData.Parameters.Append cmdData.CreateParameter("@IDListaPrecio", adInteger, adParamInput, , mIDListaPrecio)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDOrigen", adInteger, adParamInput, , mIDOrigen)
    cmdData.Parameters.Append cmdData.CreateParameter("@Sube", adVarChar, adParamInput, 50, IIf(Trim(mSube) = "", Null, mSube))
    cmdData.Parameters.Append cmdData.CreateParameter("@IDDestino", adInteger, adParamInput, , mIDDestino)
    cmdData.Parameters.Append cmdData.CreateParameter("@Baja", adVarChar, adParamInput, 50, IIf(Trim(mBaja) = "", Null, mBaja))
    cmdData.Parameters.Append cmdData.CreateParameter("@ValorDeclarado", adCurrency, adParamInput, , IIf(mValorDeclarado = 0, Null, mValorDeclarado))
    cmdData.Parameters.Append cmdData.CreateParameter("@ImporteSeguro", adCurrency, adParamInput, , IIf(mImporteSeguro = 0, Null, mImporteSeguro))
    cmdData.Parameters.Append cmdData.CreateParameter("@Importe", adCurrency, adParamInput, , mImporte)
    cmdData.Parameters.Append cmdData.CreateParameter("@ImporteContado", adCurrency, adParamInput, , mImporteContado)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDMedioPago", adTinyInt, adParamInput, , IIf(mIDMedioPago = 0, Null, mIDMedioPago))
    cmdData.Parameters.Append cmdData.CreateParameter("@Cuotas", adTinyInt, adParamInput, , IIf(mCuotas = 0, Null, mCuotas))
    cmdData.Parameters.Append cmdData.CreateParameter("@Operacion", adVarChar, adParamInput, 20, IIf(Trim(mOperacion) = "", Null, mOperacion))
    cmdData.Parameters.Append cmdData.CreateParameter("@ImporteCuentaCorriente", adCurrency, adParamInput, , mImporteCuentaCorriente)
    cmdData.Parameters.Append cmdData.CreateParameter("@ImprimirSaldo", adBoolean, adParamInput, , mImprimirSaldo)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDCuentaCorrienteCaja", adInteger, adParamInput, , IIf(mIDCuentaCorrienteCaja = 0, Null, mIDCuentaCorrienteCaja))
    cmdData.Parameters.Append cmdData.CreateParameter("@ForzarDebito", adBoolean, adParamInput, , mForzarDebito)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDPersonaCuentaCorriente", adInteger, adParamInput, , IIf(mIDPersonaCuentaCorriente = 0, Null, mIDPersonaCuentaCorriente))
    cmdData.Parameters.Append cmdData.CreateParameter("@Facturar", adBoolean, adParamInput, , mFacturar)
    cmdData.Parameters.Append cmdData.CreateParameter("@FacturarNotas", adVarChar, adParamInput, 50, IIf(Trim(mFacturarNotas) = "", Null, mFacturarNotas))
    cmdData.Parameters.Append cmdData.CreateParameter("@FacturaNumero", adVarChar, adParamInput, 20, IIf(Trim(mFacturaNumero) = "", Null, mFacturaNumero))
    cmdData.Parameters.Append cmdData.CreateParameter("@IDPersonaRecibe", adInteger, adParamInput, , IIf(mIDPersonaRecibe = 0, Null, mIDPersonaRecibe))
    cmdData.Parameters.Append cmdData.CreateParameter("@PagaQuienRecibe", adBoolean, adParamInput, , mPagaQuienRecibe)
    cmdData.Parameters.Append cmdData.CreateParameter("@Recibe", adVarChar, adParamInput, 50, IIf(Trim(mRecibe) = "", Null, mRecibe))
    cmdData.Parameters.Append cmdData.CreateParameter("@Descripcion", adVarChar, adParamInput, 100, IIf(Trim(mDescripcion) = "", Null, mDescripcion))
    cmdData.Parameters.Append cmdData.CreateParameter("@Domicilio", adVarChar, adParamInput, 100, IIf(Trim(mDomicilio) = "", Null, mDomicilio))
    cmdData.Parameters.Append cmdData.CreateParameter("@Horario", adVarChar, adParamInput, 50, IIf(Trim(mHorario) = "", Null, mHorario))
    cmdData.Parameters.Append cmdData.CreateParameter("@Telefono", adVarChar, adParamInput, 30, IIf(Trim(mTelefono) = "", Null, mTelefono))
    cmdData.Parameters.Append cmdData.CreateParameter("@DejarTraer", adChar, adParamInput, 1, IIf(Trim(mDejarTraer) = "", Null, mDejarTraer))
    cmdData.Parameters.Append cmdData.CreateParameter("@Entregada", adBoolean, adParamInput, , mEntregada)
    cmdData.Parameters.Append cmdData.CreateParameter("@EntregadaFechaHora", adDate, adParamInput, , IIf(mEntregadaFechaHora = DATE_TIME_FIELD_NULL_VALUE, Null, mEntregadaFechaHora))
    cmdData.Parameters.Append cmdData.CreateParameter("@Retira", adVarChar, adParamInput, 50, IIf(Trim(mRetira) = "", Null, mRetira))
    cmdData.Parameters.Append cmdData.CreateParameter("@ReservaTipo", adChar, adParamInput, 2, mReservaTipo)
    cmdData.Parameters.Append cmdData.CreateParameter("@CreadoEnProgreso", adBoolean, adParamInput, , mCreadoEnProgreso)
    cmdData.Parameters.Append cmdData.CreateParameter("@ModificadoEnProgreso", adBoolean, adParamInput, , mModificadoEnProgreso)
    cmdData.Parameters.Append cmdData.CreateParameter("@ReservadoPor", adVarChar, adParamInput, 100, IIf(Trim(mReservadoPor) = "", Null, mReservadoPor))
    cmdData.Parameters.Append cmdData.CreateParameter("@CanceladoPor", adVarChar, adParamInput, 100, IIf(Trim(mCanceladoPor) = "", Null, mCanceladoPor))
    cmdData.Parameters.Append cmdData.CreateParameter("@CanceladoFechaHora", adDate, adParamInput, , IIf(mCanceladoFechaHora = DATE_TIME_FIELD_NULL_VALUE, Null, mCanceladoFechaHora))
    cmdData.Parameters.Append cmdData.CreateParameter("@Notas", adVarChar, adParamInput, 8000, IIf(Trim(mNotas) = "", Null, mNotas))
    If mIsNew Then
        cmdData.Parameters.Append cmdData.CreateParameter("@IDUsuarioCreacion", adChar, adParamInput, 30, pUsuario.IDUsuario)
    Else
        cmdData.Parameters.Append cmdData.CreateParameter("@IDUsuarioModificacion", adChar, adParamInput, 30, pUsuario.IDUsuario)
    End If
    If Not mIsNew Then
        cmdData.Parameters.Append cmdData.CreateParameter("@CuentaCorrienteGrupo_ID_ViajeDebito", adInteger, adParamInput, , pParametro.CuentaCorrienteGrupo_ID_ViajeDebito)
        cmdData.Parameters.Append cmdData.CreateParameter("@CuentaCorrienteGrupo_ID_ViajeCredito", adInteger, adParamInput, , pParametro.CuentaCorrienteGrupo_ID_ViajeCredito)
    End If
    cmdData.Execute 0
    If mIsNew Then
        mIDViajeDetalle = cmdData.Parameters("@IDViajeDetalle").Value
    End If
    mIndice = cmdData.Parameters("@Indice").Value
    Set cmdData = Nothing
    
    If mOcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
        If CambioDeViaje Then
            Set ViajeOriginal = New Viaje
            ViajeOriginal.RefreshListSkip = mRefreshListSkip
            ViajeOriginal.FechaHora = mFechaHoraOriginal
            ViajeOriginal.IDRuta = mIDRutaOriginal
            Call ViajeOriginal.Asiento_Asignar
            Set ViajeOriginal = Nothing
        End If
        If mCallAsientoAsignarOnAddNew Or Not mIsNew Then
            Viaje.RefreshListSkip = mRefreshListSkip
            If mIsNew Then
                Viaje.CuentaCorrienteAsignar = False
                Call Viaje.OrdenAndCuentaCorriente_Asignar
                Call CuentaCorriente_Asignar(Viaje.Estado)
            Else
                Call Viaje.Asiento_Asignar
            End If
            
            If mIsNew And mAsiento > 0 Then
                'ACTUALIZO LOS ASIENTOS DEL VIAJE
                Set cmdData = New ADODB.command
                Set cmdData.ActiveConnection = pDatabase.Connection
                cmdData.CommandText = "sp_ViajeDetalle_AsientoCompartido"
                cmdData.CommandType = adCmdStoredProc
                cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , mFechaHora)
                cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, mIDRuta)
                cmdData.Parameters.Append cmdData.CreateParameter("Asiento_FILTER", adTinyInt, adParamInput, 20, mAsiento)
                cmdData.Parameters.Append cmdData.CreateParameter("AsientoCompartidoCount", adTinyInt, adParamOutput)
                cmdData.Execute
                If cmdData.Parameters("AsientoCompartidoCount").Value = 1 Then
                    Viaje.AsientoOcupado = Viaje.AsientoOcupado + 1
                    Call Viaje.UpdateAsientoOcupado
                End If
                Set cmdData = Nothing
            End If
        End If
    Else
        Viaje.RefreshListSkip = mRefreshListSkip
        Call Viaje.OrdenAndCuentaCorriente_Asignar
    End If

    mIsCopy = False
    mIsDirty = False
    
    If Not mRefreshListSkip Then
        Call RefreshList
    End If
    
    'Chequeo las Reservas Condicionales
    If mIsNew And mVerifyEstadoCondicionalOnAddNew Then
        If Me.Load() Then
            If Estado = VIAJE_DETALLE_ESTADO_CONDICIONAL Then
                Screen.MousePointer = vbDefault
                MsgBox "Esta Reserva se ha generado con Estado Condicional." & vbCr & vbCr & "Fecha/Hora: " & FechaHora_Formatted & vbCr & "Ruta: " & IDRuta & vbCr & "Indice: " & Indice, vbExclamation, App.Title
            End If
        End If
    End If
    
    Set Viaje = Nothing
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            ShowErrorMessage "Classes.ViajeDetalle.Update", "Ya existe un Detalle del Viaje con este Indice." & vbCr & vbCr & "FechaHora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta & vbCr & "Indice: " & Indice
        Else
            ShowErrorMessage "Classes.ViajeDetalle.Update", "Error al actualizar el Detalle del Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta & vbCr & "Indice: " & Indice
        End If
    Else
        ShowErrorMessage "Classes.ViajeDetalle.Update", "Error al actualizar el Detalle del Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta & vbCr & "Indice: " & Indice
    End If
    Set Viaje = Nothing
End Function

Public Function UpdatePagos(ByRef CPagosToAdd As Collection, ByRef CPagosToUpdate As Collection, ByRef CPagosToDelete As Collection)
    Dim Pago As CuentaCorriente
    
    If CPagosToAdd.Count = 0 And CPagosToUpdate.Count = 0 And CPagosToDelete.Count = 0 Then
        UpdatePagos = True
        Exit Function
    End If
    
    'AGREGO LOS PAGOS CORRESPONDIENTES
    For Each Pago In CPagosToAdd
        With Pago
            .Viaje_FechaHora = mFechaHora
            .Viaje_IDRuta = mIDRuta
            .Viaje_Indice = mIndice
            .Descripcion = "Viaje: " & FechaHora_Formatted & " - " & mIDRuta
            If mPagaQuienRecibe Then
                .IDPersona = IIf(mIDPersonaCuentaCorriente = 0, mIDPersonaRecibe, mIDPersonaCuentaCorriente)
                .IDPersonaOrigen = IIf(mIDPersonaCuentaCorriente = 0, 0, mIDPersonaRecibe)
            Else
                .IDPersona = IIf(mIDPersonaCuentaCorriente = 0, mIDPersona, mIDPersonaCuentaCorriente)
                .IDPersonaOrigen = IIf(mIDPersonaCuentaCorriente = 0, 0, mIDPersona)
            End If
            If Not .Update(True) Then
                Exit Function
            End If
        End With
    Next Pago
    
    'ACTUALIZO LOS PAGOS CORRESPONDIENTES
    For Each Pago In CPagosToUpdate
        With Pago
            .Viaje_FechaHora = mFechaHora
            .Viaje_IDRuta = mIDRuta
            .Viaje_Indice = mIndice
            .Descripcion = "Viaje: " & FechaHora_Formatted & " - " & mIDRuta
            If mPagaQuienRecibe Then
                .IDPersona = IIf(mIDPersonaCuentaCorriente = 0, mIDPersonaRecibe, mIDPersonaCuentaCorriente)
                .IDPersonaOrigen = IIf(mIDPersonaCuentaCorriente = 0, 0, mIDPersonaRecibe)
            Else
                .IDPersona = IIf(mIDPersonaCuentaCorriente = 0, mIDPersona, mIDPersonaCuentaCorriente)
                .IDPersonaOrigen = IIf(mIDPersonaCuentaCorriente = 0, 0, mIDPersona)
            End If
            If Not .Update(True) Then
                Exit Function
            End If
        End With
    Next Pago
    
    'ELIMINO LOS PAGOS CORRESPONDIENTES
    For Each Pago In CPagosToDelete
        If Not Pago.Delete Then
            Exit Function
        End If
    Next Pago
    
    UpdatePagos = True
End Function

Public Function CambiarEstado(ByVal BuscarConexion As Boolean) As Boolean
    Dim cmdData As ADODB.command
    
    Dim Viaje As Viaje
    Dim ViajeDetalle_Original As ViajeDetalle
    Dim ViajeDetalle_Conexion As ViajeDetalle_Conexion
    Dim CambiaConexionTramo1 As Boolean
    Dim CambiaConexionTramo2 As Boolean
    
    Dim VerificarOtrosPrepagos As Boolean
    Dim VerificarEstePrepago As Boolean
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    'CREADO EN PROGRESO
    If Not CreadoEnProgreso Then
        Set Viaje = New Viaje
        Viaje.FechaHora = mFechaHora
        Viaje.IDRuta = mIDRuta
        Call Viaje.Load
        mCreadoEnProgreso = (Viaje.Estado = VIAJE_ESTADO_EN_PROGRESO)
        Set Viaje = Nothing
    End If

    If mOcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
        If mEstado = "" Then
            VerificarEstePrepago = True
            mrecData("Estado").Value = Null
            
            'PRIORIDAD MAXIMA
            Set cmdData = New ADODB.command
            Set cmdData.ActiveConnection = pDatabase.Connection
            cmdData.CommandType = adCmdStoredProc
            cmdData.Parameters.Append cmdData.CreateParameter("FechaHora", adDate, adParamInput, , mFechaHora)
            cmdData.Parameters.Append cmdData.CreateParameter("IDRuta", adChar, adParamInput, 20, mIDRuta)
            cmdData.Parameters.Append cmdData.CreateParameter("Prioridad", adInteger, adParamOutput)
            cmdData.CommandText = "sp_ViajeDetalle_GetPrioridadNew"
            cmdData.Execute 0
            mPrioridad = cmdData.Parameters("Prioridad").Value
            Set cmdData = Nothing
            
            mrecData("Prioridad").Value = mPrioridad
            mrecData("ForzarDebito").Value = False
            mrecData("CanceladoPor").Value = Null
            mrecData("CanceladoFechaHora").Value = Null
        Else
            If mEstado = VIAJE_DETALLE_ESTADO_CANCELADO Then
                'SI CANCELÓ LA RESERVA, VERIFICO SI APLICA ALGUN PREPAGO (CON LÍMITE DE CANTIDAD) A OTRAS RESERVAS EN ESOS DIAS
                VerificarOtrosPrepagos = True
            ElseIf mrecData("Estado").Value = VIAJE_DETALLE_ESTADO_CANCELADO Then
                'SI REACTIVÓ UNA RESERVA CANCELADA, VERIFICO SI APLICA UN PREPAGO
                VerificarEstePrepago = True
            End If
            mrecData("Estado").Value = mEstado
            mrecData("Asiento").Value = Null
            mrecData("Prioridad").Value = Null
            mrecData("ForzarDebito").Value = mForzarDebito
            mrecData("CanceladoPor").Value = IIf(mCanceladoPor = "", Null, mCanceladoPor)
            mrecData("CanceladoFechaHora").Value = Now
            mrecData("IDUsuarioCancelacion").Value = pUsuario.IDUsuario
        End If
        mrecData("CreadoEnProgreso").Value = mCreadoEnProgreso
        mrecData("FechaHoraModificacion").Value = Now
        mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
        mrecData.Update
        
        Call ViajeDetalle_ShowViajeVuelta(Me, "Hay %1 Reserva(s) para este Pasajero en el mismo Día.")
    
        Set Viaje = New Viaje
        Viaje.RefreshListSkip = mRefreshListSkip
        Viaje.FechaHora = mFechaHora
        Viaje.IDRuta = mIDRuta
        Call Viaje.Asiento_Asignar
        Set Viaje = Nothing
    Else
        mrecData("Estado").Value = mEstado
        mrecData("ForzarDebito").Value = mForzarDebito
        mrecData("CanceladoPor").Value = IIf(mCanceladoPor = "", Null, mCanceladoPor)
        mrecData("CanceladoFechaHora").Value = Now
        mrecData("FechaHoraModificacion").Value = Now
        mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
        mrecData("IDUsuarioCancelacion").Value = pUsuario.IDUsuario
        mrecData.Update
        
        Set Viaje = New Viaje
        Viaje.RefreshListSkip = mRefreshListSkip
        Viaje.FechaHora = mFechaHora
        Viaje.IDRuta = mIDRuta
        Call Viaje.OrdenAndCuentaCorriente_Asignar
        Set Viaje = Nothing
    End If
    
    If BuscarConexion Then
        Set ViajeDetalle_Conexion = New ViajeDetalle_Conexion
        ViajeDetalle_Conexion.IDViajeDetalle = mIDViajeDetalle
        ViajeDetalle_Conexion.NoMatchRaiseError = False
        If ViajeDetalle_Conexion.Load() Then
            If Not ViajeDetalle_Conexion.NoMatch Then
                CambiaConexionTramo2 = (MsgBox("Hay otra " & IIf(mOcupanteTipo = OCUPANTE_TIPO_PASAJERO, "Reserva", "Comisión") & " relacionada a través de una Conexión de Rutas." & vbCr & vbCr & "Fecha/Hora: " & ViajeDetalle_Conexion.Conexion_ViajeDetalle.FechaHora_Formatted & " - " & RTrim(ViajeDetalle_Conexion.Conexion_ViajeDetalle.IDRuta) & vbCr & vbCr & "¿Desea cambiarle el estado también?", vbExclamation + vbYesNo, App.Title) = vbYes)
            Else
                ViajeDetalle_Conexion.Conexion_IDViajeDetalle = mIDViajeDetalle
                If ViajeDetalle_Conexion.LoadByConexion() Then
                    If Not ViajeDetalle_Conexion.NoMatch Then
                        CambiaConexionTramo1 = (MsgBox("Hay otra " & IIf(mOcupanteTipo = OCUPANTE_TIPO_PASAJERO, "Reserva", "Comisión") & " relacionada a través de una Conexión de Rutas." & vbCr & vbCr & "Fecha/Hora: " & ViajeDetalle_Conexion.ViajeDetalle.FechaHora_Formatted & " - " & RTrim(ViajeDetalle_Conexion.ViajeDetalle.IDRuta) & vbCr & vbCr & "¿Desea cambiarle el estado también?", vbExclamation + vbYesNo, App.Title) = vbYes)
                    End If
                End If
            End If
        End If
        
        If CambiaConexionTramo1 Or CambiaConexionTramo2 Then
            Set ViajeDetalle_Original = New ViajeDetalle
        
            If CambiaConexionTramo1 Then
                ViajeDetalle_Original.IDViajeDetalle = ViajeDetalle_Conexion.IDViajeDetalle
            ElseIf CambiaConexionTramo2 Then
                ViajeDetalle_Original.IDViajeDetalle = ViajeDetalle_Conexion.Conexion_IDViajeDetalle
            End If

            If ViajeDetalle_Original.LoadByIDViajeDetalle() Then
                ViajeDetalle_Original.Estado = mEstado
                ViajeDetalle_Original.CanceladoPor = mCanceladoPor
                ViajeDetalle_Original.CanceladoFechaHora = mCanceladoFechaHora
                ViajeDetalle_Original.ForzarDebito = mForzarDebito
                If Not ViajeDetalle_Original.CambiarEstado(False) Then
                    Screen.MousePointer = vbDefault
                    Set ViajeDetalle_Original = Nothing
                    Exit Function
                End If
            End If
            Set ViajeDetalle_Original = Nothing

        End If
    
        Set ViajeDetalle_Conexion = Nothing
    End If
    
    'VERIFICO EL TEMA PREPAGOS
    If VerificarOtrosPrepagos Then
        Call VerificarPrepagoOtrasReservas
    End If
    If VerificarEstePrepago Then
        Call VerificarPrepagoEstaReserva
    End If
    
    Screen.MousePointer = vbDefault
    CambiarEstado = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.ViajeDetale.CambiarEstado", "Error al Cambiar el Estado del Detalle del Viaje." & vbCr & vbCr & "Fecha/Hora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta & vbCr & "Indice: " & Indice
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Public Function Realizar() As Boolean
    Dim cmdData As ADODB.command
    Dim Viaje As Viaje
    Dim ViajeEstado As String
    Dim RealizadoOriginal As Long
    Dim Asignar As Boolean
    Dim StartTime As Long
    Dim RandomDelayMilliseconds As Long
    Dim ErrorCounter As Long

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If mImporteContado > 0 And IDCuentaCorrienteCaja = 0 Then
        MsgBox "Debe especificar la Caja.", vbInformation, App.Title
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    RealizadoOriginal = IIf(IsNull(mrecData("Realizado").Value), 0, IIf(mrecData("Realizado").Value, 1, 2))
    
'    If mImporte = 0 Or Viaje.Estado = VIAJE_ESTADO_ACTIVO Or Viaje.Estado = VIAJE_ESTADO_CANCELADO Or (mEstado = VIAJE_DETALLE_ESTADO_CONFIRMADO And mRealizado = VIAJE_DETALLE_REALIZADO_NO And Not mForzarDebito) Or mEstado = VIAJE_DETALLE_ESTADO_CONDICIONAL Or (mEstado = VIAJE_DETALLE_ESTADO_CANCELADO And Not mForzarDebito) Then
'        mImporteCuentaCorriente = 0
'    Else
'        mImporteCuentaCorriente = CalcularImporteCuentaCorrienteExceptCurrent()
'    End If
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ViajeDetalle_Realizar"
    cmdData.CommandType = adCmdStoredProc
    cmdData.NamedParameters = True
    cmdData.Parameters.Append cmdData.CreateParameter("@FechaHora_FILTER", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDRuta_FILTER", adChar, adParamInput, 20, mIDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("@Indice_FILTER", adInteger, adParamInput, , mIndice)
    cmdData.Parameters.Append cmdData.CreateParameter("@Realizado", adBoolean, adParamInput, , IIf(mRealizado = 0, Null, IIf(mRealizado = 1, True, False)))
    cmdData.Parameters.Append cmdData.CreateParameter("@ForzarDebito", adBoolean, adParamInput, , mForzarDebito)
    cmdData.Parameters.Append cmdData.CreateParameter("@Entregada", adBoolean, adParamInput, , mEntregada)
    cmdData.Parameters.Append cmdData.CreateParameter("@EntregadaFechaHora", adDate, adParamInput, , mEntregadaFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("@Retira", adVarChar, adParamInput, 50, IIf(Trim(mRetira) = "", Null, mRetira))
    cmdData.Parameters.Append cmdData.CreateParameter("@FacturaNumero", adVarChar, adParamInput, 20, IIf(Trim(mFacturaNumero) = "", Null, mFacturaNumero))
    cmdData.Parameters.Append cmdData.CreateParameter("@ImporteContado", adCurrency, adParamInput, , mImporteContado)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDMedioPago", adTinyInt, adParamInput, , IIf(mIDMedioPago = 0, Null, mIDMedioPago))
    cmdData.Parameters.Append cmdData.CreateParameter("@Cuotas", adTinyInt, adParamInput, , IIf(mCuotas = 0, Null, mCuotas))
    cmdData.Parameters.Append cmdData.CreateParameter("@Operacion", adVarChar, adParamInput, 20, IIf(Trim(mOperacion) = "", Null, mOperacion))
    cmdData.Parameters.Append cmdData.CreateParameter("@ImporteCuentaCorriente", adCurrency, adParamInput, , mImporteCuentaCorriente)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDCuentaCorrienteCaja", adInteger, adParamInput, , IIf(mIDCuentaCorrienteCaja = 0, Null, mIDCuentaCorrienteCaja))
    cmdData.Parameters.Append cmdData.CreateParameter("@FechaHoraSistema", adDate, adParamInput, , Now)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDUsuarioModificacion", adChar, adParamInput, 30, pUsuario.IDUsuario)
    cmdData.Execute 0
    Set cmdData = Nothing
    
    Set Viaje = New Viaje
    Viaje.RefreshListSkip = mRefreshListSkip
    Viaje.FechaHora = mFechaHora
    Viaje.IDRuta = mIDRuta
    If Viaje.Load Then
        ViajeEstado = Viaje.Estado
    End If
    
    If (RealizadoOriginal <> 2 And mRealizado = 2) Or (RealizadoOriginal = 2 And mRealizado <> 2) Then
        If mOcupanteTipo = OCUPANTE_TIPO_PASAJERO And mCallAsientoAsignarOnAddNew Then
            Call Viaje.Asiento_Asignar
        Else
            Asignar = True
        End If
    Else
        Asignar = True
    End If
    Set Viaje = Nothing
    
    If Not mRefreshListSkip Then
        RefreshList_RefreshViajeDetalle mFechaHora, mIDRuta, 0
    End If
    
    If Asignar Then
        Do While ErrorCounter < 3
            If CuentaCorriente_Asignar(ViajeEstado) Then
                Exit Do
            Else
                ErrorCounter = ErrorCounter + 1
                'HUBO ALGUN ERROR, PERO VOY A REINTENTAR 3 VECES
                'POR LO TANTO, HAGO UNA PAUSA AL AZAR POR SI PASA EN DOS PC,
                'PARA QUE NO SE QUEDEN EN UN DEAD LOCK INFINITO Y
                'REINICIO LA RUTINA
                
                Randomize Timer
                RandomDelayMilliseconds = Int(Rnd(1) * 2000) + 500
                
                StartTime = timeGetTime
                Do While (timeGetTime() - StartTime) < RandomDelayMilliseconds
                    
                Loop
            End If
        Loop
    End If
        
    Screen.MousePointer = vbDefault
    Realizar = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        ShowErrorMessage "Classes.ViajeDetale.Realizar", "Error al Actualizar el Detalle del Viaje." & vbCr & vbCr & "Fecha/Hora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta & vbCr & "Indice: " & Indice
    End If
End Function

Public Function Delete() As Boolean
    Dim cmdData As ADODB.command
    Dim Viaje As Viaje
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ViajeDetalle_Delete"
    cmdData.CommandType = adCmdStoredProc
    cmdData.NamedParameters = True
    cmdData.Parameters.Append cmdData.CreateParameter("@FechaHora", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDRuta", adChar, adParamInput, 20, mIDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("@Indice", adInteger, adParamInput, , mIndice)
    cmdData.Execute 0
    Set cmdData = Nothing
    
    Set Viaje = New Viaje
    Viaje.RefreshListSkip = mRefreshListSkip
    Viaje.FechaHora = mFechaHora
    Viaje.IDRuta = mIDRuta
    If mOcupanteTipo = OCUPANTE_TIPO_PASAJERO Then
        Call Viaje.Asiento_Asignar
    Else
        Call Viaje.OrdenAndCuentaCorriente_Asignar
    End If
    Set Viaje = Nothing
    
    Call VerificarPrepagoOtrasReservas
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede eliminar el Detalle del Viaje debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.ViajeDetalle.Delete", "Error al eliminar el Detalle del Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta & vbCr & "Indice: " & Indice
        End If
    Else
        ShowErrorMessage "Classes.ViajeDetalle.Delete", "Error al eliminar el Detalle del Viaje." & vbCr & vbCr & "FechaHora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta & vbCr & "Indice: " & Indice
    End If
End Function

Private Sub InitializeValues()
    mGrupoNumero = 0
    mReservaCodigo = ""
    mOcupanteTipo = ""
    mEstado = ""
    mPrioridad = 0
    mOrden = 0
    mAsiento = 0
    mRealizado = 0
    mIDPersona = 0
    mIDPersonaMenor = 0
    mIDListaPrecio = pParametro.ListaPrecio_ID_Predeterminada
    mIDOrigen = 0
    mSube = ""
    mIDDestino = 0
    mBaja = ""
    mValorDeclarado = 0
    mImporteSeguro = 0
    mImporte = 0
    mImporteContado = 0
    mIDMedioPago = pParametro.MedioPago_Predeterminado_ID
    mCuotas = 1
    mOperacion = ""
    mImporteCuentaCorriente = 0
    mImprimirSaldo = False
    mIDCuentaCorrienteCaja = 0
    mForzarDebito = False
    mIDPersonaCuentaCorriente = 0
    mFacturar = False
    mFacturarNotas = ""
    mFacturaNumero = ""
    mIDPersonaRecibe = 0
    mPagaQuienRecibe = False
    mRecibe = ""
    mDescripcion = ""
    mDomicilio = ""
    mHorario = ""
    mTelefono = ""
    mDejarTraer = ""
    mEntregada = False
    mEntregadaFechaHora = DATE_TIME_FIELD_NULL_VALUE
    mRetira = ""
    mReservaTipo = ""
    mCreadoEnProgreso = False
    mModificadoEnProgreso = False
    mReservadoPor = ""
    mCanceladoPor = ""
    mCanceladoFechaHora = DATE_TIME_FIELD_NULL_VALUE
    mNotas = ""
    mFechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioCreacion = 0
    mFechaHoraModificacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioModificacion = 0
    mIDUsuarioCancelacion = ""
    
    Set mCPagos = New Collection

    mIsNew = True
    mIsCopy = False
    mIsDirty = False
End Sub

Private Sub Class_Initialize()
    InitializeValues
    mNoMatchRaiseError = True
    mRefreshListSkip = False
    CallAsientoAsignarOnAddNew = True
    VerifyDuplicates = True
End Sub

Private Sub Class_Terminate()
    If Not mrecData Is Nothing Then
        If mrecData.State = adStateOpen Then
            mrecData.Close
        End If
        Set mrecData = Nothing
    End If
End Sub

Public Function GetNewValues(ByVal DiaSemanaBase As Integer) As Boolean
    Dim Persona As Persona
    Dim PersonaHorario As PersonaHorario
    Dim PersonaRuta As PersonaRuta
    Dim Ruta As Ruta
    Dim ListaPrecioDetalle As ListaPrecioDetalle
    
    'OBTENGO LOS DATOS DE LA PERSONA
    Set Persona = New Persona
    Persona.IDPersona = mIDPersona
    If Not Persona.Load() Then
        Set Persona = Nothing
        Exit Function
    End If
    IDPersonaCuentaCorriente = Persona.IDPersonaCuentaCorriente
    Set Persona = Nothing
    
    'OBTENGO LOS DATOS DEL HORARIO DE LA PERSONA
    Set PersonaHorario = New PersonaHorario
    PersonaHorario.NoMatchRaiseError = False
    PersonaHorario.IDPersona = mIDPersona
    PersonaHorario.DiaSemana = DiaSemanaBase
    PersonaHorario.Hora = FechaHora_FormattedAsTime
    PersonaHorario.IDRuta = mIDRuta
    If Not PersonaHorario.Load() Then
        Set PersonaHorario = Nothing
        Exit Function
    End If
    If Not PersonaHorario.NoMatch Then
        'TIENE ASIGNADO UN HORARIO
        mIDOrigen = PersonaHorario.IDOrigen
        mSube = PersonaHorario.Sube
        mIDDestino = PersonaHorario.IDDestino
        mBaja = PersonaHorario.Baja
        
        Set PersonaRuta = New PersonaRuta
        PersonaRuta.NoMatchRaiseError = False
        PersonaRuta.IDPersona = mIDPersona
        PersonaRuta.IDRuta = mIDRuta
        If Not PersonaRuta.Load() Then
            Set PersonaHorario = Nothing
            Set PersonaRuta = Nothing
            Exit Function
        End If
        If Not PersonaRuta.NoMatch Then
            If mIDOrigen = 0 Then
                mIDOrigen = PersonaRuta.IDOrigen
                mSube = PersonaRuta.Sube
            End If
            If mIDDestino = 0 Then
                mIDDestino = PersonaRuta.IDDestino
                mBaja = PersonaRuta.Baja
            End If
            
            'VERIFICO SI APLICA UN PREPAGO
            Call VerificarListaPrecioPorPrepago(PersonaRuta.IDListaPrecio)
        Else
            If mIDOrigen = 0 Or mIDDestino = 0 Then
                Set Ruta = New Ruta
                Ruta.IDRuta = mIDRuta
                If Not Ruta.Load() Then
                    Set PersonaHorario = Nothing
                    Set PersonaRuta = Nothing
                    Set Ruta = Nothing
                    Exit Function
                End If
                mIDOrigen = Ruta.IDOrigen
                mIDDestino = Ruta.IDDestino
                Set Ruta = Nothing
            End If
            
            'mIDListaPrecio = pParametro.ListaPrecio_ID_Predeterminada
            Call VerificarListaPrecioPorPrepago(pParametro.ListaPrecio_ID_Predeterminada)
        End If
        Set PersonaRuta = Nothing
    Else
        Set PersonaRuta = New PersonaRuta
        PersonaRuta.NoMatchRaiseError = False
        PersonaRuta.IDPersona = mIDPersona
        PersonaRuta.IDRuta = mIDRuta
        If Not PersonaRuta.Load() Then
            Set PersonaHorario = Nothing
            Set PersonaRuta = Nothing
            Exit Function
        End If
        If Not PersonaRuta.NoMatch Then
            mIDOrigen = PersonaRuta.IDOrigen
            mSube = PersonaRuta.Sube
            mIDDestino = PersonaRuta.IDDestino
            mBaja = PersonaRuta.Baja
        
            'VERIFICO SI APLICA UN PREPAGO
            Call VerificarListaPrecioPorPrepago(PersonaRuta.IDListaPrecio)
        Else
            Set Ruta = New Ruta
            Ruta.IDRuta = mIDRuta
            If Not Ruta.Load() Then
                Set PersonaHorario = Nothing
                Set PersonaRuta = Nothing
                Set Ruta = Nothing
                Exit Function
            End If
            mIDOrigen = Ruta.IDOrigen
            mIDDestino = Ruta.IDDestino
            Set Ruta = Nothing
            
            'mIDListaPrecio = pParametro.ListaPrecio_ID_Predeterminada
            Call VerificarListaPrecioPorPrepago(pParametro.ListaPrecio_ID_Predeterminada)
        End If
        Set PersonaRuta = Nothing
    End If
    Set PersonaHorario = Nothing
    
    Set ListaPrecioDetalle = New ListaPrecioDetalle
    ListaPrecioDetalle.IDListaPrecio = mIDListaPrecio
    ListaPrecioDetalle.OcupanteTipo = mOcupanteTipo
    ListaPrecioDetalle.IDRuta = mIDRuta
    ListaPrecioDetalle.IDOrigen = mIDOrigen
    ListaPrecioDetalle.IDDestino = mIDDestino
    If Not ListaPrecioDetalle.GetImporteByLugar() Then
        Exit Function
    End If
    mImporte = ListaPrecioDetalle.Importe
    
    GetNewValues = True
    Exit Function
    
ErrorHandler:
    CSM_Error.ShowErrorMessage "Classes.ViajeDetalle.GetNewValues", "Error al verificar los valores de la Reserva." & vbCr & vbCr & "IDPersona: " & mIDPersona & vbCr & "IDRuta: " & mIDRuta & vbCr & "Fecha/Hora: " & mFechaHora
End Function

Public Function VerificarListaPrecioPorPrepago(ByVal IDListaPrecioNormal As Long)
    Dim cmdPrepago As ADODB.command
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    'VERIFICO SI APLICA UN PREPAGO
    Set cmdPrepago = New ADODB.command
    Set cmdPrepago.ActiveConnection = pDatabase.Connection
    cmdPrepago.CommandText = "usp_PersonaPrepago_CheckIfApply"
    cmdPrepago.CommandType = adCmdStoredProc
    cmdPrepago.Parameters.Append cmdPrepago.CreateParameter("IDPersona", adInteger, adParamInput, , mIDPersona)
    cmdPrepago.Parameters.Append cmdPrepago.CreateParameter("IDRuta", adChar, adParamInput, 20, mIDRuta)
    cmdPrepago.Parameters.Append cmdPrepago.CreateParameter("Fecha", adDate, adParamInput, , mFechaHora)
    cmdPrepago.Parameters.Append cmdPrepago.CreateParameter("IDListaPrecio", adInteger, adParamOutput)
    cmdPrepago.Execute
    If Val(cmdPrepago.Parameters("IDListaPrecio").Value) <> 0 Then
        IDListaPrecio = Val(cmdPrepago.Parameters("IDListaPrecio").Value)
    Else
        IDListaPrecio = IDListaPrecioNormal
    End If
    Set cmdPrepago = Nothing
    Exit Function
    
ErrorHandler:
    CSM_Error.ShowErrorMessage "Classes.ViajeDetalle.VerificarListaPrecioPorPrepago", "Error al verificar la Lista de Precios de la Reserva." & vbCr & vbCr & "IDPersona: " & mIDPersona & vbCr & "IDRuta: " & mIDRuta & vbCr & "Fecha/Hora: " & mFechaHora
End Function

Private Function CalcularImporteCuentaCorrienteExceptCurrent() As Currency
    Dim ImporteRestante As Currency
    Dim Persona As Persona
    
    If mImporteContado >= mImporte Then
        ImporteRestante = 0
    Else
        ImporteRestante = mImporte
    End If
    
    'SALDO ACTUAL
    Set Persona = New Persona
    If IDPersonaCuentaCorriente = 0 Then
        Persona.IDPersona = mIDPersona
    Else
        Persona.IDPersona = mIDPersonaCuentaCorriente
    End If
    Persona.ViajeActual_FechaHora = mFechaHora
    Persona.ViajeActual_IDRuta = mIDRuta
    Persona.ViajeActual_Indice = mIndice
    If Persona.LoadSaldoActual() Then
        If Persona.SaldoActual > 0 Then
            If Persona.SaldoActual >= ImporteRestante Then
                CalcularImporteCuentaCorrienteExceptCurrent = ImporteRestante
            Else
                CalcularImporteCuentaCorrienteExceptCurrent = Persona.SaldoActual
            End If
        Else
            CalcularImporteCuentaCorrienteExceptCurrent = 0
        End If
    End If
    Set Persona = Nothing
End Function

Public Function CuentaCorriente_Asignar(ByVal ViajeEstado As String) As Boolean
    Dim ImporteTemp As Currency
    Dim CuentaCorrientePago As CuentaCorriente
    
    Screen.MousePointer = vbHourglass
    
    'CREDITO
    If pParametro.ViajeDetalle_Paquete_Permite_Multiples_Pagos And mIDRuta = pParametro.Ruta_Paquete_ID Then
        'ES UN PAQUETE Y PERMITE MULTIPLES PAGOS
    Else
        'NO ES PAQUETE O NO SE PERMITEN MULTIPLES PAGOS
        If mImporteContado = 0 Then
            If Not CuentaCorriente_Asignar_Movimiento_Eliminar(pParametro.CuentaCorrienteGrupo_ID_ViajeCredito) Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        Else
            If Not CuentaCorriente_Asignar_Movimiento_Crear(pParametro.CuentaCorrienteGrupo_ID_ViajeCredito, mIDCuentaCorrienteCaja, IIf(mOcupanteTipo = OCUPANTE_TIPO_PASAJERO, "Viaje: " & FechaHora_Formatted & " - " & mIDRuta, "Comisión: " & FechaHora_Formatted & " - " & mIDRuta), mImporteContado, mIDMedioPago, mCuotas, mOperacion) Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
    End If
        
    'DEBITO
    If mImporte = 0 Or ((pParametro.ViajeDetalle_Paquete_Permite_Multiples_Pagos = False Or mIDRuta <> pParametro.Ruta_Paquete_ID) And (ViajeEstado = VIAJE_ESTADO_ACTIVO Or ViajeEstado = VIAJE_ESTADO_CANCELADO Or (mEstado = VIAJE_DETALLE_ESTADO_CONFIRMADO And mRealizado = VIAJE_DETALLE_REALIZADO_NO And Not mForzarDebito) Or mEstado = VIAJE_DETALLE_ESTADO_CONDICIONAL Or (mEstado = VIAJE_DETALLE_ESTADO_CANCELADO And Not mForzarDebito))) Then
        If mImporteCuentaCorriente <> 0 Then
            mImporteCuentaCorriente = 0
            pDatabase.Connection.Execute "UPDATE ViajeDetalle SET ImporteCuentaCorriente = " & CSM_String.FormatCurrencyToSQL(mImporteCuentaCorriente) & " WHERE FechaHora = '" & Format(mFechaHora, "yyyy/mm/dd hh:nn") & "' AND IDRuta = '" & ReplaceQuote(mIDRuta) & "' AND Indice = " & mIndice, 0
        End If
        If Not CuentaCorriente_Asignar_Movimiento_Eliminar(pParametro.CuentaCorrienteGrupo_ID_ViajeDebito) Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    Else
        ImporteTemp = CalcularImporteCuentaCorrienteExceptCurrent()
        If ImporteTemp <> mImporteCuentaCorriente Then
            mImporteCuentaCorriente = ImporteTemp
            pDatabase.Connection.Execute "UPDATE ViajeDetalle SET ImporteCuentaCorriente = " & CSM_String.FormatCurrencyToSQL(mImporteCuentaCorriente) & " WHERE FechaHora = '" & Format(mFechaHora, "yyyy/mm/dd hh:nn") & "' AND IDRuta = '" & ReplaceQuote(mIDRuta) & "' AND Indice = " & mIndice, 0
        End If
        If Not CuentaCorriente_Asignar_Movimiento_Crear(pParametro.CuentaCorrienteGrupo_ID_ViajeDebito, pParametro.CuentaCorrienteCaja_ID_ViajeDebito, IIf(mOcupanteTipo = OCUPANTE_TIPO_PASAJERO, "Viaje: " & FechaHora_Formatted & " - " & mIDRuta, "Comisión: " & FechaHora_Formatted & " - " & mIDRuta), mImporte * -1, pParametro.MedioPago_Predeterminado_ID, 1, "") Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    If Not mRefreshListSkip Then
        RefreshList
    End If
    
    Screen.MousePointer = vbDefault
    CuentaCorriente_Asignar = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.ViajeDetalle.CuentaCorriente_Asignar", "Error al asignar los Movimientos de Cuenta Corriente." & vbCr & vbCr & "Fecha/Hora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta & vbCr & "Indice: " & Indice
End Function

Private Function CuentaCorriente_Asignar_Movimiento_Crear(ByVal IDGrupo As Long, ByVal IDCaja As Long, ByVal Descripcion As String, ByVal ImporteCC As Currency, ByVal IDMedioPagoMov As Byte, ByVal CuotasMov As Byte, ByVal OperacionMov As String) As Boolean
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
        .Viaje_Indice = mIndice
        .NoMatchRaiseError = False
        .RefreshListSkip = mRefreshListSkip
        If .LoadByViajeDetalle() Then
            .IDCuentaCorrienteCaja = IDCaja
            If mPagaQuienRecibe Then
                .IDPersona = IIf(mIDPersonaCuentaCorriente = 0, mIDPersonaRecibe, mIDPersonaCuentaCorriente)
                .IDPersonaOrigen = IIf(mIDPersonaCuentaCorriente = 0, 0, mIDPersonaRecibe)
            Else
                .IDPersona = IIf(mIDPersonaCuentaCorriente = 0, mIDPersona, mIDPersonaCuentaCorriente)
                .IDPersonaOrigen = IIf(mIDPersonaCuentaCorriente = 0, 0, mIDPersona)
            End If
            If .NoMatch Then
                .FechaHora = Now
            End If
            .Descripcion = Descripcion
            .Importe = ImporteCC
            .IDMedioPago = IDMedioPagoMov
            .Cuotas = CuotasMov
            .Operacion = OperacionMov
            .Viaje_FechaHora = mFechaHora
            .Viaje_IDRuta = mIDRuta
            .Viaje_Indice = mIndice
            .Update
        End If
    End With
    Set CuentaCorriente = Nothing
        
    Screen.MousePointer = vbDefault
    CuentaCorriente_Asignar_Movimiento_Crear = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Viaje.CuentaCorriente_Asignar_Movimiento_Crear", "Error al Crear el Movimiento en la Cuenta Corriente." & vbCr & vbCr & "Fecha/Hora: " & FechaHora_Formatted & vbCr & "IDRuta: " & IDRuta
End Function

Private Function CuentaCorriente_Asignar_Movimiento_Eliminar(ByVal IDGrupo As Long) As Boolean
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
        .Viaje_Indice = mIndice
        .NoMatchRaiseError = False
        .RefreshListSkip = mRefreshListSkip
        If .LoadByViajeDetalle() Then
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

Public Sub VerificarPrepagoOtrasReservas()
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    Dim IDListaPrecioOriginal As Long
    
    Dim ViajeDetalleVerificar As ViajeDetalle
    Dim ListaPrecioDetalle As ListaPrecioDetalle
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If Not (ListaPrecio.PrepagoEs And ListaPrecio.PrepagoReservasCantidad > 0) Then
        Exit Sub
    End If
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_PersonaPrepago_GetViajeDetalleListForRecount"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona", adInteger, adParamInput, , mIDPersona)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta", adVarChar, adParamInput, 20, mIDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("Fecha", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDListaPrecio", adInteger, adParamInput, , mIDListaPrecio)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    
    Do While Not recData.EOF
        Set ViajeDetalleVerificar = New ViajeDetalle
        ViajeDetalleVerificar.FechaHora = recData("FechaHora").Value
        ViajeDetalleVerificar.IDRuta = RTrim(recData("IDRuta").Value)
        ViajeDetalleVerificar.Indice = recData("Indice").Value
        IDListaPrecioOriginal = recData("IDListaPrecio").Value
        If ViajeDetalleVerificar.Load Then
            
            Call ViajeDetalleVerificar.VerificarListaPrecioPorPrepago(ViajeDetalleVerificar.IDListaPrecio)
            
            If ViajeDetalleVerificar.IDListaPrecio <> IDListaPrecioOriginal Then
                Set ListaPrecioDetalle = New ListaPrecioDetalle
                ListaPrecioDetalle.IDListaPrecio = ViajeDetalleVerificar.IDListaPrecio
                ListaPrecioDetalle.OcupanteTipo = ViajeDetalleVerificar.OcupanteTipo
                ListaPrecioDetalle.IDRuta = ViajeDetalleVerificar.IDRuta
                ListaPrecioDetalle.IDOrigen = ViajeDetalleVerificar.IDOrigen
                ListaPrecioDetalle.IDDestino = ViajeDetalleVerificar.IDDestino
                If ListaPrecioDetalle.GetImporteByLugar() Then
                    ViajeDetalleVerificar.Importe = ListaPrecioDetalle.Importe
                End If
                Set ListaPrecioDetalle = Nothing
                
                ViajeDetalleVerificar.Update
            End If
        End If
        Set ViajeDetalleVerificar = Nothing
    
        recData.MoveNext
    Loop
    
    recData.Close
    Set recData = Nothing
    Exit Sub
    
ErrorHandler:
    CSM_Error.ShowErrorMessage "Classes.ViajeDetalle.VerificarPrepagoOtrasReservas", "Error al verificar si se aplica un Prepago a otras reservas." & vbCr & vbCr & "IDPersona: " & mIDPersona & vbCr & "IDRuta: " & mIDRuta & vbCr & "Fecha/Hora: " & mFechaHora
End Sub

Public Sub VerificarPrepagoEstaReserva()
    Dim IDListaPrecioOriginal As Long
    Dim ListaPrecioDetalle As ListaPrecioDetalle
    
    IDListaPrecioOriginal = mIDListaPrecio
    
    Call VerificarListaPrecioPorPrepago(pParametro.ListaPrecio_ID_Predeterminada)
    
    If mIDListaPrecio <> IDListaPrecioOriginal Then
        Set ListaPrecioDetalle = New ListaPrecioDetalle
        ListaPrecioDetalle.IDListaPrecio = mIDListaPrecio
        ListaPrecioDetalle.OcupanteTipo = mOcupanteTipo
        ListaPrecioDetalle.IDRuta = mIDRuta
        ListaPrecioDetalle.IDOrigen = mIDOrigen
        ListaPrecioDetalle.IDDestino = mIDDestino
        If ListaPrecioDetalle.GetImporteByLugar() Then
            Importe = ListaPrecioDetalle.Importe
        End If
        Set ListaPrecioDetalle = Nothing
            
        Update
    End If
End Sub

Public Function ObtenerPrepagosRestantes() As Integer
    Dim cmdData As ADODB.command
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_PersonaPrepago_GetViajeDetalleRemain"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona", adInteger, adParamInput, , mIDPersona)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta", adVarChar, adParamInput, 20, mIDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("Fecha", adDate, adParamInput, , mFechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDListaPrecio", adInteger, adParamInput, , mIDListaPrecio)
    cmdData.Parameters.Append cmdData.CreateParameter("ReservasRestantes", adSmallInt, adParamOutput)
    cmdData.Execute
    
    ObtenerPrepagosRestantes = cmdData.Parameters("ReservasRestantes").Value
    Set cmdData = Nothing
    Exit Function
    
ErrorHandler:
    CSM_Error.ShowErrorMessage "Classes.ViajeDetalle.VerificarPrepagoOtrasReservas", "Error al verificar si se aplica un Prepago a otras reservas." & vbCr & vbCr & "IDPersona: " & mIDPersona & vbCr & "IDRuta: " & mIDRuta & vbCr & "Fecha/Hora: " & mFechaHora
End Function
