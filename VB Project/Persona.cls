VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Persona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDPersona As Long
Private mApellido As String
Private mNombre As String
Private mDomicilioCalle1 As String
Private mDomicilioNumero As String
Private mDomicilioPiso As String
Private mDomicilioDepartamento As String
Private mDomicilioCalle2 As String
Private mDomicilioCalle3 As String
Private mCodigoPostal As String
Private mIDProvincia As String
Private mIDLocalidad As Long
Private mIDDocumentoTipo As Byte
Private mDocumentoNumero As String
Private mIDCondicionIVA As Byte
Private mFechaNacimiento As Date
Private mEmail As String
Private mIDTelefono1Tipo As Byte
Private mTelefono1TipoOtro As String
Private mTelefono1Area As String
Private mTelefono1Numero As String
Private mIDTelefono2Tipo As Byte
Private mTelefono2TipoOtro As String
Private mTelefono2Area As String
Private mTelefono2Numero As String
Private mIDTelefono3Tipo As Byte
Private mTelefono3TipoOtro As String
Private mTelefono3Area As String
Private mTelefono3Numero As String
Private mIDTelefono4Tipo As Byte
Private mTelefono4TipoOtro As String
Private mTelefono4Area As String
Private mTelefono4Numero As String
Private mIDTelefono5Tipo As Byte
Private mTelefono5TipoOtro As String
Private mTelefono5Area As String
Private mTelefono5Numero As String
Private mEntidadTipo As String
Private mIDPersonaACargo As Long
Private mIDPersonaCuentaCorriente As Long
Private mSueldoImporte As Currency
Private mSueldoDia As Integer
Private mPermiteViajarSinPagar As Boolean
Private mHabilitadoViajar As Boolean
Private mHabilitadoInternet As Boolean
Private mPassword As String
Private mNotas As String
Private mActivo As Boolean
Private mListaPasajero As Boolean
Private mFechaHoraCreacion As Date
Private mIDUsuarioCreacion As Integer
Private mFechaHoraModificacion As Date
Private mIDUsuarioModificacion As Integer

'SALDO ACTUAL
Public mViajeActual_FechaHora As Date
Public mViajeActual_IDRuta As String
Public mViajeActual_Indice As Long
Public mSaldoActual As Currency

'STATE VARIABLES
Private mIsNew As Boolean
Private mIsCopy As Boolean
Private mIsDirty As Boolean
Private mNoMatch As Boolean

'BEHAVIOR VARIABLES
Private mNoMatchRaiseError As Boolean
Private mRefreshListSkip As Boolean

'DATABASE VARIABLES
Private mDatabase As CSC_Database_ADO_SQL

'INTERNAL VARIABLES
Private mrecData As ADODB.Recordset

Private mCuentaCorrienteCaja As CuentaCorrienteCaja
Private mUsuario As Usuario

Private mIDDocumentoTipoOriginal As Byte
Private mDocumentoNumeroOriginal As String

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

Public Property Get Apellido() As String
    Apellido = mApellido
End Property

Public Property Let Apellido(ByVal Value As String)
    If Value <> mApellido Then
        mIsDirty = True
    End If
    mApellido = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Nombre() As String
    Nombre = mNombre
End Property

Public Property Let Nombre(ByVal Value As String)
    If Value <> mNombre Then
        mIsDirty = True
    End If
    mNombre = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DomicilioCalle1() As String
    DomicilioCalle1 = mDomicilioCalle1
End Property

Public Property Let DomicilioCalle1(ByVal Value As String)
    If Value <> mDomicilioCalle1 Then
        mIsDirty = True
    End If
    mDomicilioCalle1 = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DomicilioNumero() As String
    DomicilioNumero = mDomicilioNumero
End Property

Public Property Let DomicilioNumero(ByVal Value As String)
    If Value <> mDomicilioNumero Then
        mIsDirty = True
    End If
    mDomicilioNumero = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DomicilioPiso() As String
    DomicilioPiso = mDomicilioPiso
End Property

Public Property Let DomicilioPiso(ByVal Value As String)
    If Value <> mDomicilioPiso Then
        mIsDirty = True
    End If
    mDomicilioPiso = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DomicilioDepartamento() As String
    DomicilioDepartamento = mDomicilioDepartamento
End Property

Public Property Let DomicilioDepartamento(ByVal Value As String)
    If Value <> mDomicilioDepartamento Then
        mIsDirty = True
    End If
    mDomicilioDepartamento = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DomicilioCalle2() As String
    DomicilioCalle2 = mDomicilioCalle2
End Property

Public Property Let DomicilioCalle2(ByVal Value As String)
    If Value <> mDomicilioCalle2 Then
        mIsDirty = True
    End If
    mDomicilioCalle2 = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DomicilioCalle3() As String
    DomicilioCalle3 = mDomicilioCalle3
End Property

Public Property Let DomicilioCalle3(ByVal Value As String)
    If Value <> mDomicilioCalle3 Then
        mIsDirty = True
    End If
    mDomicilioCalle3 = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get CodigoPostal() As String
    CodigoPostal = mCodigoPostal
End Property

Public Property Let CodigoPostal(ByVal Value As String)
    If Value <> mCodigoPostal Then
        mIsDirty = True
    End If
    mCodigoPostal = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDProvincia() As String
    IDProvincia = mIDProvincia
End Property

Public Property Let IDProvincia(ByVal Value As String)
    If Value <> mIDProvincia Then
        mIsDirty = True
    End If
    mIDProvincia = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDLocalidad() As Long
    IDLocalidad = mIDLocalidad
End Property

Public Property Let IDLocalidad(ByVal Value As Long)
    If Value <> mIDLocalidad Then
        mIsDirty = True
    End If
    mIDLocalidad = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDDocumentoTipo() As Byte
    IDDocumentoTipo = mIDDocumentoTipo
End Property

Public Property Let IDDocumentoTipo(ByVal Value As Byte)
    If Value <> mIDDocumentoTipo Then
        mIsDirty = True
    End If
    mIDDocumentoTipo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DocumentoTipo() As DocumentoTipo
    Static oDocumentoTipo As DocumentoTipo
    
    If oDocumentoTipo Is Nothing Then
        Set oDocumentoTipo = New DocumentoTipo
    End If
    If oDocumentoTipo.IDDocumentoTipo <> mIDDocumentoTipo Then
        If mIDDocumentoTipo = 0 Then
            Set oDocumentoTipo = New DocumentoTipo
        Else
            oDocumentoTipo.IDDocumentoTipo = mIDDocumentoTipo
            Call oDocumentoTipo.Load
        End If
    End If
    Set DocumentoTipo = oDocumentoTipo
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DocumentoNumero() As String
    DocumentoNumero = mDocumentoNumero
End Property

Public Property Let DocumentoNumero(ByVal Value As String)
    If Value <> mDocumentoNumero Then
        mIsDirty = True
    End If
    mDocumentoNumero = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDCondicionIVA() As Byte
    IDCondicionIVA = mIDCondicionIVA
End Property

Public Property Let IDCondicionIVA(ByVal Value As Byte)
    If Value <> mIDCondicionIVA Then
        mIsDirty = True
    End If
    mIDCondicionIVA = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get CondicionIVA() As CondicionIVA
    Static oCondicionIVA As CondicionIVA
    
    If mIDCondicionIVA = "" Then
        Set CondicionIVA = Nothing
    Else
        If oCondicionIVA Is Nothing Then
            Set oCondicionIVA = New CondicionIVA
        End If
        If oCondicionIVA.IDCondicionIVA <> mIDCondicionIVA Then
            oCondicionIVA.IDCondicionIVA = mIDCondicionIVA
            Call oCondicionIVA.Load
        End If
        Set CondicionIVA = oCondicionIVA
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = mFechaNacimiento
End Property

Public Property Let FechaNacimiento(ByVal Value As Date)
    If Value <> mFechaNacimiento Then
        mIsDirty = True
    End If
    mFechaNacimiento = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Email() As String
    Email = mEmail
End Property

Public Property Let Email(ByVal Value As String)
    If Value <> mEmail Then
        mIsDirty = True
    End If
    mEmail = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDTelefono1Tipo() As Byte
    IDTelefono1Tipo = mIDTelefono1Tipo
End Property

Public Property Let IDTelefono1Tipo(ByVal Value As Byte)
    If Value <> mIDTelefono1Tipo Then
        mIsDirty = True
    End If
    mIDTelefono1Tipo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono1TipoOtro() As String
    Telefono1TipoOtro = mTelefono1TipoOtro
End Property

Public Property Let Telefono1TipoOtro(ByVal Value As String)
    If Value <> mTelefono1TipoOtro Then
        mIsDirty = True
    End If
    mTelefono1TipoOtro = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono1Area() As String
    Telefono1Area = mTelefono1Area
End Property

Public Property Let Telefono1Area(ByVal Value As String)
    If Value <> mTelefono1Area Then
        mIsDirty = True
    End If
    mTelefono1Area = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono1Numero() As String
    Telefono1Numero = mTelefono1Numero
End Property

Public Property Let Telefono1Numero(ByVal Value As String)
    If Value <> mTelefono1Numero Then
        mIsDirty = True
    End If
    mTelefono1Numero = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDTelefono2Tipo() As Byte
    IDTelefono2Tipo = mIDTelefono2Tipo
End Property

Public Property Let IDTelefono2Tipo(ByVal Value As Byte)
    If Value <> mIDTelefono2Tipo Then
        mIsDirty = True
    End If
    mIDTelefono2Tipo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono2TipoOtro() As String
    Telefono2TipoOtro = mTelefono2TipoOtro
End Property

Public Property Let Telefono2TipoOtro(ByVal Value As String)
    If Value <> mTelefono2TipoOtro Then
        mIsDirty = True
    End If
    mTelefono2TipoOtro = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono2Area() As String
    Telefono2Area = mTelefono2Area
End Property

Public Property Let Telefono2Area(ByVal Value As String)
    If Value <> mTelefono2Area Then
        mIsDirty = True
    End If
    mTelefono2Area = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono2Numero() As String
    Telefono2Numero = mTelefono2Numero
End Property

Public Property Let Telefono2Numero(ByVal Value As String)
    If Value <> mTelefono2Numero Then
        mIsDirty = True
    End If
    mTelefono2Numero = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDTelefono3Tipo() As Byte
    IDTelefono3Tipo = mIDTelefono3Tipo
End Property

Public Property Let IDTelefono3Tipo(ByVal Value As Byte)
    If Value <> mIDTelefono3Tipo Then
        mIsDirty = True
    End If
    mIDTelefono3Tipo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono3TipoOtro() As String
    Telefono3TipoOtro = mTelefono3TipoOtro
End Property

Public Property Let Telefono3TipoOtro(ByVal Value As String)
    If Value <> mTelefono3TipoOtro Then
        mIsDirty = True
    End If
    mTelefono3TipoOtro = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono3Area() As String
    Telefono3Area = mTelefono3Area
End Property

Public Property Let Telefono3Area(ByVal Value As String)
    If Value <> mTelefono3Area Then
        mIsDirty = True
    End If
    mTelefono3Area = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono3Numero() As String
    Telefono3Numero = mTelefono3Numero
End Property

Public Property Let Telefono3Numero(ByVal Value As String)
    If Value <> mTelefono3Numero Then
        mIsDirty = True
    End If
    mTelefono3Numero = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDTelefono4Tipo() As Byte
    IDTelefono4Tipo = mIDTelefono4Tipo
End Property

Public Property Let IDTelefono4Tipo(ByVal Value As Byte)
    If Value <> mIDTelefono4Tipo Then
        mIsDirty = True
    End If
    mIDTelefono4Tipo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono4TipoOtro() As String
    Telefono4TipoOtro = mTelefono4TipoOtro
End Property

Public Property Let Telefono4TipoOtro(ByVal Value As String)
    If Value <> mTelefono4TipoOtro Then
        mIsDirty = True
    End If
    mTelefono4TipoOtro = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono4Area() As String
    Telefono4Area = mTelefono4Area
End Property

Public Property Let Telefono4Area(ByVal Value As String)
    If Value <> mTelefono4Area Then
        mIsDirty = True
    End If
    mTelefono4Area = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono4Numero() As String
    Telefono4Numero = mTelefono4Numero
End Property

Public Property Let Telefono4Numero(ByVal Value As String)
    If Value <> mTelefono4Numero Then
        mIsDirty = True
    End If
    mTelefono4Numero = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDTelefono5Tipo() As Byte
    IDTelefono5Tipo = mIDTelefono5Tipo
End Property

Public Property Let IDTelefono5Tipo(ByVal Value As Byte)
    If Value <> mIDTelefono5Tipo Then
        mIsDirty = True
    End If
    mIDTelefono5Tipo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono5TipoOtro() As String
    Telefono5TipoOtro = mTelefono5TipoOtro
End Property

Public Property Let Telefono5TipoOtro(ByVal Value As String)
    If Value <> mTelefono5TipoOtro Then
        mIsDirty = True
    End If
    mTelefono5TipoOtro = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono5Area() As String
    Telefono5Area = mTelefono5Area
End Property

Public Property Let Telefono5Area(ByVal Value As String)
    If Value <> mTelefono5Area Then
        mIsDirty = True
    End If
    mTelefono5Area = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Telefono5Numero() As String
    Telefono5Numero = mTelefono5Numero
End Property

Public Property Let Telefono5Numero(ByVal Value As String)
    If Value <> mTelefono5Numero Then
        mIsDirty = True
    End If
    mTelefono5Numero = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get EntidadTipo() As String
    EntidadTipo = mEntidadTipo
End Property

Public Property Let EntidadTipo(ByVal Value As String)
    If Value <> mEntidadTipo Then
        mIsDirty = True
    End If
    mEntidadTipo = Value
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

Public Property Get IDPersonaACargo() As Long
    IDPersonaACargo = mIDPersonaACargo
End Property

Public Property Let IDPersonaACargo(ByVal Value As Long)
    If Value <> mIDPersonaACargo Then
        mIsDirty = True
    End If
    mIDPersonaACargo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get PersonaCuentaCorriente() As Persona
    If IDPersonaCuentaCorriente = 0 Then
        Set PersonaCuentaCorriente = Nothing
    Else
        Set PersonaCuentaCorriente = New Persona
        PersonaCuentaCorriente.IDPersona = mIDPersonaCuentaCorriente
        If PersonaCuentaCorriente.Load() Then
        End If
    End If
End Property


'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get SueldoImporte() As Currency
    SueldoImporte = mSueldoImporte
End Property

Public Property Let SueldoImporte(ByVal Value As Currency)
    If Value <> mSueldoImporte Then
        mIsDirty = True
    End If
    mSueldoImporte = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get SueldoDia() As Integer
    SueldoDia = mSueldoDia
End Property

Public Property Let SueldoDia(ByVal Value As Integer)
    If Value <> mSueldoDia Then
        mIsDirty = True
    End If
    mSueldoDia = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get PermiteViajarSinPagar() As Boolean
    PermiteViajarSinPagar = mPermiteViajarSinPagar
End Property

Public Property Let PermiteViajarSinPagar(ByVal Value As Boolean)
    If Value <> mPermiteViajarSinPagar Then
        mIsDirty = True
    End If
    mPermiteViajarSinPagar = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get HabilitadoViajar() As Boolean
    HabilitadoViajar = mHabilitadoViajar
End Property

Public Property Let HabilitadoViajar(ByVal Value As Boolean)
    If Value <> mHabilitadoViajar Then
        mIsDirty = True
    End If
    mHabilitadoViajar = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get HabilitadoInternet() As Boolean
    HabilitadoInternet = mHabilitadoInternet
End Property

Public Property Let HabilitadoInternet(ByVal Value As Boolean)
    If Value <> mHabilitadoInternet Then
        mIsDirty = True
    End If
    mHabilitadoInternet = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Password() As String
    Password = mPassword
End Property

Public Property Let Password(ByVal Value As String)
    If Value <> mPassword Then
        mIsDirty = True
    End If
    mPassword = Value
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

Public Property Get Activo() As Boolean
    Activo = mActivo
End Property

Public Property Let Activo(ByVal Value As Boolean)
    If Value <> mActivo Then
        mIsDirty = True
    End If
    mActivo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ListaPasajero() As Boolean
    ListaPasajero = mListaPasajero
End Property

Public Property Let ListaPasajero(ByVal Value As Boolean)
    If Value <> mListaPasajero Then
        mIsDirty = True
    End If
    mListaPasajero = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaHoraCreacion() As Date
    FechaHoraCreacion = mFechaHoraCreacion
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

Public Property Get ViajeActual_FechaHora() As Date
    ViajeActual_FechaHora = mViajeActual_FechaHora
End Property

Public Property Let ViajeActual_FechaHora(ByVal Value As Date)
    mViajeActual_FechaHora = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ViajeActual_IDRuta() As String
    ViajeActual_IDRuta = mViajeActual_IDRuta
End Property

Public Property Let ViajeActual_IDRuta(ByVal Value As String)
    mViajeActual_IDRuta = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ViajeActual_Indice() As Long
    ViajeActual_Indice = mViajeActual_Indice
End Property

Public Property Let ViajeActual_Indice(ByVal Value As Long)
    mViajeActual_Indice = Value
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
    RefreshList_Module.RefreshList_RefreshPersona IDPersona
End Sub

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Database() As CSC_Database_ADO_SQL
    Set Database = mDatabase
End Property

Public Property Set Database(ByRef Value As CSC_Database_ADO_SQL)
    Set mDatabase = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get SaldoActual() As Currency
    SaldoActual = mSaldoActual
End Property

Public Property Get SaldoActual_Formatted() As String
    SaldoActual_Formatted = Format(mSaldoActual, "Currency")
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private Function OpenRecordset() As Boolean
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Persona_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , IDPersona)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    If pParametro.Database_CursorLocation = adUseClient Then
        mrecData.Properties("Update Criteria").Value = adCriteriaKey
    End If
    Set cmdData = Nothing
    
    Set mCuentaCorrienteCaja = New CuentaCorrienteCaja
    Set mUsuario = New Usuario
    
    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Persona.OpenRecordset", "Error al abrir la tabla de Personas."
End Function

Public Function Load() As Boolean
    Dim DES As CSC_Encryption_DES
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If Not OpenRecordset() Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If (Not NoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
        
        InitializeValues
    Else
        mNoMatch = False
        
        mApellido = mrecData("Apellido").Value
        mNombre = mrecData("Nombre").Value & ""
        mDomicilioCalle1 = mrecData("DomicilioCalle1").Value & ""
        mDomicilioNumero = mrecData("DomicilioNumero").Value & ""
        mDomicilioPiso = mrecData("DomicilioPiso").Value & ""
        mDomicilioDepartamento = mrecData("DomicilioDepartamento").Value & ""
        mDomicilioCalle2 = mrecData("DomicilioCalle2").Value & ""
        mDomicilioCalle3 = mrecData("DomicilioCalle3").Value & ""
        mCodigoPostal = mrecData("CodigoPostal").Value & ""
        mIDProvincia = mrecData("IDProvincia").Value & ""
        mIDLocalidad = Val(mrecData("IDLocalidad").Value & "")
        mIDDocumentoTipo = Val(mrecData("IDDocumentoTipo").Value & "")
        mDocumentoNumero = mrecData("DocumentoNumero").Value & ""
        mIDCondicionIVA = Val(mrecData("IDCondicionIVA").Value & "")
        mFechaNacimiento = IIf(IsNull(mrecData("FechaNacimiento").Value), DATE_TIME_FIELD_NULL_VALUE, mrecData("FechaNacimiento").Value)
        mEmail = mrecData("Email").Value & ""
        mIDTelefono1Tipo = Val(mrecData("IDTelefono1Tipo").Value & "")
        mTelefono1TipoOtro = mrecData("Telefono1TipoOtro").Value & ""
        mTelefono1Area = mrecData("Telefono1Area").Value & ""
        mTelefono1Numero = mrecData("Telefono1Numero").Value & ""
        mIDTelefono2Tipo = Val(mrecData("IDTelefono2Tipo").Value & "")
        mTelefono2TipoOtro = mrecData("Telefono2TipoOtro").Value & ""
        mTelefono2Area = mrecData("Telefono2Area").Value & ""
        mTelefono2Numero = mrecData("Telefono2Numero").Value & ""
        mIDTelefono3Tipo = Val(mrecData("IDTelefono3Tipo").Value & "")
        mTelefono3TipoOtro = mrecData("Telefono3TipoOtro").Value & ""
        mTelefono3Area = mrecData("Telefono3Area").Value & ""
        mTelefono3Numero = mrecData("Telefono3Numero").Value & ""
        mIDTelefono4Tipo = Val(mrecData("IDTelefono4Tipo").Value & "")
        mTelefono4TipoOtro = mrecData("Telefono4TipoOtro").Value & ""
        mTelefono4Area = mrecData("Telefono4Area").Value & ""
        mTelefono4Numero = mrecData("Telefono4Numero").Value & ""
        mIDTelefono5Tipo = Val(mrecData("IDTelefono5Tipo").Value & "")
        mTelefono5TipoOtro = mrecData("Telefono5TipoOtro").Value & ""
        mTelefono5Area = mrecData("Telefono5Area").Value & ""
        mTelefono5Numero = mrecData("Telefono5Numero").Value & ""
        mEntidadTipo = mrecData("EntidadTipo").Value
        mIDPersonaACargo = Val(mrecData("IDPersonaACargo").Value & "")
        mIDPersonaCuentaCorriente = Val(mrecData("IDPersonaCuentaCorriente").Value & "")
        mSueldoImporte = IIf(EntidadTipo <> ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO, 0, IIf(IsNull(mrecData("SueldoImporte").Value), 0, mrecData("SueldoImporte").Value))
        mSueldoDia = IIf(EntidadTipo <> ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO, 0, IIf(IsNull(mrecData("SueldoDia").Value), 0, mrecData("SueldoDia").Value))
        mPermiteViajarSinPagar = IIf(EntidadTipo <> ENTIDAD_TIPO_PERSONA_CLIENTE, False, mrecData("PermiteViajarSinPagar").Value)
        mHabilitadoViajar = mrecData("HabilitadoViajar").Value
        mHabilitadoInternet = mrecData("HabilitadoInternet").Value
        
        If IsNull(mrecData("Password").Value) Or Trim(DocumentoNumero) = "" Then
            mPassword = ""
        Else
            Set DES = New CSC_Encryption_DES
            mPassword = DES.DecryptString(mrecData("Password").Value, PASSWORD_ENCRYPTION_KEY)
            Set DES = Nothing
        End If
        
        mNotas = mrecData("Notas").Value & ""
        mActivo = mrecData("Activo").Value
        mListaPasajero = mrecData("ListaPasajero").Value
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
        
        mIDDocumentoTipoOriginal = IDDocumentoTipo
        mDocumentoNumeroOriginal = DocumentoNumero

        If EntidadTipo <> ENTIDAD_TIPO_PERSONA_CLIENTE Then
            mCuentaCorrienteCaja.IDPersona = IDPersona
            mCuentaCorrienteCaja.NoMatchRaiseError = False
            If Not mCuentaCorrienteCaja.LoadByPersona() Then
                Exit Function
            End If
        End If
    
        If mEntidadTipo = ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO Then
            mUsuario.IDPersona = IDPersona
            mUsuario.NoMatchRaiseError = False
            If Not mUsuario.LoadByPersona() Then
                Exit Function
            End If
        End If
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Persona.Load", "Error al obtener los datos de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIDPersona = 0

        mIsNew = True
        mIsCopy = True
        mIsDirty = True
    End If
End Function

Public Function AddNew() As Boolean
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    Dim CharIndex As Integer
    Dim LastSeparator As Boolean
    Dim DES As CSC_Encryption_DES

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If mrecData Is Nothing Then
        If Not OpenRecordset() Then
            Exit Function
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    'VERIFICO EL DOCUMENTO
    If IDDocumentoTipo <> 0 And DocumentoNumero <> "" Then
        Set cmdData = New ADODB.command
        Set cmdData.ActiveConnection = pDatabase.Connection
        cmdData.CommandText = "usp_Persona_ByDocumento"
        cmdData.CommandType = adCmdStoredProc
        cmdData.Parameters.Append cmdData.CreateParameter("IDDocumentoTipo_FILTER", adChar, adParamInput, 10, IDDocumentoTipo)
        cmdData.Parameters.Append cmdData.CreateParameter("DocumentoNumero_FILTER", adVarChar, adParamInput, 15, DocumentoNumero)
        Set recData = New ADODB.Recordset
        recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
        Set cmdData = Nothing
        If Not recData.EOF Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe una Persona con el mismo Tipo y Número de Documento." & vbCr & vbCr & (recData("Apellido").Value & ", " + recData("Nombre").Value), vbExclamation, App.Title
            recData.Close
            Set recData = Nothing
            Exit Function
        End If
        recData.Close
        Set recData = Nothing
    End If
    
    'MAX ID
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Persona_IDMax"
    cmdData.CommandType = adCmdStoredProc
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    IDPersona = Val(recData("IDPersonaMax").Value & "") + 1
    recData.Close
    Set recData = Nothing
    
    mrecData.AddNew
    mrecData("IDPersona").Value = IDPersona
    mrecData("Apellido").Value = Apellido
    mrecData("Nombre").Value = IIf(Nombre = "", Null, Nombre)
    mrecData("DomicilioCalle1").Value = IIf(Trim(DomicilioCalle1) = "", Null, DomicilioCalle1)
    mrecData("DomicilioNumero").Value = IIf(Trim(DomicilioNumero) = "", Null, DomicilioNumero)
    mrecData("DomicilioPiso").Value = IIf(Trim(DomicilioPiso) = "", Null, DomicilioPiso)
    mrecData("DomicilioDepartamento").Value = IIf(Trim(DomicilioDepartamento) = "", Null, DomicilioDepartamento)
    mrecData("DomicilioCalle2").Value = IIf(Trim(DomicilioCalle2) = "", Null, DomicilioCalle2)
    mrecData("DomicilioCalle3").Value = IIf(Trim(DomicilioCalle3) = "", Null, DomicilioCalle3)
    mrecData("CodigoPostal").Value = IIf(Trim(CodigoPostal) = "", Null, CodigoPostal)
    mrecData("IDProvincia").Value = IIf(Trim(IDProvincia) = "", Null, IDProvincia)
    mrecData("IDLocalidad").Value = IIf(IDLocalidad = 0, Null, IDLocalidad)
    mrecData("IDDocumentoTipo").Value = IIf(IDDocumentoTipo = 0, Null, IDDocumentoTipo)
    mrecData("DocumentoNumero").Value = IIf(Trim(DocumentoNumero) = "", Null, DocumentoNumero)
    mrecData("IDCondicionIVA").Value = IIf(IDCondicionIVA = 0, Null, IDCondicionIVA)
    mrecData("FechaNacimiento").Value = IIf(FechaNacimiento = DATE_TIME_FIELD_NULL_VALUE, Null, FechaNacimiento)
    mrecData("Email").Value = IIf(Trim(Email) = "", Null, Trim(Email))
    mrecData("IDTelefono1Tipo").Value = IIf(IDTelefono1Tipo = 0, Null, IDTelefono1Tipo)
    mrecData("Telefono1TipoOtro").Value = IIf(IDTelefono1Tipo <> pParametro.TelefonoTipo_ID_Otro Or Trim(Telefono1TipoOtro) = "", Null, Telefono1TipoOtro)
    mrecData("Telefono1Area").Value = IIf(Trim(Telefono1Area) = "", Null, Trim(Telefono1Area))
    mrecData("Telefono1Numero").Value = IIf(Trim(Telefono1Numero) = "", Null, Trim(Telefono1Numero))
    mrecData("IDTelefono2Tipo").Value = IIf(IDTelefono2Tipo = 0, Null, IDTelefono2Tipo)
    mrecData("Telefono2TipoOtro").Value = IIf(IDTelefono2Tipo <> pParametro.TelefonoTipo_ID_Otro Or Trim(Telefono2TipoOtro) = "", Null, Telefono2TipoOtro)
    mrecData("Telefono2Area").Value = IIf(Trim(Telefono2Area) = "", Null, Trim(Telefono2Area))
    mrecData("Telefono2Numero").Value = IIf(Trim(Telefono2Numero) = "", Null, Trim(Telefono2Numero))
    mrecData("IDTelefono3Tipo").Value = IIf(IDTelefono3Tipo = 0, Null, IDTelefono3Tipo)
    mrecData("Telefono3TipoOtro").Value = IIf(IDTelefono3Tipo <> pParametro.TelefonoTipo_ID_Otro Or Trim(Telefono3TipoOtro) = "", Null, Telefono3TipoOtro)
    mrecData("Telefono3Area").Value = IIf(Trim(Telefono3Area) = "", Null, Trim(Telefono3Area))
    mrecData("Telefono3Numero").Value = IIf(Trim(Telefono3Numero) = "", Null, Trim(Telefono3Numero))
    mrecData("IDTelefono4Tipo").Value = IIf(IDTelefono4Tipo = 0, Null, IDTelefono4Tipo)
    mrecData("Telefono4TipoOtro").Value = IIf(IDTelefono4Tipo <> pParametro.TelefonoTipo_ID_Otro Or Trim(Telefono4TipoOtro) = "", Null, Telefono4TipoOtro)
    mrecData("Telefono4Area").Value = IIf(Trim(Telefono4Area) = "", Null, Trim(Telefono4Area))
    mrecData("Telefono4Numero").Value = IIf(Trim(Telefono4Numero) = "", Null, Trim(Telefono4Numero))
    mrecData("IDTelefono5Tipo").Value = IIf(IDTelefono5Tipo = 0, Null, IDTelefono5Tipo)
    mrecData("Telefono5TipoOtro").Value = IIf(IDTelefono5Tipo <> pParametro.TelefonoTipo_ID_Otro Or Trim(Telefono5TipoOtro) = "", Null, Telefono5TipoOtro)
    mrecData("Telefono5Area").Value = IIf(Trim(Telefono5Area) = "", Null, Trim(Telefono5Area))
    mrecData("Telefono5Numero").Value = IIf(Trim(Telefono5Numero) = "", Null, Trim(Telefono5Numero))
    mrecData("EntidadTipo").Value = EntidadTipo
    mrecData("IDPersonaACargo").Value = IIf(IDPersonaACargo = 0, Null, IDPersonaACargo)
    mrecData("IDPersonaCuentaCorriente").Value = IIf(IDPersonaCuentaCorriente = 0, Null, IDPersonaCuentaCorriente)
    mrecData("SueldoImporte").Value = IIf(EntidadTipo <> ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO, Null, IIf(SueldoImporte = 0, Null, SueldoImporte))
    mrecData("SueldoDia").Value = IIf(EntidadTipo <> ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO, Null, IIf(SueldoDia = 0, Null, SueldoDia))
    mrecData("PermiteViajarSinPagar").Value = IIf(EntidadTipo <> ENTIDAD_TIPO_PERSONA_CLIENTE, False, PermiteViajarSinPagar)
    mrecData("Notas").Value = IIf(Trim(Notas) = "", Null, Notas)
    mrecData("HabilitadoViajar").Value = mHabilitadoViajar
    mrecData("HabilitadoInternet").Value = HabilitadoInternet
    
    If Trim(Password) = "" Then
        mrecData("Password").Value = Null
    Else
        Set DES = New CSC_Encryption_DES
        mrecData("Password").Value = DES.EncryptString(Password, PASSWORD_ENCRYPTION_KEY, False)
        Set DES = Nothing
    End If
    
    mrecData("Activo").Value = Activo
    mrecData("ListaPasajero").Value = ListaPasajero
    mrecData("FechaHoraCreacion").Value = Now
    mrecData("IDUsuarioCreacion").Value = pUsuario.IDUsuario
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If EntidadTipo <> ENTIDAD_TIPO_PERSONA_CLIENTE Then
        mCuentaCorrienteCaja.Nombre = ApellidoNombre
        mCuentaCorrienteCaja.IDPersona = IDPersona
        mCuentaCorrienteCaja.Activo = Activo
        mCuentaCorrienteCaja.Update
    End If
        
    If EntidadTipo = ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO Then
        mUsuario.LoginName = LCase(Left(Nombre, 1))
        For CharIndex = 2 To Len(Nombre)
            If Mid(Nombre, CharIndex, 1) = " " Or Mid(Nombre, CharIndex, 1) = "." Then
                LastSeparator = True
            Else
                If LastSeparator Then
                    LastSeparator = False
                    mUsuario.LoginName = mUsuario.LoginName & LCase(Mid(Nombre, CharIndex, 1))
                End If
            End If
        Next CharIndex
        mUsuario.LoginName = mUsuario.LoginName & LCase(Replace(Apellido, " ", ""))
        mUsuario.Nombre = ApellidoNombre
        mUsuario.IDPersona = IDPersona
        mUsuario.Descripcion = "Personal Administrativo"
        mUsuario.IDUsuarioGrupo = pParametro.UsuarioGrupo_ID_Predeterminado
        mUsuario.IDEmpresa = pParametro.Empresa_ID_Predeterminada
        mUsuario.Password = mUsuario.LoginName
        mUsuario.Activo = Activo
        mUsuario.Update
        
        MsgBox "Se ha generado un Usuario para esta Persona." & vbCr & vbCr & "Usuario: " & mUsuario.LoginName & vbCr & "Contraseña: " & mUsuario.Password, vbInformation, App.Title
    End If
        
    If Not mRefreshListSkip Then
        RefreshList_RefreshPersona IDPersona
    End If
    
    Screen.MousePointer = vbDefault
    AddNew = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_PRIMARYKEY Or pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe una Persona con el mismo Apellido y Nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Persona.AddNew", "Error al Agregar la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona
        End If
    Else
        ShowErrorMessage "Classes.Persona.AddNew", "Error al Agregar la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona
    End If
    mrecData.CancelUpdate
End Function

Public Function Update() As Boolean
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    Dim CharIndex As Integer
    Dim LastSeparator As Boolean
    Dim DES As CSC_Encryption_DES
    
    If Not mIsDirty Then
        Update = True
        Exit Function
    End If

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    'VERIFICO EL DOCUMENTO
    If IDDocumentoTipo <> 0 And DocumentoNumero <> "" And (IDDocumentoTipo <> mIDDocumentoTipoOriginal Or DocumentoNumero <> mDocumentoNumeroOriginal) Then
        Set cmdData = New ADODB.command
        Set cmdData.ActiveConnection = pDatabase.Connection
        cmdData.CommandText = "usp_Persona_ByDocumento"
        cmdData.CommandType = adCmdStoredProc
        cmdData.Parameters.Append cmdData.CreateParameter("IDDocumentoTipo", adTinyInt, adParamInput, , IDDocumentoTipo)
        cmdData.Parameters.Append cmdData.CreateParameter("DocumentoNumero", adVarChar, adParamInput, 15, DocumentoNumero)
        Set recData = New ADODB.Recordset
        recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
        Set cmdData = Nothing
        If Not recData.EOF Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe una Persona con el mismo Tipo y Número de Documento." & vbCr & vbCr & (recData("Apellido").Value & ", " + recData("Nombre").Value), vbExclamation, App.Title
            recData.Close
            Set recData = Nothing
            Exit Function
        End If
        recData.Close
        Set recData = Nothing
    End If
    
    mrecData("Apellido").Value = Apellido
    mrecData("Nombre").Value = IIf(Nombre = "", Null, Nombre)
    mrecData("DomicilioCalle1").Value = IIf(Trim(DomicilioCalle1) = "", Null, DomicilioCalle1)
    mrecData("DomicilioNumero").Value = IIf(Trim(DomicilioNumero) = "", Null, DomicilioNumero)
    mrecData("DomicilioPiso").Value = IIf(Trim(DomicilioPiso) = "", Null, DomicilioPiso)
    mrecData("DomicilioDepartamento").Value = IIf(Trim(DomicilioDepartamento) = "", Null, DomicilioDepartamento)
    mrecData("DomicilioCalle2").Value = IIf(Trim(DomicilioCalle2) = "", Null, DomicilioCalle2)
    mrecData("DomicilioCalle3").Value = IIf(Trim(DomicilioCalle3) = "", Null, DomicilioCalle3)
    mrecData("CodigoPostal").Value = IIf(Trim(CodigoPostal) = "", Null, CodigoPostal)
    mrecData("IDProvincia").Value = IIf(Trim(IDProvincia) = "", Null, IDProvincia)
    mrecData("IDLocalidad").Value = IIf(IDLocalidad = 0, Null, IDLocalidad)
    mrecData("IDDocumentoTipo").Value = IIf(IDDocumentoTipo = 0, Null, IDDocumentoTipo)
    mrecData("DocumentoNumero").Value = IIf(Trim(DocumentoNumero) = "", Null, DocumentoNumero)
    mrecData("IDCondicionIVA").Value = IIf(IDCondicionIVA = 0, Null, IDCondicionIVA)
    mrecData("FechaNacimiento").Value = IIf(FechaNacimiento = DATE_TIME_FIELD_NULL_VALUE, Null, FechaNacimiento)
    mrecData("Email").Value = IIf(Trim(Email) = "", Null, Trim(Email))
    mrecData("IDTelefono1Tipo").Value = IIf(IDTelefono1Tipo = 0, Null, IDTelefono1Tipo)
    mrecData("Telefono1TipoOtro").Value = IIf(IDTelefono1Tipo = 0 Or IDTelefono1Tipo <> pParametro.TelefonoTipo_ID_Otro Or Trim(Telefono1TipoOtro) = "", Null, Telefono1TipoOtro)
    mrecData("Telefono1Area").Value = IIf(IDTelefono1Tipo = 0 Or Trim(Telefono1Area) = "", Null, Trim(Telefono1Area))
    mrecData("Telefono1Numero").Value = IIf(IDTelefono1Tipo = 0 Or Trim(Telefono1Numero) = "", Null, Trim(Telefono1Numero))
    mrecData("IDTelefono2Tipo").Value = IIf(IDTelefono2Tipo = 0, Null, IDTelefono2Tipo)
    mrecData("Telefono2TipoOtro").Value = IIf(IDTelefono2Tipo = 0 Or IDTelefono2Tipo <> pParametro.TelefonoTipo_ID_Otro Or Trim(Telefono2TipoOtro) = "", Null, Telefono2TipoOtro)
    mrecData("Telefono2Area").Value = IIf(IDTelefono2Tipo = 0 Or Trim(Telefono2Area) = "", Null, Trim(Telefono2Area))
    mrecData("Telefono2Numero").Value = IIf(IDTelefono2Tipo = 0 Or Trim(Telefono2Numero) = "", Null, Trim(Telefono2Numero))
    mrecData("IDTelefono3Tipo").Value = IIf(IDTelefono3Tipo = 0, Null, IDTelefono3Tipo)
    mrecData("Telefono3TipoOtro").Value = IIf(IDTelefono3Tipo = 0 Or IDTelefono3Tipo <> pParametro.TelefonoTipo_ID_Otro Or Trim(Telefono3TipoOtro) = "", Null, Telefono3TipoOtro)
    mrecData("Telefono3Area").Value = IIf(IDTelefono3Tipo = 0 Or Trim(Telefono3Area) = "", Null, Trim(Telefono3Area))
    mrecData("Telefono3Numero").Value = IIf(IDTelefono3Tipo = 0 Or Trim(Telefono3Numero) = "", Null, Trim(Telefono3Numero))
    mrecData("IDTelefono4Tipo").Value = IIf(IDTelefono4Tipo = 0, Null, IDTelefono4Tipo)
    mrecData("Telefono4TipoOtro").Value = IIf(IDTelefono4Tipo = 0 Or IDTelefono4Tipo <> pParametro.TelefonoTipo_ID_Otro Or Trim(Telefono4TipoOtro) = "", Null, Telefono4TipoOtro)
    mrecData("Telefono4Area").Value = IIf(IDTelefono4Tipo = 0 Or Trim(Telefono4Area) = "", Null, Trim(Telefono4Area))
    mrecData("Telefono4Numero").Value = IIf(IDTelefono4Tipo = 0 Or Trim(Telefono4Numero) = "", Null, Trim(Telefono4Numero))
    mrecData("IDTelefono5Tipo").Value = IIf(IDTelefono5Tipo = 0, Null, IDTelefono5Tipo)
    mrecData("Telefono5TipoOtro").Value = IIf(IDTelefono5Tipo = 0 Or IDTelefono5Tipo <> pParametro.TelefonoTipo_ID_Otro Or Trim(Telefono5TipoOtro) = "", Null, Telefono5TipoOtro)
    mrecData("Telefono5Area").Value = IIf(IDTelefono5Tipo = 0 Or Trim(Telefono5Area) = "", Null, Trim(Telefono5Area))
    mrecData("Telefono5Numero").Value = IIf(IDTelefono5Tipo = 0 Or Trim(Telefono5Numero) = "", Null, Trim(Telefono5Numero))
    mrecData("EntidadTipo").Value = EntidadTipo
    mrecData("IDPersonaACargo").Value = IIf(IDPersonaACargo = 0, Null, IDPersonaACargo)
    mrecData("IDPersonaCuentaCorriente").Value = IIf(IDPersonaCuentaCorriente = 0, Null, IDPersonaCuentaCorriente)
    mrecData("SueldoImporte").Value = IIf(EntidadTipo <> ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO, Null, IIf(SueldoImporte = 0, Null, SueldoImporte))
    mrecData("SueldoDia").Value = IIf(EntidadTipo <> ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO, Null, IIf(SueldoDia = 0, Null, SueldoDia))
    mrecData("PermiteViajarSinPagar").Value = IIf(EntidadTipo <> ENTIDAD_TIPO_PERSONA_CLIENTE, False, PermiteViajarSinPagar)
    mrecData("Notas").Value = IIf(Trim(Notas) = "", Null, Notas)
    mrecData("HabilitadoViajar").Value = mHabilitadoViajar
    mrecData("HabilitadoInternet").Value = HabilitadoInternet
    
    If Trim(Password) = "" Then
        mrecData("Password").Value = Null
    Else
        Set DES = New CSC_Encryption_DES
        mrecData("Password").Value = DES.EncryptString(Password, PASSWORD_ENCRYPTION_KEY, False)
        Set DES = Nothing
    End If
    
    mrecData("Activo").Value = Activo
    mrecData("ListaPasajero").Value = ListaPasajero
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If EntidadTipo <> ENTIDAD_TIPO_PERSONA_CLIENTE Then
        mCuentaCorrienteCaja.Nombre = ApellidoNombre
        mCuentaCorrienteCaja.IDPersona = IDPersona
        mCuentaCorrienteCaja.Activo = Activo
        mCuentaCorrienteCaja.Update
    End If
        
    If EntidadTipo = ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO Then
        If mUsuario.NoMatch Then
            mUsuario.LoginName = LCase(Left(Nombre, 1))
            For CharIndex = 2 To Len(Nombre)
                If Mid(Nombre, CharIndex, 1) = " " Or Mid(Nombre, CharIndex, 1) = "." Then
                    LastSeparator = True
                Else
                    If LastSeparator Then
                        LastSeparator = False
                        mUsuario.LoginName = mUsuario.LoginName & LCase(Mid(Nombre, CharIndex, 1))
                    End If
                End If
            Next CharIndex
            mUsuario.LoginName = mUsuario.LoginName & LCase(Replace(Apellido, " ", ""))
            mUsuario.Nombre = ApellidoNombre
            mUsuario.IDPersona = IDPersona
            mUsuario.Descripcion = "Personal Administrativo"
            mUsuario.IDUsuarioGrupo = pParametro.UsuarioGrupo_ID_Predeterminado
            mUsuario.IDEmpresa = pParametro.Empresa_ID_Predeterminada
            mUsuario.Password = mUsuario.LoginName
            mUsuario.Activo = Activo
            mUsuario.Update
            
            MsgBox "Se ha generado un Usuario para esta Persona." & vbCr & vbCr & "Usuario: " & mUsuario.LoginName & vbCr & "Contraseña: " & mUsuario.Password, vbInformation, App.Title
        Else
            mUsuario.Nombre = ApellidoNombre
            mUsuario.Activo = Activo
            mUsuario.Update
        End If
    End If
    
    If Not mRefreshListSkip Then
        RefreshList_RefreshPersona IDPersona
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe una Persona con el mismo Apellido y Nombre." & vbCr & "Si lo buscó y no lo encontró, verifique que no esté inactivo.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Persona.Update", "Error al Actualizar la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona
        End If
    Else
        ShowErrorMessage "Classes.Persona.Update", "Error al Actualizar la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona
    End If
    mrecData.CancelUpdate
End Function

Public Function Delete() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    mrecData.Delete
    
    If Not mRefreshListSkip Then
        RefreshList_RefreshPersona 0
        RefreshList_RefreshCuentaCorrienteCaja 0
        RefreshList_RefreshUsuario ""
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede Eliminar la Persona debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Persona.Delete", "Error al Eliminar la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona
        End If
    Else
        ShowErrorMessage "Classes.Persona.Delete", "Error al Eliminar la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona
    End If
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Private Sub InitializeValues()
    mApellido = ""
    mNombre = ""
    mDomicilioCalle1 = ""
    mDomicilioNumero = ""
    mDomicilioPiso = ""
    mDomicilioDepartamento = ""
    mDomicilioCalle2 = ""
    mDomicilioCalle3 = ""
    mCodigoPostal = ""
    mIDProvincia = pParametro.Provincia_ID_Predeterminada
    mIDLocalidad = 0
    mIDDocumentoTipo = 0
    mDocumentoNumero = ""
    mIDCondicionIVA = 0
    mFechaNacimiento = DATE_TIME_FIELD_NULL_VALUE
    mEmail = ""
    mIDTelefono1Tipo = 0
    mTelefono1TipoOtro = ""
    mTelefono1Area = ""
    mTelefono1Numero = ""
    mIDTelefono2Tipo = 0
    mTelefono2TipoOtro = ""
    mTelefono2Area = ""
    mTelefono2Numero = ""
    mIDTelefono3Tipo = 0
    mTelefono3TipoOtro = ""
    mTelefono3Area = ""
    mTelefono3Numero = ""
    mIDTelefono4Tipo = 0
    mTelefono4TipoOtro = ""
    mTelefono4Area = ""
    mTelefono4Numero = ""
    mIDTelefono5Tipo = 0
    mTelefono5TipoOtro = ""
    mTelefono5Area = ""
    mTelefono5Numero = ""
    mEntidadTipo = ""
    mIDPersonaACargo = 0
    mIDPersonaCuentaCorriente = 0
    mSueldoImporte = 0
    mSueldoDia = 0
    mPermiteViajarSinPagar = False
    mHabilitadoViajar = True
    mHabilitadoInternet = False
    mPassword = ""
    mNotas = ""
    mActivo = True
    mListaPasajero = False
    mFechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioCreacion = 0
    mFechaHoraModificacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioModificacion = 0

    mIsNew = True
    mIsCopy = False
    mIsDirty = False
End Sub

Private Sub Class_Initialize()
    InitializeValues
    NoMatchRaiseError = True
    mRefreshListSkip = False
End Sub

Private Sub Class_Terminate()
    If Not mrecData Is Nothing Then
        If mrecData.State = adStateOpen Then
            If Not (mrecData.BOF Or mrecData.EOF) Then
                If mrecData.EditMode <> adEditNone Then
                    mrecData.CancelUpdate
                End If
            End If
            mrecData.Close
        End If
        Set mrecData = Nothing
    End If
    If Not mCuentaCorrienteCaja Is Nothing Then
        Set mCuentaCorrienteCaja = Nothing
    End If
End Sub

Public Property Get ApellidoNombre() As String
    If Apellido = "" Then
        If Nombre <> "" Then
            ApellidoNombre = Nombre
        End If
    Else
        If Nombre = "" Then
            ApellidoNombre = Apellido
        Else
            ApellidoNombre = Apellido & ", " & Nombre
        End If
    End If
End Property

Public Function LoadSaldoActual() As Boolean
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    Select Case mEntidadTipo
        Case ENTIDAD_TIPO_PERSONA_CLIENTE
        Case ENTIDAD_TIPO_PERSONA_CONDUCTOR
            If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CONDUCTOR_SELECT, False) Then
                mSaldoActual = 0
                Exit Function
            End If
        Case ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO
            If Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_ADMINISTRATIVO_SELECT, False) Then
                mSaldoActual = 0
                Exit Function
            End If
    End Select
        
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass
    
    mSaldoActual = 0
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandType = adCmdStoredProc
    If mViajeActual_Indice = 0 Then
        cmdData.CommandText = "sp_Persona_SaldoActual"
        cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , mIDPersona)
    Else
        cmdData.CommandText = "sp_Persona_SaldoActual_ExceptoViaje"
        cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , mIDPersona)
        cmdData.Parameters.Append cmdData.CreateParameter("Viaje_FechaHora_FILTER", adDate, adParamInput, , mViajeActual_FechaHora)
        cmdData.Parameters.Append cmdData.CreateParameter("Viaje_IDRuta_FILTER", adChar, adParamInput, 20, mViajeActual_IDRuta)
        cmdData.Parameters.Append cmdData.CreateParameter("Viaje_Indice_FILTER", adInteger, adParamInput, , mViajeActual_Indice)
    End If
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    
    If Not IsNull(recData("SaldoActual").Value) Then
        mSaldoActual = recData("SaldoActual").Value
    End If
    recData.Close
    Set recData = Nothing
    
    Screen.MousePointer = vbDefault
    LoadSaldoActual = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Persona.LoadSaldoActual", "Error al obtener el Saldo Actual de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona
End Function

Public Function VerificarInasistencias() As Long
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandType = adCmdStoredProc
    cmdData.CommandText = "sp_Persona_Inasistencias"
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , IDPersona)
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_Desde", adDate, adParamInput, , DateAdd("d", -pParametro.Pasajero_AvisoInasistencia_Dias, Date))

    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    
    VerificarInasistencias = recData("Inasistencias").Value

    recData.Close
    Set recData = Nothing
    
    Screen.MousePointer = vbDefault
    
    If VerificarInasistencias >= pParametro.Pasajero_AvisoInasistencia_Veces Then
        MsgBox "Este Pasajero tiene " & VerificarInasistencias & " inasistencias en los últimos " & pParametro.Pasajero_AvisoInasistencia_Dias & " días.", vbExclamation, App.Title
    End If
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Persona.GetInasistencias", "Error al obtener las Inasistencias del Pasajero." & vbCr & vbCr & "IDPersona: " & IDPersona
End Function

Public Function GenerateNewPassword() As String
    Dim aObjetos As Variant
    Dim aColores As Variant
    
    aObjetos = Array("botellones", "vasos", "platos", "cuchillos", "tenedores", "cucharones", "manteles", "jarrones", _
                     "sillones", "armarios", "veladores", "portones", "encendedores", "zapatos", _
                     "autos", "motos", "camiones", "acoplados", "tanques", _
                     "caballos", "toros", "elefantes", "monos", "perros", "gatos", _
                     "duraznos", "monitores", "teclados", "discos", "tambores", "cigarrillos")
                     
    aColores = Array("azules", "verdes", "amarillos", "rojos", "negros", "blancos", "violetas", _
                     "naranjas", "marrones", "celestes", "grises", "rosas", "turquesas")
    
    Randomize Timer
    
    'NUMERO
    GenerateNewPassword = CLng(Rnd(1) * 18) + 2
        
    'OBJETO
    GenerateNewPassword = GenerateNewPassword & aObjetos(CLng(Rnd(1) * UBound(aObjetos)))

    'COLOR
    GenerateNewPassword = GenerateNewPassword & aColores(CLng(Rnd(1) * UBound(aColores)))
End Function
