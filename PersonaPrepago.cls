VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PersonaPrepago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDPersona As Long
Private mIDRutaGrupo_Original As Long
Private mFechaInicio_Original As Date
Private mIDRutaGrupo As Long
Private mFechaInicio As Date
Private mFechaFin As Date
Private mIDListaPrecio As Long
Private mImporteOriginal As Currency
Private mImporte As Currency
Private mIDMedioPago As Byte
Private mCuotas As Byte
Private mOperacion As String
Private mFacturaNumero As String
Private mIDCuentaCorrienteCaja As Long
Private mIDMovimiento_Credito As Long
Private mIDMovimiento_Debito As Long
Private mFechaHoraCreacion As Date
Private mIDUsuarioCreacion As Integer
Private mFechaHoraModificacion As Date
Private mIDUsuarioModificacion As Integer

'STATE VARIABLES
Private mIsNew As Boolean
Private mIsCopy As Boolean
Private mIsDirty As Boolean
Private mNoMatch As Boolean

'BEHAVIOR VARIABLES
Private mNoMatchRaiseError As Boolean
Private mRefreshListSkip As Boolean

'INTERNAL VARIABLES

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

Public Property Get IDRutaGrupo() As Long
    IDRutaGrupo = mIDRutaGrupo
End Property

Public Property Let IDRutaGrupo(ByVal Value As Long)
    If Value <> mIDRutaGrupo Then
        mIsDirty = True
    End If
    mIDRutaGrupo = Value
End Property

Public Property Get RutaGrupo() As RutaGrupo
    Static oRutaGrupo As RutaGrupo
    
    If mIDRutaGrupo = 0 Then
        Set RutaGrupo = Nothing
    Else
        If oRutaGrupo Is Nothing Then
            Set oRutaGrupo = New RutaGrupo
        End If
        If oRutaGrupo.IDRutaGrupo <> mIDRutaGrupo Then
            oRutaGrupo.IDRutaGrupo = mIDRutaGrupo
            Call oRutaGrupo.Load
        End If
        Set RutaGrupo = oRutaGrupo
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaInicio() As Date
    FechaInicio = mFechaInicio
End Property

Public Property Get FechaInicio_Formatted() As String
    FechaInicio_Formatted = Format(mFechaInicio, "Short Date")
End Property

Public Property Let FechaInicio(ByVal Value As Date)
    If Value <> mFechaInicio Then
        mIsDirty = True
    End If
    mFechaInicio = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get FechaFin() As Date
    FechaFin = mFechaFin
End Property

Public Property Get FechaFin_Formatted() As String
    FechaFin_Formatted = Format(mFechaFin, "Short Date")
End Property

Public Property Let FechaFin(ByVal Value As Date)
    If Value <> mFechaFin Then
        mIsDirty = True
    End If
    mFechaFin = Value
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

Public Property Get ImporteOriginal() As Currency
    ImporteOriginal = mImporteOriginal
End Property

Public Property Get ImporteOriginal_Formatted() As String
    ImporteOriginal_Formatted = Format(mImporteOriginal, "Currency")
End Property

Public Property Let ImporteOriginal(ByVal Value As Currency)
    If Value <> mImporteOriginal Then
        mIsDirty = True
    End If
    mImporteOriginal = Value
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

Public Property Get IDMovimiento_Credito() As Long
    IDMovimiento_Credito = mIDMovimiento_Credito
End Property

Public Property Let IDMovimiento_Credito(ByVal Value As Long)
    If Value <> mIDMovimiento_Credito Then
        mIsDirty = True
    End If
    mIDMovimiento_Credito = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDMovimiento_Debito() As Long
    IDMovimiento_Debito = mIDMovimiento_Debito
End Property

Public Property Let IDMovimiento_Debito(ByVal Value As Long)
    If Value <> mIDMovimiento_Debito Then
        mIsDirty = True
    End If
    mIDMovimiento_Debito = Value
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
    RefreshList_Module.RefreshList_RefreshPersonaPrepago IDPersona, IDRutaGrupo, FechaInicio
End Sub

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function Load() As Boolean
    Dim cmdSP As ADODB.command
    Dim recData As ADODB.Recordset

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Call InitializeValues

    Set cmdSP = New ADODB.command
    With cmdSP
        Set .ActiveConnection = pDatabase.Connection
        .CommandText = "usp_PersonaPrepago_Get"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDPersona", adInteger, adParamInput, , mIDPersona)
        .Parameters.Append .CreateParameter("IDRutaGrupo", adInteger, adParamInput, , mIDRutaGrupo)
        .Parameters.Append .CreateParameter("FechaInicio", adDate, adParamInput, , mFechaInicio)
    End With

    Set recData = New ADODB.Recordset
    recData.Open cmdSP, , adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

    Screen.MousePointer = vbHourglass

    Load = GetDataFromRecordset(recData)

    recData.Close
    Set recData = Nothing
    Set cmdSP = Nothing
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.PersonaPrepago.Load", "Error al obtener los datos del Prepago de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "IDRutaGrupo: " & IDRutaGrupo & vbCr & "FechaInicio: " & FechaInicio
End Function

Private Function GetDataFromRecordset(ByRef recData As ADODB.Recordset) As Boolean

    If (Not mNoMatchRaiseError) And recData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False

        mIDPersona = recData("IDPersona").Value
        mIDRutaGrupo_Original = recData("IDRutaGrupo").Value
        mFechaInicio_Original = recData("FechaInicio").Value
        mIDRutaGrupo = recData("IDRutaGrupo").Value
        mFechaInicio = recData("FechaInicio").Value
        mFechaFin = recData("FechaFin").Value
        mIDListaPrecio = recData("IDListaPrecio").Value
        mImporteOriginal = CSM_Function.IfIsNull_Zero(recData("ImporteOriginal").Value)
        mImporte = CSM_Function.IfIsNull_Zero(recData("Importe").Value)
        mIDMedioPago = Val(recData("IDMedioPago").Value & "")
        mCuotas = Val(recData("Cuotas").Value & "")
        mOperacion = recData("Operacion").Value & ""
        mFacturaNumero = recData("FacturaNumero").Value & ""
        mIDCuentaCorrienteCaja = Val(recData("IDCuentaCorrienteCaja").Value & "")
        mIDMovimiento_Credito = Val(recData("IDMovimiento_Credito").Value & "")
        mIDMovimiento_Debito = Val(recData("IDMovimiento_Debito").Value & "")
        mFechaHoraCreacion = recData("FechaHoraCreacion").Value
        mIDUsuarioCreacion = recData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = recData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = recData("IDUsuarioModificacion").Value
    End If

    Screen.MousePointer = vbDefault
    GetDataFromRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.PersonaPrepago.GetDataFromRecordset", "Error al obtener los datos del Prepago de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "IDRutaGrupo: " & IDRutaGrupo & vbCr & "FechaInicio: " & FechaInicio
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIsNew = True
        mIsCopy = True
        mIsDirty = True
    End If
End Function

Public Function Update() As Boolean
    Dim cmdSP As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Set cmdSP = New ADODB.command
    With cmdSP
        Set .ActiveConnection = pDatabase.Connection
        .CommandType = adCmdStoredProc
        If mIsNew Then
            .CommandText = "usp_PersonaPrepago_Add"
        Else
            .CommandText = "usp_PersonaPrepago_Update"
            .Parameters.Append .CreateParameter("IDRutaGrupo_Original", adInteger, adParamInput, , mIDRutaGrupo_Original)
            .Parameters.Append .CreateParameter("FechaInicio_Original", adDate, adParamInput, , mFechaInicio_Original)
        End If

        .Parameters.Append .CreateParameter("IDPersona", adInteger, adParamInput, , mIDPersona)
        .Parameters.Append .CreateParameter("IDRutaGrupo", adInteger, adParamInput, , mIDRutaGrupo)
        .Parameters.Append .CreateParameter("FechaInicio", adDate, adParamInput, , mFechaInicio)
        .Parameters.Append .CreateParameter("FechaFin", adDate, adParamInput, , mFechaFin)
        .Parameters.Append .CreateParameter("IDListaPrecio", adInteger, adParamInput, , mIDListaPrecio)
        .Parameters.Append .CreateParameter("ImporteOriginal", adCurrency, adParamInput, , CSM_Function.IfIsZero_Null(mImporteOriginal))
        .Parameters.Append .CreateParameter("Importe", adCurrency, adParamInput, , mImporte)
        .Parameters.Append .CreateParameter("IDMedioPago", adTinyInt, adParamInput, , IIf(mIDMedioPago = 0, pParametro.MedioPago_Predeterminado_ID, mIDMedioPago))
        .Parameters.Append .CreateParameter("Cuotas", adTinyInt, adParamInput, , CSM_Function.IfIsZero_Null(mCuotas))
        .Parameters.Append .CreateParameter("Operacion", adVarChar, adParamInput, 20, CSM_Function.IfIsZeroLenghtString_Null(mOperacion))
        .Parameters.Append .CreateParameter("FacturaNumero", adVarChar, adParamInput, 20, CSM_Function.IfIsZeroLenghtString_Null(mFacturaNumero))
        .Parameters.Append .CreateParameter("IDCuentaCorrienteCaja", adInteger, adParamInput, , CSM_Function.IfIsZero_Null(mIDCuentaCorrienteCaja))
        .Parameters.Append .CreateParameter("IDMovimiento_Credito", adInteger, adParamInput, , CSM_Function.IfIsZero_Null(mIDMovimiento_Credito))
        .Parameters.Append .CreateParameter("IDMovimiento_Debito", adInteger, adParamInput, , CSM_Function.IfIsZero_Null(mIDMovimiento_Debito))
        .Parameters.Append .CreateParameter("IDUsuario", adSmallInt, adParamInput, , pUsuario.IDUsuario)

        .Execute
    End With

    mIsCopy = False
    mIsDirty = False

    If Not mRefreshListSkip Then
        Call RefreshList
    End If

    Screen.MousePointer = vbDefault
    Update = True
    Exit Function

ErrorHandler:
    If pDatabase.Connection.Errors(0).NativeError = 50000 Then
        MsgBox "Ya existe un Prepago para el mismo Grupo de Rutas y Fecha de inicio", vbInformation, App.Title
    Else
        ShowErrorMessage "Classes.PersonaPrepago.Update", "Error al actualizar el Prepago de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "IDRutaGrupo: " & IDRutaGrupo & vbCr & "FechaInicio: " & FechaInicio
    End If
End Function

Public Function Delete() As Boolean
    Dim cmdSP As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Set cmdSP = New ADODB.command
    With cmdSP
        Set .ActiveConnection = pDatabase.Connection
        .CommandText = "usp_PersonaPrepago_Delete"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDPersona", adInteger, adParamInput, , mIDPersona)
        .Parameters.Append .CreateParameter("IDRutaGrupo", adInteger, adParamInput, , mIDRutaGrupo)
        .Parameters.Append .CreateParameter("FechaInicio", adDate, adParamInput, , mFechaInicio)

        .Execute
    End With

    If Not mRefreshListSkip Then
        Call RefreshList
    End If

    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.PersonaPrepago.Delete", "Error al eliminar el Prepago de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "IDRutaGrupo: " & IDRutaGrupo & vbCr & "FechaInicio: " & FechaInicio
End Function

Private Sub InitializeValues()
    mIDRutaGrupo_Original = 0
    mFechaInicio_Original = DATE_TIME_FIELD_NULL_VALUE
    mFechaFin = DATE_TIME_FIELD_NULL_VALUE
    mIDListaPrecio = 0
    mImporteOriginal = 0
    mImporte = 0
    mIDMedioPago = 0
    mCuotas = 0
    mOperacion = ""
    mFacturaNumero = ""
    mIDCuentaCorrienteCaja = 0
    mIDMovimiento_Credito = 0
    mIDMovimiento_Debito = 0
    mFechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioCreacion = 0
    mFechaHoraModificacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioModificacion = 0

    mIsNew = True
    mIsCopy = False
    mIsDirty = False
End Sub

Private Sub Class_Initialize()
    Call InitializeValues
    mNoMatchRaiseError = True
    mRefreshListSkip = False
End Sub

Private Sub Class_Terminate()
    'TERMINATE CODE
End Sub
