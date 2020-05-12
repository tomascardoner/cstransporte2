VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MedioPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDMedioPago As Byte
Private mAbreviatura As String
Private mNombre As String
Private mUtilizaOperacion As Boolean
Private mIDMedioPagoPlan As Byte
Private mIDCuentaCorrienteCaja As Long
Private mActivo As Boolean
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

Public Property Get Abreviatura() As String
    Abreviatura = mAbreviatura
End Property

Public Property Let Abreviatura(ByVal Value As String)
    If Value <> mAbreviatura Then
        mIsDirty = True
    End If
    mAbreviatura = Value
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

Public Property Get UtilizaOperacion() As Boolean
    UtilizaOperacion = mUtilizaOperacion
End Property

Public Property Let UtilizaOperacion(ByVal Value As Boolean)
    If Value <> mUtilizaOperacion Then
        mIsDirty = True
    End If
    mUtilizaOperacion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDMedioPagoPlan() As Byte
    IDMedioPagoPlan = mIDMedioPagoPlan
End Property

Public Property Let IDMedioPagoPlan(ByVal Value As Byte)
    If Value <> mIDMedioPagoPlan Then
        mIsDirty = True
    End If
    mIDMedioPagoPlan = Value
End Property

Public Property Get MedioPagoPlan() As MedioPagoPlan
    Static oMedioPagoPlan As MedioPagoPlan
    
    If mIDMedioPagoPlan = 0 Then
        Set MedioPagoPlan = Nothing
    Else
        If oMedioPagoPlan Is Nothing Then
            Set oMedioPagoPlan = New MedioPagoPlan
        End If
        If oMedioPagoPlan.IDMedioPagoPlan <> mIDMedioPagoPlan Then
            oMedioPagoPlan.IDMedioPagoPlan = mIDMedioPagoPlan
            Call oMedioPagoPlan.Load
        End If
        Set MedioPagoPlan = oMedioPagoPlan
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
    RefreshList_Module.RefreshList_RefreshMedioPago IDMedioPago
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
        .CommandText = "usp_MedioPago_Get"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDMedioPago", adTinyInt, adParamInput, , mIDMedioPago)
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
    ShowErrorMessage "Classes.MedioPago.Load", "Error al obtener los datos del Medio de Pago." & vbCr & vbCr & "IDMedioPago: " & IDMedioPago
End Function

Private Function GetDataFromRecordset(ByRef recData As ADODB.Recordset) As Boolean

    If (Not mNoMatchRaiseError) And recData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False

        mIDMedioPago = recData("IDMedioPago").Value
        mAbreviatura = recData("Abreviatura").Value
        mNombre = recData("Nombre").Value
        mUtilizaOperacion = recData("UtilizaOperacion").Value
        mIDMedioPagoPlan = Val(recData("IDMedioPagoPlan").Value & "")
        mIDCuentaCorrienteCaja = Val(recData("IDCuentaCorrienteCaja").Value & "")
        mActivo = recData("Activo").Value
        mFechaHoraCreacion = CDate(recData("FechaHoraCreacion").Value)
        mIDUsuarioCreacion = recData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = recData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = recData("IDUsuarioModificacion").Value
    End If

    Screen.MousePointer = vbDefault
    GetDataFromRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.MedioPago.GetDataFromRecordset", "Error al obtener los datos del Medio de Pago." & vbCr & vbCr & "IDMedioPago: " & IDMedioPago
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIDMedioPago = 0

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
        If mIsNew Then
            .CommandText = "usp_MedioPago_Add"
        Else
            .CommandText = "usp_MedioPago_Update"
        End If
        .CommandType = adCmdStoredProc

        If mIsNew Then
            .Parameters.Append .CreateParameter("IDMedioPago", adTinyInt, adParamOutput)
        Else
            .Parameters.Append .CreateParameter("IDMedioPago", adTinyInt, adParamInput, , mIDMedioPago)
        End If
        .Parameters.Append .CreateParameter("Abreviatura", adVarChar, adParamInput, 4, mAbreviatura)
        .Parameters.Append .CreateParameter("Nombre", adVarChar, adParamInput, 50, mNombre)
        .Parameters.Append .CreateParameter("UtilizaOperacion", adBoolean, adParamInput, , mUtilizaOperacion)
        .Parameters.Append .CreateParameter("IDMedioPagoPlan", adTinyInt, adParamInput, , IIf(mIDMedioPagoPlan = 0, Null, mIDMedioPagoPlan))
        .Parameters.Append .CreateParameter("IDCuentaCorrienteCaja", adInteger, adParamInput, , IIf(mIDCuentaCorrienteCaja = 0, Null, mIDCuentaCorrienteCaja))
        .Parameters.Append .CreateParameter("Activo", adBoolean, adParamInput, , mActivo)
        .Parameters.Append .CreateParameter("IDUsuario", adSmallInt, adParamInput, , pUsuario.IDUsuario)

        .Execute

        If mIsNew Then
            mIDMedioPago = .Parameters("IDMedioPago").Value
        End If
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
    ShowErrorMessage "Classes.MedioPago.Update", "Error al actualizar el Medio de Pago." & vbCr & vbCr & "IDMedioPago: " & IDMedioPago
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
        .CommandText = "usp_MedioPago_Delete"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDMedioPago", adTinyInt, adParamInput, , mIDMedioPago)

        .Execute
    End With

    If Not mRefreshListSkip Then
        Call RefreshList
    End If

    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.MedioPago.Delete", "Error al eliminar el Medio de Pago." & vbCr & vbCr & "IDMedioPago: " & IDMedioPago
End Function

Private Sub InitializeValues()
    mAbreviatura = ""
    mNombre = ""
    mUtilizaOperacion = False
    mIDMedioPagoPlan = 0
    mIDCuentaCorrienteCaja = 0
    Activo = True
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
