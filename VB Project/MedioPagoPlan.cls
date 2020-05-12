VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MedioPagoPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDMedioPagoPlan As Byte
Private mNombre As String
Private mActivo As Boolean
Private mFechaHoraCreacion As Date
Private mIDUsuarioCreacion As Integer
Private mFechaHoraModificacion As Date
Private mIDUsuarioModificacion As Integer

Private mCCuotas As Collection

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

Public Property Get IDMedioPagoPlan() As Byte
    IDMedioPagoPlan = mIDMedioPagoPlan
End Property

Public Property Let IDMedioPagoPlan(ByVal Value As Byte)
    If Value <> mIDMedioPagoPlan Then
        mIsDirty = True
    End If
    mIDMedioPagoPlan = Value
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

Public Property Get CCuotas() As Collection
    Set CCuotas = mCCuotas
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
    'RefreshList_Module.MedioPagoPlan IDMedioPagoPlan
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
        .CommandText = "usp_MedioPagoPlan_Get"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDMedioPagoPlan", adTinyInt, adParamInput, , mIDMedioPagoPlan)
    End With

    Set recData = New ADODB.Recordset
    recData.Open cmdSP, , adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

    Load = GetDataFromRecordset(recData)

    recData.Close
    Set recData = Nothing
    Set cmdSP = Nothing
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.MedioPagoPlan.Load", "Error al obtener los datos del Plan de Medios de Pago." & vbCr & vbCr & "IDMedioPagoPlan: " & IDMedioPagoPlan
End Function

Private Function GetDataFromRecordset(ByRef recData As ADODB.Recordset) As Boolean
    If (Not mNoMatchRaiseError) And recData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False

        mNombre = recData("Nombre").Value
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
    ShowErrorMessage "Classes.MedioPagoPlan.LoadCuotas", "Error al obtener los datos del Plan de Medios de Pago." & vbCr & vbCr & "IDMedioPagoPlan: " & IDMedioPagoPlan
End Function

Public Function LoadCuotas() As Boolean
    Dim cmdSP As ADODB.command
    Dim recData As ADODB.Recordset
    Dim MedioPagoPlanCuota As MedioPagoPlanCuota

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Set mCCuotas = New Collection

    Set cmdSP = New ADODB.command
    With cmdSP
        Set .ActiveConnection = pDatabase.Connection
        .CommandText = "usp_MedioPagoPlanCuota_List"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDMedioPagoPlan", adTinyInt, adParamInput, , mIDMedioPagoPlan)
    End With

    Set recData = New ADODB.Recordset
    recData.Open cmdSP, , adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

    Do While Not recData.EOF
        Set MedioPagoPlanCuota = New MedioPagoPlanCuota
        MedioPagoPlanCuota.IDMedioPagoPlan = mIDMedioPagoPlan
        MedioPagoPlanCuota.Cuota = recData("Cuota").Value
        MedioPagoPlanCuota.Coeficiente = recData("Coeficiente").Value
        MedioPagoPlanCuota.CoeficientePrepago = recData("CoeficientePrepago").Value
        
        mCCuotas.Add MedioPagoPlanCuota, KEY_STRINGER & MedioPagoPlanCuota.Cuota
        
        recData.MoveNext
    Loop

    Screen.MousePointer = vbDefault

    LoadCuotas = True

    recData.Close
    Set recData = Nothing
    Set cmdSP = Nothing
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.MedioPagoPlan.LoadCuotas", "Error al obtener los datos de las Cuotas del Plan de Medios de Pago." & vbCr & vbCr & "IDMedioPagoPlan: " & IDMedioPagoPlan
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIDMedioPagoPlan = 0

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
            .CommandText = "usp_MedioPagoPlan_Add"
        Else
            .CommandText = "usp_MedioPagoPlan_Update"
        End If
        .CommandType = adCmdStoredProc

        If mIsNew Then
            .Parameters.Append .CreateParameter("IDMedioPagoPlan", adTinyInt, adParamOutput)
        Else
            .Parameters.Append .CreateParameter("IDMedioPagoPlan", adTinyInt, adParamInput, , mIDMedioPagoPlan)
        End If
        .Parameters.Append .CreateParameter("Nombre", adVarChar, adParamInput, 50, mNombre)
        .Parameters.Append .CreateParameter("Activo", adBoolean, adParamInput, , mActivo)
        .Parameters.Append .CreateParameter("IDUsuario", adSmallInt, adParamInput, , pUsuario.IDUsuario)

        .Execute

        If mIsNew Then
            mIDMedioPagoPlan = .Parameters("IDMedioPagoPlan").Value
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
    ShowErrorMessage "Classes.MedioPagoPlan.Update", "Error al actualizar el Plan de Medios de Pago." & vbCr & vbCr & "IDMedioPagoPlan: " & IDMedioPagoPlan
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
        .CommandText = "usp_MedioPagoPlan_Delete"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("IDMedioPagoPlan", adTinyInt, adParamInput, , mIDMedioPagoPlan)

        .Execute
    End With

    If Not mRefreshListSkip Then
        Call RefreshList
    End If

    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.MedioPagoPlan.Delete", "Error al eliminar el MedioPagoPlan." & vbCr & vbCr & "IDMedioPagoPlan: " & IDMedioPagoPlan
End Function

Private Sub InitializeValues()
    mNombre = ""
    mActivo = True
    mFechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioCreacion = 0
    mFechaHoraModificacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioModificacion = 0
    
    Set mCCuotas = New Collection

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
