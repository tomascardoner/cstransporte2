VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogAccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDLogAccion As Long
Private mEntidadTipo As String
Private mFechaHora As Date
Private mDescripcion As String
Private mIDUsuario As String

'STATE VARIABLES
Private mIsNew As Boolean
Private mIsCopy As Boolean
Private mIsDirty As Boolean
Private mNoMatch As Boolean

'BEHAVIOR VARIABLES
Private mNoMatchRaiseError As Boolean
Private mRefreshListSkip As Boolean

'INTERNAL VARIABLES
Private mrecData As ADODB.Recordset

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDLogAccion() As Long
    IDLogAccion = mIDLogAccion
End Property

Public Property Let IDLogAccion(ByVal Value As Long)
    If Value <> mIDLogAccion Then
        mIsDirty = True
    End If
    mIDLogAccion = Value
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

Public Property Get FechaHora() As Date
    FechaHora = mFechaHora
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

Public Property Get IDUsuario() As String
    IDUsuario = mIDUsuario
End Property

Public Property Let IDUsuario(ByVal Value As String)
    If Value <> mIDUsuario Then
        mIsDirty = True
    End If
    mIDUsuario = Value
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
    'RefreshList_RefreshLogAccion IDLogAccion
End Sub

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function OpenRecordset() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    Screen.MousePointer = vbHourglass
      
    Set mrecData = New ADODB.Recordset
    Set mrecData.ActiveConnection = pDatabase.Connection
    mrecData.CursorType = adOpenKeyset
    mrecData.LockType = adLockOptimistic
    mrecData.Source = "SELECT * FROM LogAccion"
    mrecData.Open , , , , adCmdText
    mrecData.Properties("Update Criteria").Value = adCriteriaKey

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.LogAccion.OpenRecordset", "Error al abrir la tabla de Logs de Acciones."
    If Not mrecData Is Nothing Then
        Set mrecData = Nothing
    End If
End Function

Public Function Load() As Boolean
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

    Call InitializeValues

    mrecData.Filter = "IDLogAccion = " & mIDLogAccion

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mEntidadTipo = mrecData("EntidadTipo").Value & ""
        mFechaHora = mrecData("FechaHora").Value
        mDescripcion = mrecData("Descripcion").Value
        mIDUsuario = mrecData("IDUsuario").Value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.LogAccion.Load", "Error al obtener los datos de el Log de Acción." & vbCr & vbCr & "IDLogAccion: " & IDLogAccion
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIDLogAccion = 0

        mIsNew = True
        mIsCopy = True
        mIsDirty = True
    End If
End Function

Public Function Update() As Boolean
    Dim recMaxID As ADODB.Recordset

    If Not mIsDirty Then
        Update = True
        Exit Function
    End If

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
    
    If mIsNew Then
        'GET NEW ID
        Set recMaxID = New ADODB.Recordset
        Set recMaxID.ActiveConnection = pDatabase.Connection
        recMaxID.CursorType = adOpenForwardOnly
        recMaxID.LockType = adLockReadOnly
        recMaxID.Source = "SELECT MAX(IDLogAccion) AS MaxID FROM LogAccion"
        recMaxID.Open , , , , adCmdText
        mIDLogAccion = Val(recMaxID("MaxID").Value & "") + 1
        recMaxID.Close
        Set recMaxID = Nothing
        mrecData.AddNew
        mrecData("IDLogAccion").Value = mIDLogAccion
    End If
    mrecData("EntidadTipo").Value = IIf(mEntidadTipo = "", Null, mEntidadTipo)
    mrecData("FechaHora").Value = mFechaHora
    mrecData("Descripcion").Value = mDescripcion
    mrecData("IDUsuario").Value = mIDUsuario
    mrecData.Update

    mIsCopy = False
    mIsDirty = False
    
    If Not mRefreshListSkip Then
        'RefreshList_RefreshLogAccion IDLogAccion
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe un Log de Acción con el mismo nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.LogAccion.Update", "Error al actualizar el Log de Acción." & vbCr & vbCr & "IDLogAccion: " & IDLogAccion
        End If
        If mrecData.EditMode = adEditAdd Or mrecData.EditMode = adEditInProgress Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.LogAccion.Update", "Error al actualizar el Log de Acción." & vbCr & vbCr & "IDLogAccion: " & IDLogAccion
    End If
    If Not recMaxID Is Nothing Then
        If recMaxID.State = adStateOpen Then
            recMaxID.Close
        End If
        Set recMaxID = Nothing
    End If
End Function

Public Function Delete() As Boolean
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
    
    mrecData.Delete
    
    If Not mRefreshListSkip Then
        'RefreshList_RefreshLogAccion 0
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede eliminar el Log de Acción debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.LogAccion.Delete", "Error al eliminar el Log de Acción." & vbCr & vbCr & "IDLogAccion: " & IDLogAccion
        End If
        If mrecData.EditMode = adEditDelete Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.LogAccion.Delete", "Error al eliminar el Log de Acción." & vbCr & vbCr & "IDLogAccion: " & IDLogAccion
    End If
End Function

Private Sub InitializeValues()
    mEntidadTipo = ""
    mFechaHora = Now
    mDescripcion = ""
    mIDUsuario = pUsuario.IDUsuario

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
