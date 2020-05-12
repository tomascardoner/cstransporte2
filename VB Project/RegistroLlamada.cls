VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistroLlamada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDRegistroLlamada As Long
Private mFechaHora As Date
Private mIDSucursalTelefono As Long
Private mTelefonoNumero As String

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

Public Property Get IDRegistroLlamada() As Long
    IDRegistroLlamada = mIDRegistroLlamada
End Property

Public Property Let IDRegistroLlamada(ByVal Value As Long)
    If Value <> mIDRegistroLlamada Then
        mIsDirty = True
    End If
    mIDRegistroLlamada = Value
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

Public Property Get FechaHora_Formatted() As String
    FechaHora_Formatted = Format(mFechaHora, "Short Date") & " " & Format(mFechaHora, "Short Time")
End Property

Public Property Get FechaHora_FormattedAsDate() As String
    FechaHora_FormattedAsDate = Format(mFechaHora, "Short Date")
End Property

Public Property Get FechaHora_FormattedAsTime() As String
    FechaHora_FormattedAsTime = Format(mFechaHora, "Short Time")
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDSucursalTelefono() As Long
    IDSucursalTelefono = mIDSucursalTelefono
End Property

Public Property Let IDSucursalTelefono(ByVal Value As Long)
    If Value <> mIDSucursalTelefono Then
        mIsDirty = True
    End If
    mIDSucursalTelefono = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get TelefonoNumero() As String
    TelefonoNumero = mTelefonoNumero
End Property

Public Property Let TelefonoNumero(ByVal Value As String)
    If Value <> mTelefonoNumero Then
        mIsDirty = True
    End If
    mTelefonoNumero = Value
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
    RefreshList_RefreshRegistroLlamada 0
End Sub

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function OpenRecordset() As Boolean
    Dim ConnectionRetryCount As Long
    
RETRY:
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    Screen.MousePointer = vbHourglass
      
    Set mrecData = New ADODB.Recordset
    Set mrecData.ActiveConnection = pDatabase.Connection
    mrecData.CursorType = adOpenKeyset
    mrecData.LockType = adLockOptimistic
    mrecData.Source = "SELECT * FROM RegistroLlamada"
    mrecData.Open , , , , adCmdText
    mrecData.Properties("Update Criteria").Value = adCriteriaKey

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    If ConnectionRetryCount = 0 Then
        ConnectionRetryCount = ConnectionRetryCount + 1
        If pDatabase.Connect() Then
            Resume RETRY
        End If
    End If
    ShowErrorMessage "Classes.RegistroLlamada.OpenRecordset", "Error al abrir la tabla de Registros de Llamadas."
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

    mrecData.Filter = "IDRegistroLlamada = " & mIDRegistroLlamada

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mFechaHora = mrecData("FechaHora").Value
        mIDSucursalTelefono = mrecData("IDSucursalTelefono").Value
        mTelefonoNumero = mrecData("TelefonoNumero").Value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.RegistroLlamada.Load", "Error al obtener los datos de el Registro de Llamada." & vbCr & vbCr & "IDRegistroLlamada: " & IDRegistroLlamada
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIDRegistroLlamada = 0

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
        recMaxID.Source = "SELECT MAX(IDRegistroLlamada) AS MaxID FROM RegistroLlamada"
        recMaxID.Open , , , , adCmdText
        mIDRegistroLlamada = Val(recMaxID("MaxID").Value & "") + 1
        recMaxID.Close
        Set recMaxID = Nothing
        mrecData.AddNew
        mrecData("IDRegistroLlamada").Value = mIDRegistroLlamada
    End If

    mrecData("FechaHora").Value = mFechaHora
    mrecData("IDSucursalTelefono").Value = mIDSucursalTelefono
    mrecData("TelefonoNumero").Value = mTelefonoNumero
    mrecData.Update

    mIsCopy = False
    mIsDirty = False
    
    If Not mRefreshListSkip Then
        RefreshList_RefreshRegistroLlamada IDRegistroLlamada
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
'            MsgBox "Ya existe un Registro de Llamada con el mismo nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.RegistroLlamada.Update", "Error al actualizar el Registro de Llamada." & vbCr & vbCr & "IDRegistroLlamada: " & IDRegistroLlamada
        End If
        If mrecData.EditMode = adEditAdd Or mrecData.EditMode = adEditInProgress Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.RegistroLlamada.Update", "Error al actualizar el Registro de Llamada." & vbCr & vbCr & "IDRegistroLlamada: " & IDRegistroLlamada
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
        RefreshList_RefreshRegistroLlamada 0
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede eliminar el Registro de Llamada debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.RegistroLlamada.Delete", "Error al eliminar el Registro de Llamada." & vbCr & vbCr & "IDRegistroLlamada: " & IDRegistroLlamada
        End If
        If mrecData.EditMode = adEditDelete Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.RegistroLlamada.Delete", "Error al eliminar el Registro de Llamada." & vbCr & vbCr & "IDRegistroLlamada: " & IDRegistroLlamada
    End If
End Function

Private Sub InitializeValues()
    mFechaHora = DATE_TIME_FIELD_NULL_VALUE
    mIDSucursalTelefono = 0
    mTelefonoNumero = ""

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
