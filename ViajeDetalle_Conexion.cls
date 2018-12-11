VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViajeDetalle_Conexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDViajeDetalle As Long
Private mConexion_IDViajeDetalle As Long

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

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDViajeDetalle() As Long
    IDViajeDetalle = mIDViajeDetalle
End Property

Public Property Get IDViajeDetalle_Formatted() As String
    IDViajeDetalle_Formatted = Format(mIDViajeDetalle, "#,###")
End Property

Public Property Let IDViajeDetalle(ByVal Value As Long)
    If Value <> mIDViajeDetalle Then
        mIsDirty = True
    End If
    mIDViajeDetalle = Value
End Property

Public Property Get ViajeDetalle() As ViajeDetalle
    Static oViajeDetalle As ViajeDetalle
    
    If mIDViajeDetalle = 0 Then
        Set ViajeDetalle = Nothing
    Else
        If oViajeDetalle Is Nothing Then
            Set oViajeDetalle = New ViajeDetalle
        End If
        If oViajeDetalle.IDViajeDetalle <> mIDViajeDetalle Then
            oViajeDetalle.IDViajeDetalle = mIDViajeDetalle
            Call oViajeDetalle.LoadByIDViajeDetalle
        End If
        Set ViajeDetalle = oViajeDetalle
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Conexion_IDViajeDetalle() As Long
    Conexion_IDViajeDetalle = mConexion_IDViajeDetalle
End Property

Public Property Get Conexion_IDViajeDetalle_Formatted() As String
    Conexion_IDViajeDetalle_Formatted = Format(mConexion_IDViajeDetalle, "#,###")
End Property

Public Property Let Conexion_IDViajeDetalle(ByVal Value As Long)
    If Value <> mConexion_IDViajeDetalle Then
        mIsDirty = True
    End If
    mConexion_IDViajeDetalle = Value
End Property

Public Property Get Conexion_ViajeDetalle() As ViajeDetalle
    Static oViajeDetalle As ViajeDetalle
    
    If mConexion_IDViajeDetalle = 0 Then
        Set Conexion_ViajeDetalle = Nothing
    Else
        If oViajeDetalle Is Nothing Then
            Set oViajeDetalle = New ViajeDetalle
        End If
        If oViajeDetalle.IDViajeDetalle <> mConexion_IDViajeDetalle Then
            oViajeDetalle.IDViajeDetalle = mConexion_IDViajeDetalle
            Call oViajeDetalle.LoadByIDViajeDetalle
        End If
        Set Conexion_ViajeDetalle = oViajeDetalle
    End If
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
    'RefreshList_Module.ViajeDetalle_Conexion FechaHora, IDRuta, Indice
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

Public Function OpenRecordset() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    Screen.MousePointer = vbHourglass
      
    Set mrecData = New ADODB.Recordset
    Set mrecData.ActiveConnection = pDatabase.Connection
    mrecData.CursorType = adOpenKeyset
    mrecData.LockType = adLockOptimistic
    mrecData.Source = "SELECT * FROM ViajeDetalle_Conexion"
    mrecData.Open , , , , adCmdText
    mrecData.Properties("Update Criteria").Value = adCriteriaKey

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.ViajeDetalle_Conexion.OpenRecordset", "Error al abrir la tabla de Conexiones de Reservas."
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

    mrecData.Filter = "IDViajeDetalle = " & mIDViajeDetalle

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mConexion_IDViajeDetalle = mrecData("Conexion_IDViajeDetalle").Value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.ViajeDetalle_Conexion.Load", "Error al obtener los datos de la Conexión de la Reserva." & vbCr & vbCr & "IDViajeDetalle: " & mIDViajeDetalle
End Function

Public Function LoadByConexion() As Boolean
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

    mIDViajeDetalle = 0

    mrecData.Filter = "Conexion_IDViajeDetalle = " & mConexion_IDViajeDetalle

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDViajeDetalle = mrecData("IDViajeDetalle").Value
    End If
    
    Screen.MousePointer = vbDefault
    LoadByConexion = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.ViajeDetalle_Conexion.LoadByConexion", "Error al obtener los datos de la Conexión de la Reserva." & vbCr & vbCr & "Conexion_IDViajeDetalle: " & Conexion_IDViajeDetalle
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True
        
        mIDViajeDetalle = 0

        mIsNew = True
        mIsCopy = True
        mIsDirty = True
    End If
End Function

Public Function Update() As Boolean
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
        mrecData.AddNew
        mrecData("IDViajeDetalle").Value = mIDViajeDetalle
    End If

    mrecData("Conexion_IDViajeDetalle").Value = mConexion_IDViajeDetalle
    mrecData.Update

    mIsNew = False
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
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe una Conexión para esta Reserva.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.ViajeDetalle_Conexion.Update", "Error al actualizar la Conexión de la Reserva." & vbCr & vbCr & "IDViajeDetalle: " & mIDViajeDetalle
        End If
        If mrecData.EditMode = adEditAdd Or mrecData.EditMode = adEditInProgress Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.ViajeDetalle_Conexion.Update", "Error al actualizar la Conexión de la Reserva." & vbCr & vbCr & "IDViajeDetalle: " & mIDViajeDetalle
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
        Call RefreshList
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede eliminar la Conexión de la Reserva debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.ViajeDetalle_Conexion.Delete", "Error al eliminar la Conexión de la Reserva." & vbCr & vbCr & "IDViajeDetalle: " & mIDViajeDetalle
        End If
        If mrecData.EditMode = adEditDelete Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.ViajeDetalle_Conexion.Delete", "Error al eliminar la Conexión de la Reserva." & vbCr & vbCr & "IDViajeDetalle: " & mIDViajeDetalle
    End If
End Function

Private Sub InitializeValues()
    mConexion_IDViajeDetalle = 0

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
