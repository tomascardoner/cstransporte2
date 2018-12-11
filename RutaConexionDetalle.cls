VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RutaConexionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDRutaConexion As Long
Private mTramo1_IDRuta As String
Private mTramo2_IDRuta As String

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

Public Property Get IDRutaConexion() As Long
    IDRutaConexion = mIDRutaConexion
End Property

Public Property Let IDRutaConexion(ByVal Value As Long)
    If Value <> mIDRutaConexion Then
        mIsDirty = True
    End If
    mIDRutaConexion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Tramo1_IDRuta() As String
    Tramo1_IDRuta = mTramo1_IDRuta
End Property

Public Property Let Tramo1_IDRuta(ByVal Value As String)
    If Value <> mTramo1_IDRuta Then
        mIsDirty = True
    End If
    mTramo1_IDRuta = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Tramo2_IDRuta() As String
    Tramo2_IDRuta = mTramo2_IDRuta
End Property

Public Property Let Tramo2_IDRuta(ByVal Value As String)
    If Value <> mTramo2_IDRuta Then
        mIsDirty = True
    End If
    mTramo2_IDRuta = Value
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
    'RefreshList_Module.RutaConexionDetalle IDRutaConexion, Tramo1_IDRuta, Tramo2_IDRuta
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
    mrecData.Source = "SELECT * FROM RutaConexionDetalle"
    mrecData.Open , , , , adCmdText
    mrecData.Properties("Update Criteria").Value = adCriteriaKey

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.RutaConexionDetalle.OpenRecordset", "Error al abrir la tabla de Detalles de Conexiones de Ruta."
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

    mrecData.Filter = "IDRutaConexion = " & mIDRutaConexion & " AND Tramo1_IDRuta = '" & mTramo1_IDRuta & "' AND Tramo2_IDRuta = '" & mTramo2_IDRuta & "'"

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDRutaConexion = mrecData("IDRutaConexion").Value
        mTramo1_IDRuta = RTrim(mrecData("Tramo1_IDRuta").Value)
        mTramo2_IDRuta = RTrim(mrecData("Tramo2_IDRuta").Value)
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.RutaConexionDetalle.Load", "Error al obtener los datos del Detalle de Conexión de Ruta." & vbCr & vbCr & "IDRutaConexion: " & IDRutaConexion & vbCr & "Tramo1_IDRuta: " & Tramo1_IDRuta & vbCr & "Tramo2_IDRuta: " & Tramo2_IDRuta
End Function

Public Function LoadFirstWhereTramo1() As Boolean
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

    mrecData.Filter = "IDRutaConexion = " & mIDRutaConexion & " AND Tramo1_IDRuta = '" & mTramo1_IDRuta & "'"

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDRutaConexion = mrecData("IDRutaConexion").Value
        mTramo1_IDRuta = RTrim(mrecData("Tramo1_IDRuta").Value)
        mTramo2_IDRuta = RTrim(mrecData("Tramo2_IDRuta").Value)
    End If
    
    Screen.MousePointer = vbDefault
    LoadFirstWhereTramo1 = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.RutaConexionDetalle.LoadFirstWhereTramo1", "Error al obtener los datos del Detalle de Conexión de Ruta." & vbCr & vbCr & "IDRutaConexion: " & IDRutaConexion & vbCr & "Tramo1_IDRuta: " & Tramo1_IDRuta
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIDRutaConexion = 0
        mTramo1_IDRuta = ""
        mTramo2_IDRuta = ""

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
        mrecData("IDRutaConexion").Value = mIDRutaConexion
        mrecData("Tramo1_IDRuta").Value = mTramo1_IDRuta
        mrecData("Tramo2_IDRuta").Value = mTramo2_IDRuta
    End If

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
            MsgBox "Ya existe un Detalle de Conexión de Ruta con el mismo nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.RutaConexionDetalle.Update", "Error al actualizar el Detalle de Conexión de Ruta." & vbCr & vbCr & "IDRutaConexion: " & IDRutaConexion & vbCr & "Tramo1_IDRuta: " & Tramo1_IDRuta & vbCr & "Tramo2_IDRuta: " & Tramo2_IDRuta
        End If
        If mrecData.EditMode = adEditAdd Or mrecData.EditMode = adEditInProgress Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.RutaConexionDetalle.Update", "Error al actualizar el Detalle de Conexión de Ruta." & vbCr & vbCr & "IDRutaConexion: " & IDRutaConexion & vbCr & "Tramo1_IDRuta: " & Tramo1_IDRuta & vbCr & "Tramo2_IDRuta: " & Tramo2_IDRuta
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
            MsgBox "No se puede eliminar el Detalle de Conexión de Ruta debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.RutaConexionDetalle.Delete", "Error al eliminar el Detalle de Conexión de Ruta." & vbCr & vbCr & "IDRutaConexion: " & IDRutaConexion & vbCr & "Tramo1_IDRuta: " & Tramo1_IDRuta & vbCr & "Tramo2_IDRuta: " & Tramo2_IDRuta
        End If
        If mrecData.EditMode = adEditDelete Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.RutaConexionDetalle.Delete", "Error al eliminar el Detalle de Conexión de Ruta." & vbCr & vbCr & "IDRutaConexion: " & IDRutaConexion & vbCr & "Tramo1_IDRuta: " & Tramo1_IDRuta & vbCr & "Tramo2_IDRuta: " & Tramo2_IDRuta
    End If
End Function

Private Sub InitializeValues()
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
