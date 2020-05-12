VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListaPrecio_Ruta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDListaPrecio As Long
Private mIDRuta As String
Private mRutaNombre As String
Private mOrigenNombre As String
Private mDestinoNombre As String

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

Public Property Get RutaNombre() As String
    RutaNombre = mRutaNombre
End Property

Public Property Let RutaNombre(ByVal Value As String)
    If Value <> mRutaNombre Then
        mIsDirty = True
    End If
    mRutaNombre = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get OrigenNombre() As String
    OrigenNombre = mOrigenNombre
End Property

Public Property Let OrigenNombre(ByVal Value As String)
    If Value <> mOrigenNombre Then
        mIsDirty = True
    End If
    mOrigenNombre = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get DestinoNombre() As String
    DestinoNombre = mDestinoNombre
End Property

Public Property Let DestinoNombre(ByVal Value As String)
    If Value <> mDestinoNombre Then
        mIsDirty = True
    End If
    mDestinoNombre = Value
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
    'RefreshList_Module.RefreshList_RefreshListaPrecio_Ruta IDListaPrecio, IDRuta
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
    mrecData.Source = "SELECT * FROM ListaPrecio_Ruta"
    mrecData.Open , , , , adCmdText
    mrecData.Properties("Update Criteria").Value = adCriteriaKey

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.ListaPrecio_Ruta.OpenRecordset", "Error al abrir la tabla de Listas de Precios por Ruta."
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

    mrecData.Filter = "IDListaPrecio = " & mIDListaPrecio & " AND IDRuta = '" & mIDRuta & "'"

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDListaPrecio = mrecData("IDListaPrecio").Value
        mIDRuta = RTrim(mrecData("IDRuta").Value)
        mRutaNombre = mrecData("RutaNombre").Value & ""
        mOrigenNombre = mrecData("OrigenNombre").Value & ""
        mDestinoNombre = mrecData("DestinoNombre").Value & ""
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.ListaPrecio_Ruta.Load", "Error al obtener los datos de la Lista de Precios por Ruta." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio & vbCr & "IDRuta: " & IDRuta
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIDListaPrecio = 0
        mIDRuta = ""

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
        mrecData("IDListaPrecio").Value = mIDListaPrecio
        mrecData("IDRuta").Value = mIDRuta
    End If
    mrecData("RutaNombre").Value = IIf(Trim(mRutaNombre) = "", Null, mRutaNombre)
    mrecData("OrigenNombre").Value = IIf(Trim(mOrigenNombre) = "", Null, mOrigenNombre)
    mrecData("DestinoNombre").Value = IIf(Trim(mDestinoNombre) = "", Null, mDestinoNombre)
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
            MsgBox "Ya existe una la Lista de Precios por Ruta con el mismo nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.ListaPrecio_Ruta.Update", "Error al actualizar la Lista de Precios por Ruta." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio & vbCr & "IDRuta: " & IDRuta
        End If
        If mrecData.EditMode = adEditAdd Or mrecData.EditMode = adEditInProgress Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.ListaPrecio_Ruta.Update", "Error al actualizar la Lista de Precios por Ruta." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio & vbCr & "IDRuta: " & IDRuta
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
            MsgBox "No se puede eliminar la Lista de Precios por Ruta debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.ListaPrecio_Ruta.Delete", "Error al eliminar la Lista de Precios por Ruta." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio & vbCr & "IDRuta: " & IDRuta
        End If
        If mrecData.EditMode = adEditDelete Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.ListaPrecio_Ruta.Delete", "Error al eliminar la Lista de Precios por Ruta." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio & vbCr & "IDRuta: " & IDRuta
    End If
End Function

Private Sub InitializeValues()
    mRutaNombre = ""
    mOrigenNombre = ""
    mDestinoNombre = ""

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
