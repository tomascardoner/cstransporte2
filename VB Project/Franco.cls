VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Franco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mFecha As Date
Private mIDPersona As Long
Private mImporte As Currency
Private mIDMovimientoCuentaCorriente As Long
Private mIDUsuarioCreacion As Integer
Private mFechaHoraCreacion As Date
Private mIDUsuarioModificacion As Integer
Private mFechaHoraModificacion As Date

Private mFechaOriginal As Date
Private mIDPersonaOriginal As Long

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

Public Property Get Fecha() As Date
    Fecha = mFecha
End Property

Public Property Get Fecha_Formatted() As String
    Fecha_Formatted = Format(mFecha, "Short Date") & " " & Format(mFecha, "Short Time")
End Property

Public Property Get Fecha_FormattedAsDate() As String
    Fecha_FormattedAsDate = Format(mFecha, "Short Date")
End Property

Public Property Get Fecha_FormattedAsTime() As String
    Fecha_FormattedAsTime = Format(mFecha, "Short Time")
End Property

Public Property Let Fecha(ByVal value As Date)
    If value <> mFecha Then
        mIsDirty = True
    End If
    mFecha = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDPersona() As Long
    IDPersona = mIDPersona
End Property

Public Property Let IDPersona(ByVal value As Long)
    If value <> mIDPersona Then
        mIsDirty = True
    End If
    mIDPersona = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Importe() As Currency
    Importe = mImporte
End Property

Public Property Get Importe_FormattedAsString() As String
    Importe_FormattedAsString = IIf(mImporte = -1, "", Format(mImporte, "Currency"))
End Property

Public Property Let Importe(ByVal value As Currency)
    If value <> mImporte Then
        mIsDirty = True
    End If
    mImporte = value
End Property

Public Property Let Importe_FormattedAsString(ByVal value As String)
    If IsNumeric(value) Then
        Importe = CCur(value)
    Else
        Importe = -1
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDMovimientoCuentaCorriente() As Long
    IDMovimientoCuentaCorriente = mIDMovimientoCuentaCorriente
End Property

Public Property Let IDMovimientoCuentaCorriente(ByVal value As Long)
    If value <> mIDMovimientoCuentaCorriente Then
        mIsDirty = True
    End If
    mIDMovimientoCuentaCorriente = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDUsuarioCreacion() As String
    IDUsuarioCreacion = mIDUsuarioCreacion
End Property

Public Property Let IDUsuarioCreacion(ByVal value As String)
    If value <> mIDUsuarioCreacion Then
        mIsDirty = True
    End If
    mIDUsuarioCreacion = value
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

Public Property Let FechaHoraCreacion(ByVal value As Date)
    If value <> mFechaHoraCreacion Then
        mIsDirty = True
    End If
    mFechaHoraCreacion = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDUsuarioModificacion() As String
    IDUsuarioModificacion = mIDUsuarioModificacion
End Property

Public Property Let IDUsuarioModificacion(ByVal value As String)
    If value <> mIDUsuarioModificacion Then
        mIsDirty = True
    End If
    mIDUsuarioModificacion = value
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

Public Property Let FechaHoraModificacion(ByVal value As Date)
    If value <> mFechaHoraModificacion Then
        mIsDirty = True
    End If
    mFechaHoraModificacion = value
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

Public Property Let NoMatchRaiseError(ByVal value As Boolean)
    mNoMatchRaiseError = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get RefreshListSkip() As Boolean
    RefreshListSkip = mRefreshListSkip
End Property

Public Property Let RefreshListSkip(ByVal value As Boolean)
    mRefreshListSkip = value
End Property

Public Sub RefreshList()
    RefreshList_Module.Franco Fecha, IDPersona
End Sub

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Database() As CSC_Database_ADO_SQL
    Set Database = mDatabase
End Property

Public Property Set Database(ByRef value As CSC_Database_ADO_SQL)
    Set mDatabase = value
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
    mrecData.Source = "SELECT * FROM Franco"
    mrecData.Open , , , , adCmdText
    mrecData.Properties("Update Criteria").value = adCriteriaKey

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Franco.OpenRecordset", "Error al abrir la tabla de Francos."
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

    mrecData.Filter = "Fecha = #" & Format(mFecha, "yyyy/mm/dd") & "# AND IDPersona = " & mIDPersona

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mFecha = mrecData("Fecha").value
        mIDPersona = mrecData("IDPersona").value
        mImporte = IIf(IsNull(mrecData("Importe").value), -1, mrecData("Importe").value)
        mIDMovimientoCuentaCorriente = Val(mrecData("IDMovimientoCuentaCorriente").value & "")
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").value
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").value
    
        mFechaOriginal = mFecha
        mIDPersonaOriginal = mIDPersona
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Franco.Load", "Error al obtener los datos del Franco." & vbCr & vbCr & "Fecha: " & Fecha_FormattedAsDate & vbCr & "IDPersona: " & IDPersona
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mFecha = DATE_TIME_FIELD_NULL_VALUE
        mIDPersona = 0

        mIsNew = True
        mIsCopy = True
        mIsDirty = True
    End If
End Function

Public Function Update() As Boolean
    Dim cmdData As ADODB.command
    
    If Not mIsDirty Then
        Update = True
        Exit Function
    End If

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Franco_Update"
    cmdData.CommandType = adCmdStoredProc
    cmdData.NamedParameters = True
    cmdData.Parameters.Append cmdData.CreateParameter("@FechaOriginal", adDate, adParamInput, , mFechaOriginal)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDPersonaOriginal", adInteger, adParamInput, , mIDPersonaOriginal)
    cmdData.Parameters.Append cmdData.CreateParameter("@Fecha", adDate, adParamInput, , mFecha)
    cmdData.Parameters.Append cmdData.CreateParameter("@IDPersona", adInteger, adParamInput, , mIDPersona)
    cmdData.Parameters.Append cmdData.CreateParameter("@Importe", adCurrency, adParamInput, , IIf(mImporte = -1, Null, mImporte))
    cmdData.Parameters.Append cmdData.CreateParameter("@IDUsuario", adSmallInt, adParamInput, 30, pUsuario.IDUsuario)
    cmdData.Execute 0
    Set cmdData = Nothing
    
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
            MsgBox "Ya existe un Franco con el mismo nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Franco.Update", "Error al actualizar el Franco." & vbCr & vbCr & "Fecha: " & mFecha & vbCr & "IDPersona: " & mIDPersona
        End If
    Else
        ShowErrorMessage "Classes.Franco.Update", "Error al actualizar el Franco." & vbCr & vbCr & "Fecha: " & Fecha_FormattedAsDate & vbCr & "IDPersona: " & IDPersona
    End If
End Function

Public Function Delete() As Boolean
    Dim CuentaCorriente As CuentaCorriente
    
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
    
    Set CuentaCorriente = New CuentaCorriente
    With CuentaCorriente
        If mIDMovimientoCuentaCorriente <> 0 Then
            .IDMovimiento = mIDMovimientoCuentaCorriente
            .NoMatchRaiseError = False
            If Not .Load() Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            If Not .NoMatch Then
                If Not .Delete() Then
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
            End If
        End If
    End With
    Set CuentaCorriente = Nothing
    
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
            MsgBox "No se puede eliminar el Franco debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Franco.Delete", "Error al eliminar el Franco." & vbCr & vbCr & "Fecha: " & Fecha_FormattedAsDate & vbCr & "IDPersona: " & IDPersona
        End If
        If mrecData.EditMode = adEditDelete Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.Franco.Delete", "Error al eliminar el Franco." & vbCr & vbCr & "Fecha: " & Fecha_FormattedAsDate & vbCr & "IDPersona: " & IDPersona
    End If
End Function

Private Sub InitializeValues()
    mImporte = -1
    mIDMovimientoCuentaCorriente = 0
    mIDUsuarioCreacion = 0
    mFechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    mIDUsuarioModificacion = 0
    mFechaHoraModificacion = DATE_TIME_FIELD_NULL_VALUE
    
    mFechaOriginal = DATE_TIME_FIELD_NULL_VALUE
    mIDPersonaOriginal = 0

    mIsNew = True
    mIsCopy = False
    mIsDirty = False
End Sub

Private Sub Class_Initialize()
    mFecha = DATE_TIME_FIELD_NULL_VALUE
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
