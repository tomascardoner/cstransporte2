VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuentaCorrienteCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDCuentaCorrienteCaja As Long
Private mNombre As String
Private mIDPersona As Long
Private mMostrarSiempre As Boolean
Private mOcultarSaldo As Boolean
Private mNotas As String
Private mActivo As Boolean
Private mFechaHoraCreacion As Date
Private mIDUsuarioCreacion As Integer
Private mFechaHoraModificacion As Date
Private mIDUsuarioModificacion As Integer

Private mSaldoActual_Efectivo As Currency
Private mSaldoActual_Tarjeta As Currency

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

Public Property Get MostrarSiempre() As Boolean
    MostrarSiempre = mMostrarSiempre
End Property

Public Property Let MostrarSiempre(ByVal Value As Boolean)
    If Value <> mMostrarSiempre Then
        mIsDirty = True
    End If
    mMostrarSiempre = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get OcultarSaldo() As Boolean
    OcultarSaldo = mOcultarSaldo
End Property

Public Property Let OcultarSaldo(ByVal Value As Boolean)
    If Value <> mOcultarSaldo Then
        mIsDirty = True
    End If
    mOcultarSaldo = Value
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

Public Property Get SaldoActual_Efectivo_Formatted() As String
    SaldoActual_Efectivo_Formatted = Format(mSaldoActual_Efectivo, "Currency")
End Property

Public Property Get SaldoActual_Tarjeta_Formatted() As String
    SaldoActual_Tarjeta_Formatted = Format(mSaldoActual_Tarjeta, "Currency")
End Property

Public Property Get SaldoActual_Formatted() As String
    SaldoActual_Formatted = Format(mSaldoActual_Efectivo + mSaldoActual_Tarjeta, "Currency")
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
    RefreshList_Module.RefreshList_RefreshCuentaCorrienteCaja IDCuentaCorrienteCaja
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
    mrecData.Source = "SELECT * FROM CuentaCorrienteCaja"
    mrecData.Open , , , , adCmdText
    mrecData.Properties("Update Criteria").Value = adCriteriaKey

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.CuentaCorrienteCaja.OpenRecordset", "Error al abrir la tabla de Cajas de Cuenta Corriente."
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

    mrecData.Filter = "IDCuentaCorrienteCaja = " & mIDCuentaCorrienteCaja

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDCuentaCorrienteCaja = mrecData("IDCuentaCorrienteCaja").Value
        mNombre = mrecData("Nombre").Value
        mIDPersona = Val(mrecData("IDPersona").Value & "")
        mMostrarSiempre = mrecData("MostrarSiempre").Value
        mOcultarSaldo = mrecData("OcultarSaldo").Value
        mNotas = mrecData("Notas").Value & ""
        mActivo = mrecData("Activo").Value
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.CuentaCorrienteCaja.Load", "Error al obtener los datos de la Caja de Cuenta Corriente." & vbCr & vbCr & "IDCuentaCorrienteCaja: " & IDCuentaCorrienteCaja
End Function

Public Function LoadByPersona() As Boolean
    Dim SaveIDPersona As Long
    
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

    SaveIDPersona = mIDPersona
    Call InitializeValues
    mIDPersona = SaveIDPersona

    mrecData.Filter = "IDPersona = " & mIDPersona

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDCuentaCorrienteCaja = mrecData("IDCuentaCorrienteCaja").Value
        mNombre = mrecData("Nombre").Value
        mIDPersona = Val(mrecData("IDPersona").Value & "")
        mMostrarSiempre = mrecData("MostrarSiempre").Value
        mOcultarSaldo = mrecData("OcultarSaldo").Value
        mNotas = mrecData("Notas").Value & ""
        mActivo = mrecData("Activo").Value
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    
    Screen.MousePointer = vbDefault
    LoadByPersona = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.CuentaCorrienteCaja.LoadByPersona", "Error al obtener los datos de la Caja de Cuenta Corriente a partir de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIDCuentaCorrienteCaja = 0

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
        recMaxID.Source = "SELECT MAX(IDCuentaCorrienteCaja) AS MaxID FROM CuentaCorrienteCaja"
        recMaxID.Open , , , , adCmdText
        mIDCuentaCorrienteCaja = Val(recMaxID("MaxID").Value & "") + 1
        recMaxID.Close
        Set recMaxID = Nothing
        mrecData.AddNew
        mrecData("IDCuentaCorrienteCaja").Value = mIDCuentaCorrienteCaja
    End If

    mrecData("Nombre").Value = mNombre
    mrecData("IDPersona").Value = IIf(mIDPersona = 0, Null, mIDPersona)
    mrecData("MostrarSiempre").Value = mMostrarSiempre
    mrecData("OcultarSaldo").Value = mOcultarSaldo
    mrecData("Notas").Value = IIf(Trim(mNotas) = "", Null, mNotas)
    mrecData("Activo").Value = mActivo
    If mIsNew Then
        mrecData("FechaHoraCreacion").Value = Now
        mrecData("IDUsuarioCreacion").Value = pUsuario.IDUsuario
    End If
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If mIsNew Then
        pCPermiso.Add PERMISO_CUENTA_CORRIENTE_CAJA_CAJA & IDCuentaCorrienteCaja, "Cajas de Cuenta Corriente", "Ver Movimientos de: " & Nombre
    Else
        pCPermiso(PERMISO_CUENTA_CORRIENTE_CAJA_CAJA & IDCuentaCorrienteCaja).Descripcion = "Ver Movimientos de: " & Nombre
    End If

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
            MsgBox "Ya existe una Caja de Cuenta Corriente con el mismo nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.CuentaCorrienteCaja.Update", "Error al actualizar la Caja de Cuenta Corriente." & vbCr & vbCr & "IDCuentaCorrienteCaja: " & IDCuentaCorrienteCaja
        End If
        If mrecData.EditMode = adEditAdd Or mrecData.EditMode = adEditInProgress Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.CuentaCorrienteCaja.Update", "Error al actualizar la Caja de Cuenta Corriente." & vbCr & vbCr & "IDCuentaCorrienteCaja: " & IDCuentaCorrienteCaja
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
        Call RefreshList
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede eliminar la Caja de Cuenta Corriente debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.CuentaCorrienteCaja.Delete", "Error al eliminar la Caja de Cuenta Corriente." & vbCr & vbCr & "IDCuentaCorrienteCaja: " & IDCuentaCorrienteCaja
        End If
        If mrecData.EditMode = adEditDelete Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.CuentaCorrienteCaja.Delete", "Error al eliminar la Caja de Cuenta Corriente." & vbCr & vbCr & "IDCuentaCorrienteCaja: " & IDCuentaCorrienteCaja
    End If
End Function

Private Sub InitializeValues()
    mNombre = ""
    mIDPersona = 0
    mMostrarSiempre = False
    mOcultarSaldo = False
    mNotas = ""
    mActivo = True
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

Public Function LoadSaldoActual() As Boolean
    Dim cmdData As ADODB.command
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass
    
    mSaldoActual_Efectivo = 0
    mSaldoActual_Tarjeta = 0
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandType = adCmdStoredProc
    cmdData.CommandText = "sp_CuentaCorrienteCaja_SaldoActual"
    cmdData.Parameters.Append cmdData.CreateParameter("IDCuentaCorrienteCaja", adInteger, adParamInput, , IDCuentaCorrienteCaja)
    cmdData.Parameters.Append cmdData.CreateParameter("SaldoActual_Efectivo", adCurrency, adParamOutput)
    cmdData.Parameters.Append cmdData.CreateParameter("SaldoActual_Tarjeta", adCurrency, adParamOutput)
    cmdData.Execute
    
    mSaldoActual_Efectivo = cmdData.Parameters("SaldoActual_Efectivo").Value
    mSaldoActual_Tarjeta = cmdData.Parameters("SaldoActual_Tarjeta").Value
    
    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    LoadSaldoActual = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.CuentaCorrienteCaja.LoadSaldoActual", "Error al obtener el Saldo Actual de la Caja de Cuenta Corriente." & vbCr & vbCr & "IDCuentaCorrienteCaja: " & IDCuentaCorrienteCaja
End Function
