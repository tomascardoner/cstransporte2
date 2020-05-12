VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ruta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDRuta As String
Private mNombre As String
Private mIDOrigen As Long
Private mIDDestino As Long
Private mIDRutaGrupo As Long
Private mKilometro As Integer
Private mDuracion As Integer
Private mLimiteCancelacionIDLugar As Long
Private mLimiteCancelacionDuracion As Integer
Private mPermite2Conductores As Boolean
Private mConductorImporteTramoCompleto As Currency
Private mConductorImporteTramo1 As Currency
Private mConductorImporteTramo2 As Currency
Private mNotas As String
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

'DATABASE VARIABLES
Private mDatabase As CSC_Database_ADO_SQL

'INTERNAL VARIABLES
Private mrecData As ADODB.Recordset

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

Public Property Get IDOrigen() As Long
    IDOrigen = mIDOrigen
End Property

Public Property Let IDOrigen(ByVal Value As Long)
    If Value <> mIDOrigen Then
        mIsDirty = True
    End If
    mIDOrigen = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDDestino() As Long
    IDDestino = mIDDestino
End Property

Public Property Let IDDestino(ByVal Value As Long)
    If Value <> mIDDestino Then
        mIsDirty = True
    End If
    mIDDestino = Value
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

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Kilometro() As Integer
    Kilometro = mKilometro
End Property

Public Property Let Kilometro(ByVal Value As Integer)
    If Value <> mKilometro Then
        mIsDirty = True
    End If
    mKilometro = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Duracion() As Integer
    Duracion = mDuracion
End Property

Public Property Let Duracion(ByVal Value As Integer)
    If Value <> mDuracion Then
        mIsDirty = True
    End If
    mDuracion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get LimiteCancelacionIDLugar() As Long
    LimiteCancelacionIDLugar = mLimiteCancelacionIDLugar
End Property

Public Property Let LimiteCancelacionIDLugar(ByVal Value As Long)
    If Value <> mLimiteCancelacionIDLugar Then
        mIsDirty = True
    End If
    mLimiteCancelacionIDLugar = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get LimiteCancelacionDuracion() As Integer
    LimiteCancelacionDuracion = mLimiteCancelacionDuracion
End Property

Public Property Let LimiteCancelacionDuracion(ByVal Value As Integer)
    If Value <> mLimiteCancelacionDuracion Then
        mIsDirty = True
    End If
    mLimiteCancelacionDuracion = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Permite2Conductores() As Boolean
    Permite2Conductores = mPermite2Conductores
End Property

Public Property Let Permite2Conductores(ByVal Value As Boolean)
    If Value <> mPermite2Conductores Then
        mIsDirty = True
    End If
    mPermite2Conductores = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ConductorImporteTramoCompleto() As Currency
    ConductorImporteTramoCompleto = mConductorImporteTramoCompleto
End Property

Public Property Get ConductorImporteTramoCompleto_FormattedAsString() As String
    ConductorImporteTramoCompleto_FormattedAsString = IIf(mConductorImporteTramoCompleto = -1, "", Format(mConductorImporteTramoCompleto, "Currency"))
End Property

Public Property Let ConductorImporteTramoCompleto(ByVal Value As Currency)
    If Value <> mConductorImporteTramoCompleto Then
        mIsDirty = True
    End If
    mConductorImporteTramoCompleto = Value
End Property

Public Property Let ConductorImporteTramoCompleto_FormattedAsString(ByVal Value As String)
    If IsNumeric(Value) Then
        ConductorImporteTramoCompleto = CCur(Value)
    Else
        ConductorImporteTramoCompleto = -1
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ConductorImporteTramo1() As Currency
    ConductorImporteTramo1 = mConductorImporteTramo1
End Property

Public Property Get ConductorImporteTramo1_FormattedAsString() As String
    ConductorImporteTramo1_FormattedAsString = IIf(mConductorImporteTramo1 = -1, "", Format(mConductorImporteTramo1, "Currency"))
End Property

Public Property Let ConductorImporteTramo1(ByVal Value As Currency)
    If Value <> mConductorImporteTramo1 Then
        mIsDirty = True
    End If
    mConductorImporteTramo1 = Value
End Property

Public Property Let ConductorImporteTramo1_FormattedAsString(ByVal Value As String)
    If IsNumeric(Value) Then
        ConductorImporteTramo1 = CCur(Value)
    Else
        ConductorImporteTramo1 = -1
    End If
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ConductorImporteTramo2() As Currency
    ConductorImporteTramo2 = mConductorImporteTramo2
End Property

Public Property Get ConductorImporteTramo2_FormattedAsString() As String
    ConductorImporteTramo2_FormattedAsString = IIf(mConductorImporteTramo2 = -1, "", Format(mConductorImporteTramo2, "Currency"))
End Property

Public Property Let ConductorImporteTramo2(ByVal Value As Currency)
    If Value <> mConductorImporteTramo2 Then
        mIsDirty = True
    End If
    mConductorImporteTramo2 = Value
End Property

Public Property Let ConductorImporteTramo2_FormattedAsString(ByVal Value As String)
    If IsNumeric(Value) Then
        ConductorImporteTramo2 = CCur(Value)
    Else
        ConductorImporteTramo2 = -1
    End If
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

Public Property Get Activo() As Integer
    Activo = mActivo
End Property

Public Property Let Activo(ByVal Value As Integer)
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
    RefreshList_Module.RefreshList_RefreshRuta IDRuta
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
    mrecData.Source = "SELECT * FROM Ruta"
    mrecData.Open , , , , adCmdText
    mrecData.Properties("Update Criteria").Value = adCriteriaKey

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Ruta.OpenRecordset", "Error al abrir la tabla de Rutas."
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

    mrecData.Filter = "IDRuta = '" & mIDRuta & "'"

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDRuta = RTrim(mrecData("IDRuta").Value)
        mNombre = mrecData("Nombre").Value
        mIDOrigen = mrecData("IDOrigen").Value
        mIDDestino = mrecData("IDDestino").Value
        mIDRutaGrupo = mrecData("IDRutaGrupo").Value
        mKilometro = Val(mrecData("Kilometro").Value & "")
        mDuracion = Val(mrecData("Duracion").Value & "")
        mLimiteCancelacionIDLugar = Val(mrecData("LimiteCancelacionIDLugar").Value & "")
        mLimiteCancelacionDuracion = Val(mrecData("LimiteCancelacionDuracion").Value & "")
        mPermite2Conductores = mrecData("Permite2Conductores").Value
        mConductorImporteTramoCompleto = IIf(IsNull(mrecData("ConductorImporteTramoCompleto").Value), -1, mrecData("ConductorImporteTramoCompleto").Value)
        mConductorImporteTramo1 = IIf(IsNull(mrecData("ConductorImporteTramo1").Value), -1, mrecData("ConductorImporteTramo1").Value)
        mConductorImporteTramo2 = IIf(IsNull(mrecData("ConductorImporteTramo2").Value), -1, mrecData("ConductorImporteTramo2").Value)
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
    ShowErrorMessage "Classes.Ruta.Load", "Error al obtener los datos de la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

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
        mrecData("IDRuta").Value = mIDRuta
        mrecData("FechaHoraCreacion").Value = Now
        mrecData("IDUsuarioCreacion").Value = pUsuario.IDUsuario
    End If
    mrecData("Nombre").Value = mNombre
    mrecData("IDOrigen").Value = mIDOrigen
    mrecData("IDDestino").Value = mIDDestino
    mrecData("IDRutaGrupo").Value = mIDRutaGrupo
    mrecData("Kilometro").Value = IIf(mKilometro = 0, Null, mKilometro)
    mrecData("Duracion").Value = IIf(mDuracion = 0, Null, mDuracion)
    mrecData("LimiteCancelacionIDLugar").Value = IIf(mLimiteCancelacionIDLugar = 0, Null, mLimiteCancelacionIDLugar)
    mrecData("LimiteCancelacionDuracion").Value = IIf(mLimiteCancelacionDuracion = 0, Null, mLimiteCancelacionDuracion)
    mrecData("Permite2Conductores").Value = mPermite2Conductores
    mrecData("ConductorImporteTramoCompleto").Value = IIf(mConductorImporteTramoCompleto = -1, Null, mConductorImporteTramoCompleto)
    mrecData("ConductorImporteTramo1").Value = IIf(mConductorImporteTramo1 = -1, Null, mConductorImporteTramo1)
    mrecData("ConductorImporteTramo2").Value = IIf(mConductorImporteTramo2 = -1, Null, mConductorImporteTramo2)
    mrecData("Notas").Value = IIf(Trim(Notas) = "", Null, Notas)
    mrecData("Activo").Value = mActivo
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
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
            MsgBox "Ya existe una Ruta con el mismo nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Ruta.Update", "Error al actualizar la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta
        End If
        If mrecData.EditMode = adEditAdd Or mrecData.EditMode = adEditInProgress Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.Ruta.Update", "Error al actualizar la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta
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
            MsgBox "No se puede eliminar la Ruta debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Ruta.Delete", "Error al eliminar la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta
        End If
        If mrecData.EditMode = adEditDelete Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.Ruta.Delete", "Error al eliminar la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta
    End If
End Function

Private Sub InitializeValues()
    mNombre = ""
    mIDOrigen = 0
    mIDDestino = 0
    mIDRutaGrupo = 0
    mKilometro = 0
    mDuracion = 0
    mLimiteCancelacionIDLugar = 0
    mLimiteCancelacionDuracion = 0
    mPermite2Conductores = pParametro.Viaje_Permite_2_Conductores
    mConductorImporteTramoCompleto = -1
    mConductorImporteTramo1 = -1
    mConductorImporteTramo2 = -1
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

Public Function GetStatistics(ByRef CantidadLugares As Long, ByRef IndiceMinimo As Long, ByRef IndiceMaximo As Long) As Boolean
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Ruta_Statistics"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, mIDRuta)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    
    CantidadLugares = recData("CantidadLugares").Value
    IndiceMinimo = Val(recData("IndiceMinimo").Value & "")
    IndiceMaximo = Val(recData("IndiceMaximo").Value & "")
    
    recData.Close
    Set recData = Nothing
    
    Screen.MousePointer = vbDefault
    GetStatistics = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Ruta.GetStatistics", "Error al obtener las Estadísticas de la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta
End Function
