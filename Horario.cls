VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Horario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mDiaSemana As Integer
Private mHora As Date
Private mIDRuta As String
Private mIDConductor As Long
Private mIDConductor2 As Long
Private mIDVehiculo As Long
Private mConductorImporteTramoCompleto As Currency
Private mConductorImporteTramo1 As Currency
Private mConductorImporteTramo2 As Currency
Private mNotas As String
Private mActivo As Boolean
Private mPersonal As Boolean
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

Public Property Get DiaSemana() As Integer
    DiaSemana = mDiaSemana
End Property

Public Property Let DiaSemana(ByVal Value As Integer)
    If Value <> mDiaSemana Then
        mIsDirty = True
    End If
    mDiaSemana = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Hora() As Date
    Hora = mHora
End Property

Public Property Get Hora_Formatted() As String
    Hora_Formatted = Format(mHora, "Short Time")
End Property

Public Property Let Hora(ByVal Value As Date)
    If Value <> mHora Then
        mIsDirty = True
    End If
    mHora = Value
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

Public Property Get IDConductor() As Long
    IDConductor = mIDConductor
End Property

Public Property Let IDConductor(ByVal Value As Long)
    If Value <> mIDConductor Then
        mIsDirty = True
    End If
    mIDConductor = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDConductor2() As Long
    IDConductor2 = mIDConductor2
End Property

Public Property Let IDConductor2(ByVal Value As Long)
    If Value <> mIDConductor2 Then
        mIsDirty = True
    End If
    mIDConductor2 = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDVehiculo() As Long
    IDVehiculo = mIDVehiculo
End Property

Public Property Let IDVehiculo(ByVal Value As Long)
    If Value <> mIDVehiculo Then
        mIsDirty = True
    End If
    mIDVehiculo = Value
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

Public Property Get Personal() As Boolean
    Personal = mPersonal
End Property

Public Property Let Personal(ByVal Value As Boolean)
    If Value <> mPersonal Then
        mIsDirty = True
    End If
    mPersonal = Value
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
    RefreshList_Module.RefreshList_RefreshHorario DiaSemana, Hora, IDRuta
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
    mrecData.Source = "SELECT * FROM Horario"
    mrecData.Open , , , , adCmdText
    mrecData.Properties("Update Criteria").Value = adCriteriaKey

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Horario.OpenRecordset", "Error al abrir la tabla de Horario."
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

    mrecData.Filter = "DiaSemana = " & mDiaSemana & " AND Hora = #" & DATE_TIME_FIELD_NULL_VALUE & " " & mHora & "# AND IDRuta = '" & mIDRuta & "'"

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mDiaSemana = mrecData("DiaSemana").Value
        mHora = Format(mrecData("Hora").Value, "Long Time")
        mIDRuta = RTrim(mrecData("IDRuta").Value)
        mIDConductor = Val(mrecData("IDConductor").Value & "")
        mIDConductor2 = Val(mrecData("IDConductor2").Value & "")
        mIDVehiculo = Val(mrecData("IDVehiculo").Value & "")
        mConductorImporteTramoCompleto = IIf(IsNull(mrecData("ConductorImporteTramoCompleto").Value), -1, mrecData("ConductorImporteTramoCompleto").Value)
        mConductorImporteTramo1 = IIf(IsNull(mrecData("ConductorImporteTramo1").Value), -1, mrecData("ConductorImporteTramo1").Value)
        mConductorImporteTramo2 = IIf(IsNull(mrecData("ConductorImporteTramo2").Value), -1, mrecData("ConductorImporteTramo2").Value)
        mNotas = mrecData("Notas").Value & ""
        mActivo = mrecData("Activo").Value
        mPersonal = mrecData("Personal").Value
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Horario.Load", "Error al obtener los datos del Horario." & vbCr & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora & vbCr & "IDRuta: " & IDRuta
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mDiaSemana = 0
        mHora = DATE_TIME_FIELD_NULL_VALUE
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
    End If
    mrecData("DiaSemana").Value = mDiaSemana
    mrecData("Hora").Value = DATE_TIME_FIELD_NULL_VALUE & " " & mHora
    mrecData("IDRuta").Value = mIDRuta
    mrecData("IDConductor").Value = IIf(mIDConductor = 0, Null, mIDConductor)
    mrecData("IDConductor2").Value = IIf(mIDConductor2 = 0, Null, mIDConductor2)
    mrecData("IDVehiculo").Value = IIf(mIDVehiculo = 0, Null, mIDVehiculo)
    mrecData("ConductorImporteTramoCompleto").Value = IIf(mConductorImporteTramoCompleto = -1, Null, mConductorImporteTramoCompleto)
    mrecData("ConductorImporteTramo1").Value = IIf(mConductorImporteTramo1 = -1, Null, mConductorImporteTramo1)
    mrecData("ConductorImporteTramo2").Value = IIf(mConductorImporteTramo2 = -1, Null, mConductorImporteTramo2)
    mrecData("Notas").Value = IIf(Trim(mNotas) = "", Null, mNotas)
    mrecData("Activo").Value = mActivo
    mrecData("Personal").Value = mPersonal
    If mIsNew Then
        mrecData("FechaHoraCreacion").Value = mFechaHoraCreacion
        mrecData("IDUsuarioCreacion").Value = pUsuario.IDUsuario
    End If
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
            MsgBox "Ya existe un Horario con el mismo nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Horario.Update", "Error al actualizar el Horario." & vbCr & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora & vbCr & "IDRuta: " & IDRuta
        End If
        If mrecData.EditMode = adEditAdd Or mrecData.EditMode = adEditInProgress Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.Horario.Update", "Error al actualizar el Horario." & vbCr & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora & vbCr & "IDRuta: " & IDRuta
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
            MsgBox "No se puede eliminar el Horario debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Horario.Delete", "Error al eliminar el Horario." & vbCr & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora & vbCr & "IDRuta: " & IDRuta
        End If
        If mrecData.EditMode = adEditDelete Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.Horario.Delete", "Error al eliminar el Horario." & vbCr & vbCr & "DiaSemana: " & DiaSemana & vbCr & "Hora: " & Hora & vbCr & "IDRuta: " & IDRuta
    End If
End Function

Private Sub InitializeValues()
    mIDConductor = 0
    mIDConductor2 = 0
    mIDVehiculo = 0
    mConductorImporteTramoCompleto = -1
    mConductorImporteTramo1 = -1
    mConductorImporteTramo2 = -1
    mNotas = ""
    mActivo = True
    mPersonal = False
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
