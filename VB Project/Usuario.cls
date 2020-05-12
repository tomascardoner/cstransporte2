VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES VARIABLES
Private mIDUsuario As Integer
Private mLoginName As String
Private mNombre As String
Private mPassword As String
Private mDescripcion As String
Private mIDUsuarioGrupo As Long
Private mIDEmpresa As Byte
Private mIDPersona As Long
Private mNotas As String
Private mActivo As Boolean
Private mSemaforo As Long
Private mSemaforoAnterior As Long
Private mFechaHoraCreacion As Date
Private mIDUsuarioCreacion As Integer
Private mFechaHoraModificacion As Date
Private mIDUsuarioModificacion As Integer

Private mIDCuentaCorrienteCaja As Long

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

Public Property Get IDUsuario() As Integer
    IDUsuario = mIDUsuario
End Property

Public Property Let IDUsuario(ByVal Value As Integer)
    If Value <> mIDUsuario Then
        mIsDirty = True
    End If
    mIDUsuario = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get LoginName() As String
    LoginName = mLoginName
End Property

Public Property Let LoginName(ByVal Value As String)
    If Value <> mLoginName Then
        mIsDirty = True
    End If
    mLoginName = Value
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

Public Property Get Password() As String
    Password = mPassword
End Property

Public Property Let Password(ByVal Value As String)
    If Value <> mPassword Then
        mIsDirty = True
    End If
    mPassword = Value
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

Public Property Get IDUsuarioGrupo() As Long
    IDUsuarioGrupo = mIDUsuarioGrupo
End Property

Public Property Let IDUsuarioGrupo(ByVal Value As Long)
    If Value <> mIDUsuarioGrupo Then
        mIsDirty = True
    End If
    mIDUsuarioGrupo = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get IDEmpresa() As Byte
    IDEmpresa = mIDEmpresa
End Property

Public Property Let IDEmpresa(ByVal Value As Byte)
    If Value <> mIDEmpresa Then
        mIsDirty = True
    End If
    mIDEmpresa = Value
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

Public Property Get Semaforo() As Long
    Semaforo = mSemaforo
End Property

Public Property Let Semaforo(ByVal Value As Long)
    If Value <> mSemaforo Then
        mIsDirty = True
    End If
    mSemaforo = Value
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

Public Property Get IDCuentaCorrienteCaja() As Long
    IDCuentaCorrienteCaja = mIDCuentaCorrienteCaja
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
    RefreshList_Module.RefreshList_RefreshUsuario IDUsuario
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
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_Usuario_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDUsuario", adSmallInt, adParamInput, , IDUsuario)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Usuario.OpenRecordset", "Error al abrir la tabla de Usuarios."
    If Not mrecData Is Nothing Then
        Set mrecData = Nothing
    End If
End Function

Public Function Requery() As Boolean
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

    mrecData.Requery
    
    Requery = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Usuario.Requery", "Error al releer los datos del Usuario." & vbCr & vbCr & "IDUsuario: " & IDUsuario
End Function

Public Function Load() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    If Not OpenRecordset() Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
        
        InitializeValues
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDUsuario = mrecData("IDUsuario").Value
        mLoginName = mrecData("LoginName").Value
        mNombre = mrecData("Nombre").Value
        mPassword = mrecData("Password").Value
        mDescripcion = mrecData("Descripcion").Value & ""
        mIDUsuarioGrupo = mrecData("IDUsuarioGrupo").Value
        mIDEmpresa = mrecData("IDEmpresa").Value
        mIDPersona = Val(mrecData("IDPersona").Value & "")
        mNotas = mrecData("Notas").Value & ""
        mActivo = mrecData("Activo").Value
        mSemaforo = mrecData("Semaforo").Value
        mSemaforoAnterior = mSemaforo
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Usuario.Load", "Error al obtener los datos del Usuario." & vbCr & vbCr & "IDUsuario: " & IDUsuario
End Function

Public Function OpenRecordsetByLoginName() As Boolean
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_Usuario_Data_ByLoginName"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("LoginName", adVarChar, adParamInput, 30, LoginName)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    OpenRecordsetByLoginName = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Usuario.OpenRecordset", "Error al abrir la tabla de Usuarios por medio de un Login Name."
    If Not mrecData Is Nothing Then
        Set mrecData = Nothing
    End If
End Function

Public Function LoadByLoginName() As Boolean
    Dim SaveLoginName As String
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    If Not OpenRecordsetByLoginName() Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass

    SaveLoginName = mLoginName
    Call InitializeValues
    mLoginName = SaveLoginName

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDUsuario = mrecData("IDUsuario").Value
        mLoginName = mrecData("LoginName").Value
        mNombre = mrecData("Nombre").Value
        mPassword = mrecData("Password").Value
        mDescripcion = mrecData("Descripcion").Value & ""
        mIDUsuarioGrupo = mrecData("IDUsuarioGrupo").Value
        mIDEmpresa = mrecData("IDEmpresa").Value
        mIDPersona = Val(mrecData("IDPersona").Value & "")
        mNotas = mrecData("Notas").Value & ""
        mActivo = mrecData("Activo").Value
        mSemaforo = mrecData("Semaforo").Value
        mSemaforoAnterior = mSemaforo
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    
    Screen.MousePointer = vbDefault
    LoadByLoginName = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Usuario.LoadByPersona", "Error al obtener los datos del Usuario a partir del Login Name." & vbCr & vbCr & "LoginName: " & LoginName
End Function

Public Function OpenRecordsetByPersona() As Boolean
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_Usuario_Data_ByPersona"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona", adInteger, adParamInput, , IDPersona)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    OpenRecordsetByPersona = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.Usuario.OpenRecordset", "Error al abrir la tabla de Usuarios por medio de una Persona."
    If Not mrecData Is Nothing Then
        Set mrecData = Nothing
    End If
End Function

Public Function LoadByPersona() As Boolean
    Dim SaveIDPersona As Long
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    If Not OpenRecordsetByPersona() Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass

    SaveIDPersona = mIDPersona
    Call InitializeValues
    mIDPersona = SaveIDPersona

    If (Not mNoMatchRaiseError) And mrecData.EOF Then
        mNoMatch = True
    Else
        mNoMatch = False
        mIsNew = False
    
        mIDUsuario = mrecData("IDUsuario").Value
        mLoginName = mrecData("LoginName").Value
        mNombre = mrecData("Nombre").Value
        mPassword = mrecData("Password").Value
        mDescripcion = mrecData("Descripcion").Value & ""
        mIDUsuarioGrupo = mrecData("IDUsuarioGrupo").Value
        mIDEmpresa = mrecData("IDEmpresa").Value
        mIDPersona = Val(mrecData("IDPersona").Value & "")
        mNotas = mrecData("Notas").Value & ""
        mActivo = mrecData("Activo").Value
        mSemaforo = mrecData("Semaforo").Value
        mSemaforoAnterior = mSemaforo
        mFechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        mIDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        mFechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        mIDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    
    Screen.MousePointer = vbDefault
    LoadByPersona = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Usuario.LoadByPersona", "Error al obtener los datos del Usuario a partir de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona
End Function

Public Function CambioSemaforo() As Boolean
    Requery
    mSemaforo = mrecData("Semaforo").Value
    If mSemaforoAnterior = mSemaforo Then
        CambioSemaforo = False
    Else
        mSemaforoAnterior = mSemaforo
        CambioSemaforo = True
    End If
End Function

Public Function Copy() As Boolean
    If Load() Then
        Copy = True

        mIDUsuario = 0
        mLoginName = ""

        mIsNew = True
        mIsCopy = True
        mIsDirty = True
    End If
End Function

Public Function Update() As Boolean
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
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
        Set cmdData = New ADODB.command
        Set cmdData.ActiveConnection = pDatabase.Connection
        cmdData.CommandText = "usp_Usuario_IDMax"
        cmdData.CommandType = adCmdStoredProc
        Set recData = New ADODB.Recordset
        recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
        Set cmdData = Nothing
        mIDUsuario = Val(recData("IDUsuarioMax").Value & "") + 1
        recData.Close
        Set recData = Nothing
    
        mrecData.AddNew
        mrecData("IDUsuario").Value = mIDUsuario
        mrecData("FechaHoraCreacion").Value = Now
        mrecData("IDUsuarioCreacion").Value = pUsuario.IDUsuario
    End If
    mrecData("LoginName").Value = mLoginName
    mrecData("Nombre").Value = mNombre
    mrecData("Password").Value = mPassword
    mrecData("Descripcion").Value = IIf(Trim(mDescripcion) = "", Null, mDescripcion)
    mrecData("IDUsuarioGrupo").Value = mIDUsuarioGrupo
    mrecData("IDEmpresa").Value = mIDEmpresa
    mrecData("IDPersona").Value = IIf(mIDPersona = 0, Null, mIDPersona)
    mrecData("Notas").Value = IIf(Trim(mNotas) = "", Null, mNotas)
    mrecData("Activo").Value = mActivo
    mrecData("Semaforo").Value = mSemaforo
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
            MsgBox "Ya existe un Usuario con el mismo nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Usuario.Update", "Error al actualizar el Usuario." & vbCr & vbCr & "IDUsuario: " & IDUsuario
        End If
        If mrecData.EditMode = adEditAdd Or mrecData.EditMode = adEditInProgress Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.Usuario.Update", "Error al actualizar el Usuario." & vbCr & vbCr & "IDUsuario: " & IDUsuario
    End If
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
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
            MsgBox "No se puede eliminar el Usuario debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Usuario.Delete", "Error al eliminar el Usuario." & vbCr & vbCr & "IDUsuario: " & IDUsuario
        End If
        If mrecData.EditMode = adEditDelete Then
            mrecData.CancelUpdate
        End If
    Else
        ShowErrorMessage "Classes.Usuario.Delete", "Error al eliminar el Usuario." & vbCr & vbCr & "IDUsuario: " & IDUsuario
    End If
End Function

Public Sub InitializeValues()
    mLoginName = ""
    mNombre = ""
    mPassword = ""
    mDescripcion = ""
    mIDUsuarioGrupo = 0
    mIDEmpresa = 0
    mIDPersona = 0
    mNotas = ""
    mActivo = True
    mSemaforo = 0
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

Public Function LogIn() As Boolean
    Dim VehiculoMantenimiento As VehiculoMantenimiento
    Dim PersonaAlarma As PersonaAlarma
    Dim Alarma As Alarma
    Dim Avisos As Long
    Dim CuentaCorrienteCaja As CuentaCorrienteCaja
    
    Dim Reporte As Reporte
    
    WriteLogEvent "User Login Succesfully - LoginName: " & LoginName, vbLogEventTypeInformation, pParametro.LogAccion_Enabled
    Call CSM_Registry.SetValue_ToApplication_CurrentUser("", "LastUserID", LoginName)
    frmMDI.stbMain.Panels("USUARIO").Text = Nombre
    
    Set pCPermiso = New CPermiso
    pCPermiso.IDUsuarioGrupo = IDUsuarioGrupo
    pCPermiso.OpenRecordset
    pCPermiso.Load

    '//////////////////////////////////////////////////////////////////
    'OBTENGO LOS PARAMETROS DEL USUARIO
    If Not pParametro.InitializeUserParameters Then
        Exit Function
    End If
    If Not pParametro.LoadUserParameters() Then
        Exit Function
    End If
    pMessengerEnabled = (pParametro.Messenger_Enabled And pParametro.Usuario_Messenger_Enabled And pCPermiso.GotPermission(PERMISO_MESSENGER, False))
    
    '//////////////////////////////////////////////////////////////////
    'OBTENGO LA CAJA DE CUENTA CORRIENTE
    WriteLogEvent "Reading Cuenta Corriente Caja of User: INIT", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
    Set CuentaCorrienteCaja = New CuentaCorrienteCaja
    CuentaCorrienteCaja.IDPersona = pUsuario.IDPersona
    CuentaCorrienteCaja.NoMatchRaiseError = False
    CuentaCorrienteCaja.LoadByPersona
    mIDCuentaCorrienteCaja = CuentaCorrienteCaja.IDCuentaCorrienteCaja
    Set CuentaCorrienteCaja = Nothing
    WriteLogEvent "Reading Cuenta Corriente Caja of User: DONE", vbLogEventTypeInformation, pParametro.LogAccion_Enabled

    '//////////////////////////////////////////////////////////////////
    'SINCRONIZO LA HORA CON EL SERVER
    ' Updated: 2018-04-14 - La aplicación se ejecuta en el servidor, por lo tanto, esto no tiene sentido
    If pParametro.System_SyncDateTimeWithDBServer_Enabled And pParametro.Workstation_SyncDateTimeWithDBServer_Enabled Then
        Call CSM_Time.SyncronizeWithSQLServer
    End If
    
    '//////////////////////////////////////////////////////////////////
    'SHOW/HIDE MENU ITEMS AND BUTTONS
    WriteLogEvent "Showing/Hiding Utility menu: INIT", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
    frmMDI.mnuUtility.Visible = pCPermiso.GotPermission(PERMISO_UTILIDAD, False)
    frmMDI.mnuUtilityParameter.Enabled = pCPermiso.GotPermission(PERMISO_UTILIDAD_EDITOR_DIRECTO_PARAMETRO, False)
    frmMDI.mnuUtilityCallerIDSimulate.Enabled = (IDUsuario = 0)
    frmMDI.mnuUtilityExecute.Enabled = (IDUsuario = USUARIO_ID_ADMINISTRATOR)
    frmMDI.mnuUtilityHistory.Enabled = (IDUsuario = USUARIO_ID_ADMINISTRATOR)
    frmMDI.mnuUtilityUpdatePrecio.Enabled = pCPermiso.GotPermission(PERMISO_UTILIDAD_ACTUALIZAR_PRECIO_RESERVA, False)
    frmMDI.mnuUtilityUpdateSueldo.Enabled = pCPermiso.GotPermission(PERMISO_UTILIDAD_ACTUALIZAR_SUELDO_VIAJE, False)
    frmMDI.tlbMain.Buttons("MESSENGER").Visible = pMessengerEnabled
    frmMDI.tlbMain.Buttons("SEPARATOR_7").Visible = pMessengerEnabled
    WriteLogEvent "Showing/Hiding Utility menu: DONE", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
    
    LogIn = True
    
    'AVISOS DE MANTENIMIENTO DE VEHICULOS
    If pParametro.VehiculoMantenimiento_Aviso Then
        WriteLogEvent "Verifying Mantenimiento de Vehículos Alerts: INIT", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
        Set VehiculoMantenimiento = New VehiculoMantenimiento
        Avisos = VehiculoMantenimiento.Check
        Set VehiculoMantenimiento = Nothing
        
        If Avisos > 0 Then
            If MsgBox("Hay " & Avisos & " Aviso" & IIf(Avisos = 1, "", "s") & " de Mantenimiento de Vehículos." & vbCr & vbCr & "¿Desea ver el Reporte?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Vehiculo_Mantenimiento_Aviso") Then
                    Set Reporte = New Reporte
                    Reporte.IDReporte = "Vehiculo_Mantenimiento_Aviso"
                    If Reporte.Load() Then
                        If Reporte.OpenReport() Then
                            Reporte.PrintReport True
                        End If
                    End If
                    Set Reporte = Nothing
                End If
            End If
        End If
        WriteLogEvent "Verifying Mantenimiento de Vehículos Alerts: DONE", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
    End If

    'AVISOS DE ALARMAS DE PERSONAS
    If pParametro.PersonaAlarma_Aviso Then
        WriteLogEvent "Verifying Alarmas de Personas Alerts: INIT", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
        Set PersonaAlarma = New PersonaAlarma
        Avisos = PersonaAlarma.Check
        Set PersonaAlarma = Nothing
        
        If Avisos > 0 Then
            If MsgBox(IIf(Avisos = 1, "Hay un Aviso de Alarma de una Persona.", "Hay " & Avisos & " de Alarmas de Personas.") & vbCr & vbCr & "¿Desea ver el Reporte?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Persona_Alarma_Aviso") Then
                    Set Reporte = New Reporte
                    Reporte.IDReporte = "Persona_Alarma_Aviso"
                    If Reporte.Load() Then
                        If Reporte.OpenReport() Then
                            Reporte.PrintReport True
                        End If
                    End If
                    Set Reporte = Nothing
                End If
            End If
        End If
        WriteLogEvent "Verifying Alarmas de Personas Alerts: DONE", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
    End If

    'AVISOS DE ALARMAS GENERALES
    If pParametro.Alarma_Aviso Then
        WriteLogEvent "Verifying Alarmas Generales Alerts: INIT", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
        Set Alarma = New Alarma
        Avisos = Alarma.Check
        Set Alarma = Nothing
        
        If Avisos > 0 Then
            If MsgBox(IIf(Avisos = 1, "Hay un Aviso de Alarma General.", "Hay " & Avisos & " de Alarmas Generales.") & vbCr & vbCr & "¿Desea ver el Reporte?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Alarma_Aviso") Then
                    Set Reporte = New Reporte
                    Reporte.IDReporte = "Alarma_Aviso"
                    If Reporte.Load() Then
                        If Reporte.OpenReport() Then
                            Reporte.PrintReport True
                        End If
                    End If
                    Set Reporte = Nothing
                End If
            End If
        End If
        WriteLogEvent "Verifying Alarmas Generales Alerts: DONE", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
    End If
    
    'AVISO DE VIAJES CON ESTADO VENCIDO
    If pParametro.Viaje_EstadoVencido_Check Then
        WriteLogEvent "Verifying Viaje Estado Vencido Alerts: INIT", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
        Call Viaje_EstadoVencido_List
        WriteLogEvent "Verifying Viaje Estado Vencido Alerts: DONE", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
    End If
    
    'MENSAJES GENERALES
    If frmMensaje.CheckMessages() Then
        frmMensaje.Show vbModal, frmMDI
    Else
        Unload frmMensaje
    End If
End Function

Public Sub LogOut()
    CSM_Forms.UnloadAll "frmMDI"
    frmMDI.stbMain.Panels("USUARIO").Text = ""
    WriteLogEvent "User Logout Succesfully - Username: " & IDUsuario, vbLogEventTypeInformation, pParametro.LogAccion_Enabled
End Sub
