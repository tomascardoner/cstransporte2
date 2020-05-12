VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RutaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IDRuta As String
Public IDLugar As Long
Public Indice As Long
Public IDLugarGrupo As Long
Public Kilometro As Integer
Public Duracion As Integer
Public Espera As Integer
Public HoraInicio As Date
Public HoraFin As Date
Public FechaHoraCreacion As Date
Public IDUsuarioCreacion As Integer
Public FechaHoraModificacion As Date
Public IDUsuarioModificacion As Integer

Public NoMatchRaiseError As Boolean
Public NoMatch As Boolean

Public RefreshList As Boolean

Private mrecData As ADODB.Recordset

Private Function OpenRecordset() As Boolean
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_RutaDetalle_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta", adChar, adParamInput, 20, IDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("IDLugar", adInteger, adParamInput, , IDLugar)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing
    
    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.RutaDetalle.OpenRecordset", "Error al abrir la tabla de Detalles de Rutas."
End Function

Public Function Load() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    If Not OpenRecordset() Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If (Not NoMatchRaiseError) And mrecData.EOF Then
        NoMatch = True
        
        InitializeValues
    Else
        NoMatch = False
        
        Indice = mrecData("Indice").Value
        IDLugarGrupo = mrecData("IDLugarGrupo").Value
        Kilometro = IIf(IsNull(mrecData("Kilometro").Value), -1, mrecData("Kilometro").Value)
        Duracion = IIf(IsNull(mrecData("Duracion").Value), -1, mrecData("Duracion").Value)
        Espera = IIf(IsNull(mrecData("Espera").Value), -1, mrecData("Espera").Value)
        HoraInicio = IIf(IsNull(mrecData("HoraInicio").Value), DATE_TIME_FIELD_NULL_VALUE, mrecData("HoraInicio").Value)
        HoraFin = IIf(IsNull(mrecData("HoraFin").Value), DATE_TIME_FIELD_NULL_VALUE, mrecData("HoraFin").Value)
        FechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        IDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        FechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        IDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.RutaDetalle.Load", "Error al obtener los datos del Detalle de la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta & vbCr & "IDLugar: " & IDLugar
End Function

Public Function AddNew() As Boolean
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If mrecData Is Nothing Then
        If Not OpenRecordset() Then
            Exit Function
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_RutaDetalle_IndiceMax"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta", adChar, adParamInput, 20, IDRuta)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    Indice = Val(recData("IndiceMax").Value & "") + 10
    recData.Close
    Set recData = Nothing
    
    mrecData.AddNew
    mrecData("IDRuta").Value = IDRuta
    mrecData("IDLugar").Value = IDLugar
    mrecData("Indice").Value = Indice
    mrecData("IDLugarGrupo").Value = IDLugarGrupo
    mrecData("Kilometro").Value = IIf(Kilometro = -1, Null, Kilometro)
    mrecData("Duracion").Value = IIf(Duracion = -1, Null, Duracion)
    mrecData("Espera").Value = IIf(Espera = -1, Null, Espera)
    mrecData("HoraInicio").Value = IIf(HoraInicio = DATE_TIME_FIELD_NULL_VALUE, Null, HoraInicio)
    mrecData("HoraFin").Value = IIf(HoraFin = DATE_TIME_FIELD_NULL_VALUE, Null, HoraFin)
    mrecData("FechaHoraCreacion").Value = Now
    mrecData("IDUsuarioCreacion").Value = pUsuario.IDUsuario
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If RefreshList Then
        RefreshList_RefreshRutaDetalle IDRuta, IDLugar
    End If
    
    Screen.MousePointer = vbDefault
    AddNew = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe el Lugar en este Detalle de Ruta.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.RutaDetalle.AddNew", "Error al Agregar el Detalle de la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta & vbCr & "IDLugar: " & IDLugar
        End If
    Else
        ShowErrorMessage "Classes.RutaDetalle.AddNew", "Error al Agregar el Detalle de la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta & vbCr & "IDLugar: " & IDLugar
    End If
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Public Function Update() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    mrecData("IDRuta").Value = IDRuta
    mrecData("IDLugar").Value = IDLugar
    mrecData("Indice").Value = Indice
    mrecData("IDLugarGrupo").Value = IDLugarGrupo
    mrecData("Kilometro").Value = IIf(Kilometro = -1, Null, Kilometro)
    mrecData("Duracion").Value = IIf(Duracion = -1, Null, Duracion)
    mrecData("Espera").Value = IIf(Espera = -1, Null, Espera)
    mrecData("HoraInicio").Value = IIf(HoraInicio = DATE_TIME_FIELD_NULL_VALUE, Null, CStr(HoraInicio))
    mrecData("HoraFin").Value = IIf(HoraFin = DATE_TIME_FIELD_NULL_VALUE, Null, CStr(HoraFin))
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If RefreshList Then
        RefreshList_RefreshRutaDetalle IDRuta, IDLugar
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe el Lugar en este Detalle de Ruta.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.RutaDetalle.Update", "Error al Actualizar el Detalle de la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta & vbCr & "IDLugar: " & IDLugar
        End If
    Else
        ShowErrorMessage "Classes.RutaDetalle.Update", "Error al Actualizar el Detalle de la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta & vbCr & "IDLugar: " & IDLugar
    End If
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Public Function Delete() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    mrecData.Delete
    
    Call ReIndex(True)
        
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede Eliminar el Detalle de la Ruta debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.RutaDetalle.Delete", "Error al Eliminar el Detalle de la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta & vbCr & "IDLugar: " & IDLugar
        End If
    Else
        ShowErrorMessage "Classes.RutaDetalle.Delete", "Error al Eliminar el Detalle de la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta & vbCr & "IDLugar: " & IDLugar
    End If
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Private Sub InitializeValues()
    Indice = 0
    IDLugarGrupo = 0
    Kilometro = -1
    Duracion = -1
    Espera = -1
    HoraInicio = DATE_TIME_FIELD_NULL_VALUE
    HoraFin = DATE_TIME_FIELD_NULL_VALUE
    FechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    IDUsuarioCreacion = 0
    FechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    IDUsuarioModificacion = 0
End Sub

Private Sub Class_Initialize()
    InitializeValues
    NoMatchRaiseError = True
    RefreshList = True
End Sub

Private Sub Class_Terminate()
    If Not mrecData Is Nothing Then
        If mrecData.State = adStateOpen Then
            If Not (mrecData.BOF Or mrecData.EOF) Then
                If mrecData.EditMode <> adEditNone Then
                    mrecData.CancelUpdate
                End If
            End If
            mrecData.Close
        End If
        Set mrecData = Nothing
    End If
End Sub

Public Sub ReIndex(ByVal Ascending As Boolean)
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    Dim IndiceCurrent As Long
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_RutaDetalle_List"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, IDRuta)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenKeyset, adLockOptimistic
    Set cmdData = Nothing
    
    If Not recData.EOF Then
        pDatabase.Connection.BeginTrans
        If Ascending Then
            IndiceCurrent = 10
            Do While Not recData.EOF
                recData("Indice").Value = IndiceCurrent
                recData.Update
                recData.MoveNext
                IndiceCurrent = IndiceCurrent + 10
            Loop
        Else
            recData.MoveLast
            IndiceCurrent = recData.RecordCount * 10
            Do While Not recData.BOF
                recData("Indice").Value = IndiceCurrent
                recData.Update
                recData.MovePrevious
                IndiceCurrent = IndiceCurrent - 10
            Loop
        End If
        pDatabase.Connection.CommitTrans
    End If
    recData.Close
    Set recData = Nothing
    
    If RefreshList Then
        RefreshList_RefreshRutaDetalle IDRuta, IDLugar
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.RutaDetalle.ReIndex", "Error al Reindexar los Detalles de la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta
    recData.CancelUpdate
    pDatabase.Connection.RollbackTrans
End Sub
