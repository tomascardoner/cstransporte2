VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VehiculoMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IDVehiculo As Long
Public IDVehiculoMantenimientoGrupo As Long
Public Tipo As String
Public KilometrajeLapso As Long
Public KilometrajePreaviso As Long
Public DiasLapso As Integer
Public DiasPreaviso As Integer
Public FechaFecha As Date
Public FechaPreaviso As Integer
Public Notas As String
Public Activo As Boolean
Public FechaHoraCreacion As Date
Public IDUsuarioCreacion As Integer
Public FechaHoraModificacion As Date
Public IDUsuarioModificacion As Integer

Public NoMatchRaiseError As Boolean
Public NoMatch As Boolean

Public RefreshList As Boolean

Private mrecData As ADODB.Recordset

Public Function Tipo_ToString() As String
    Select Case Tipo
        Case VEHICULO_MATENIMIENTO_TIPO_KILOMETRAJE
            Tipo_ToString = VEHICULO_MATENIMIENTO_TIPO_KILOMETRAJE_NOMBRE
        Case VEHICULO_MATENIMIENTO_TIPO_DIAS
            Tipo_ToString = VEHICULO_MATENIMIENTO_TIPO_DIAS_NOMBRE
        Case VEHICULO_MATENIMIENTO_TIPO_FECHA
            Tipo_ToString = VEHICULO_MATENIMIENTO_TIPO_FECHA_NOMBRE
        Case VEHICULO_MATENIMIENTO_TIPO_NINGUNO
            Tipo_ToString = VEHICULO_MATENIMIENTO_TIPO_NINGUNO_NOMBRE
        Case Else
            Tipo_ToString = ""
    End Select
End Function

Private Function OpenRecordset() As Boolean
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_VehiculoMantenimiento_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDVehiculo_FILTER", adInteger, adParamInput, , IDVehiculo)
    cmdData.Parameters.Append cmdData.CreateParameter("IDVehiculoMantenimientoGrupo_FILTER", adInteger, adParamInput, , IDVehiculoMantenimientoGrupo)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.VehiculoMantenimiento.OpenRecordset", "Error al abrir la tabla de Mantenimiento de Vehículos."
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
        
        Tipo = mrecData("Tipo").Value
        Select Case Tipo
            Case VEHICULO_MATENIMIENTO_TIPO_KILOMETRAJE
                KilometrajeLapso = mrecData("KilometrajeLapso").Value
                KilometrajePreaviso = mrecData("KilometrajePreaviso").Value
                DiasLapso = 0
                DiasPreaviso = 0
                FechaFecha = DATE_TIME_FIELD_NULL_VALUE
                FechaPreaviso = 0
            Case VEHICULO_MATENIMIENTO_TIPO_DIAS
                KilometrajeLapso = 0
                KilometrajePreaviso = 0
                DiasLapso = mrecData("DiasLapso").Value
                DiasPreaviso = mrecData("DiasPreaviso").Value
                FechaFecha = DATE_TIME_FIELD_NULL_VALUE
                FechaPreaviso = 0
            Case VEHICULO_MATENIMIENTO_TIPO_FECHA
                KilometrajeLapso = 0
                KilometrajePreaviso = 0
                DiasLapso = 0
                DiasPreaviso = 0
                FechaFecha = mrecData("FechaFecha").Value
                FechaPreaviso = mrecData("FechaPreaviso").Value
            Case VEHICULO_MATENIMIENTO_TIPO_NINGUNO
                KilometrajeLapso = 0
                KilometrajePreaviso = 0
                DiasLapso = 0
                DiasPreaviso = 0
                FechaFecha = DATE_TIME_FIELD_NULL_VALUE
                FechaPreaviso = 0
        End Select
        Notas = mrecData("Notas").Value & ""
        Activo = mrecData("Activo").Value
        FechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        IDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        FechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        IDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.VehiculoMantenimiento.Load", "Error al obtener los datos del Mantenimiento del Vehículo." & vbCr & vbCr & "IDVehiculo: " & IDVehiculo & vbCr & "IDVehiculoMantenimientoGrupo: " & IDVehiculoMantenimientoGrupo
End Function

Public Function AddNew() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If mrecData Is Nothing Then
        If Not OpenRecordset() Then
            Exit Function
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    mrecData.AddNew
    mrecData("IDVehiculo").Value = IDVehiculo
    mrecData("IDVehiculoMantenimientoGrupo").Value = IDVehiculoMantenimientoGrupo
    mrecData("Tipo").Value = Tipo
    Select Case Tipo
        Case VEHICULO_MATENIMIENTO_TIPO_KILOMETRAJE
            mrecData("KilometrajeLapso").Value = KilometrajeLapso
            mrecData("KilometrajePreaviso").Value = KilometrajePreaviso
            mrecData("DiasLapso").Value = Null
            mrecData("DiasPreaviso").Value = Null
            mrecData("FechaFecha").Value = Null
            mrecData("FechaPreaviso").Value = Null
        Case VEHICULO_MATENIMIENTO_TIPO_DIAS
            mrecData("KilometrajeLapso").Value = Null
            mrecData("KilometrajePreaviso").Value = Null
            mrecData("DiasLapso").Value = DiasLapso
            mrecData("DiasPreaviso").Value = DiasPreaviso
            mrecData("FechaFecha").Value = Null
            mrecData("FechaPreaviso").Value = Null
        Case VEHICULO_MATENIMIENTO_TIPO_FECHA
            mrecData("KilometrajeLapso").Value = Null
            mrecData("KilometrajePreaviso").Value = Null
            mrecData("DiasLapso").Value = Null
            mrecData("DiasPreaviso").Value = Null
            mrecData("FechaFecha").Value = FechaFecha
            mrecData("FechaPreaviso").Value = FechaPreaviso
        Case VEHICULO_MATENIMIENTO_TIPO_NINGUNO
            mrecData("KilometrajeLapso").Value = Null
            mrecData("KilometrajePreaviso").Value = Null
            mrecData("DiasLapso").Value = Null
            mrecData("DiasPreaviso").Value = Null
            mrecData("FechaFecha").Value = Null
            mrecData("FechaPreaviso").Value = Null
    End Select
    mrecData("Notas").Value = IIf(Trim(Notas) = "", Null, Notas)
    mrecData("Activo").Value = Activo
    mrecData("FechaHoraCreacion").Value = Now
    mrecData("IDUsuarioCreacion").Value = pUsuario.IDUsuario
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If RefreshList Then
        RefreshList_RefreshVehiculoMantenimiento IDVehiculo, IDVehiculoMantenimientoGrupo
    End If
    
    Screen.MousePointer = vbDefault
    AddNew = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Or pDatabase.Connection.Errors(0).NativeError = pDatabase.ERRORNATIVE_DUPLICATEKEY_PRIMARYKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe el Grupo para este Vehículo.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.VehiculoMantenimiento.AddNew", "Error al Agregar el Mantenimiento del Vehículo." & vbCr & vbCr & "IDVehiculo: " & IDVehiculo & vbCr & "IDVehiculoMantenimientoGrupo: " & IDVehiculoMantenimientoGrupo
        End If
    Else
        ShowErrorMessage "Classes.VehiculoMantenimiento.AddNew", "Error al Agregar el Mantenimiento del Vehículo." & vbCr & vbCr & "IDVehiculo: " & IDVehiculo & vbCr & "IDVehiculoMantenimientoGrupo: " & IDVehiculoMantenimientoGrupo
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
    
    If mrecData("IDVehiculo").Value <> IDVehiculo Then
        mrecData("IDVehiculo").Value = IDVehiculo
    End If
    If mrecData("IDVehiculoMantenimientoGrupo").Value <> IDVehiculoMantenimientoGrupo Then
        mrecData("IDVehiculoMantenimientoGrupo").Value = IDVehiculoMantenimientoGrupo
    End If
    mrecData("Tipo").Value = Tipo
    Select Case Tipo
        Case VEHICULO_MATENIMIENTO_TIPO_KILOMETRAJE
            mrecData("KilometrajeLapso").Value = KilometrajeLapso
            mrecData("KilometrajePreaviso").Value = KilometrajePreaviso
            mrecData("DiasLapso").Value = Null
            mrecData("DiasPreaviso").Value = Null
            mrecData("FechaFecha").Value = Null
            mrecData("FechaPreaviso").Value = Null
        Case VEHICULO_MATENIMIENTO_TIPO_DIAS
            mrecData("KilometrajeLapso").Value = Null
            mrecData("KilometrajePreaviso").Value = Null
            mrecData("DiasLapso").Value = DiasLapso
            mrecData("DiasPreaviso").Value = DiasPreaviso
            mrecData("FechaFecha").Value = Null
            mrecData("FechaPreaviso").Value = Null
        Case VEHICULO_MATENIMIENTO_TIPO_FECHA
            mrecData("KilometrajeLapso").Value = Null
            mrecData("KilometrajePreaviso").Value = Null
            mrecData("DiasLapso").Value = Null
            mrecData("DiasPreaviso").Value = Null
            mrecData("FechaFecha").Value = FechaFecha
            mrecData("FechaPreaviso").Value = FechaPreaviso
        Case VEHICULO_MATENIMIENTO_TIPO_NINGUNO
            mrecData("KilometrajeLapso").Value = Null
            mrecData("KilometrajePreaviso").Value = Null
            mrecData("DiasLapso").Value = Null
            mrecData("DiasPreaviso").Value = Null
            mrecData("FechaFecha").Value = Null
            mrecData("FechaPreaviso").Value = Null
    End Select
    mrecData("Notas").Value = IIf(Trim(Notas) = "", Null, Notas)
    mrecData("Activo").Value = Activo
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If RefreshList Then
        RefreshList_RefreshVehiculoMantenimiento IDVehiculo, IDVehiculoMantenimientoGrupo
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe el Grupo para este Vehículo.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.VehiculoMantenimiento.Update", "Error al Actualizar el Mantenimiento del Vehículo." & vbCr & vbCr & "IDVehiculo: " & IDVehiculo & vbCr & "IDVehiculoMantenimientoGrupo: " & IDVehiculoMantenimientoGrupo
        End If
    Else
        ShowErrorMessage "Classes.VehiculoMantenimiento.Update", "Error al Actualizar el Mantenimiento del Vehículo." & vbCr & vbCr & "IDVehiculo: " & IDVehiculo & vbCr & "IDVehiculoMantenimientoGrupo: " & IDVehiculoMantenimientoGrupo
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
    
    If RefreshList Then
        RefreshList_RefreshVehiculoMantenimiento 0, 0
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede Eliminar el Mantenimiento del Vehículo debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.VehiculoMantenimiento.Delete", "Error al Eliminar el Mantenimiento del Vehículo." & vbCr & vbCr & "IDVehiculo: " & IDVehiculo & vbCr & "IDVehiculoMantenimientoGrupo: " & IDVehiculoMantenimientoGrupo
        End If
    Else
        ShowErrorMessage "Classes.VehiculoMantenimiento.Delete", "Error al Eliminar el Mantenimiento del Vehículo." & vbCr & vbCr & "IDVehiculo: " & IDVehiculo & vbCr & "IDVehiculoMantenimientoGrupo: " & IDVehiculoMantenimientoGrupo
    End If
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Private Sub InitializeValues()
    Tipo = ""
    KilometrajeLapso = 0
    KilometrajePreaviso = 0
    DiasLapso = 0
    DiasPreaviso = 0
    FechaFecha = DATE_TIME_FIELD_NULL_VALUE
    FechaPreaviso = 0
    Notas = ""
    Activo = True
    FechaHoraCreacion = DATE_TIME_FIELD_NULL_VALUE
    IDUsuarioCreacion = 0
    FechaHoraModificacion = DATE_TIME_FIELD_NULL_VALUE
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

Public Function Check() As Long
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Report_Vehiculo_Mantenimiento_Aviso_List"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDVehiculo_FILTER", adInteger, adParamInput, , Null)
    cmdData.Parameters.Append cmdData.CreateParameter("IDVehiculoMantenimientoGrupo_FILTER", adInteger, adParamInput, , Null)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    
    If Not recData.EOF Then
        recData.MoveLast
        
        Check = recData.RecordCount
    End If
    
    recData.Close
    Set recData = Nothing

    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.VehiculoMantenimiento.Check", "Error verificar los Avisos de Mantenimiento de Vehículos."
End Function
