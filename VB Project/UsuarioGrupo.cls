VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UsuarioGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IDUsuarioGrupo As Byte
Public Nombre As String
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

Private Function OpenRecordset() As Boolean
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_UsuarioGrupo_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDUsuarioGrupo", adTinyInt, adParamInput, , IDUsuarioGrupo)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing
    
    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.UsuarioGrupo.OpenRecordset", "Error al abrir la tabla de Grupos de Usuarios."
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
        
        Nombre = mrecData("Nombre").Value
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
    ShowErrorMessage "Classes.UsuarioGrupo.Load", "Error al obtener los datos del Grupo de Usuarios." & vbCr & vbCr & "IDUsuarioGrupo: " & IDUsuarioGrupo
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
    cmdData.CommandText = "usp_UsuarioGrupo_IDMax"
    cmdData.CommandType = adCmdStoredProc
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
    IDUsuarioGrupo = Val(recData("IDUsuarioGrupoMax").Value & "") + 1
    recData.Close
    Set recData = Nothing
    
    mrecData.AddNew
    mrecData("IDUsuarioGrupo").Value = IDUsuarioGrupo
    mrecData("Nombre").Value = Nombre
    mrecData("Notas").Value = IIf(Trim(Notas) = "", Null, Notas)
    mrecData("Activo").Value = Activo
    mrecData("FechaHoraCreacion").Value = Now
    mrecData("IDUsuarioCreacion").Value = pUsuario.IDUsuario
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If RefreshList Then
        RefreshList_RefreshUsuarioGrupo IDUsuarioGrupo
    End If
    
    Screen.MousePointer = vbDefault
    AddNew = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe un Grupo de Usuarios con el mismo Nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.UsuarioGrupo.AddNew", "Error al Agregar el Grupo de Usuarios." & vbCr & vbCr & "IDUsuarioGrupo: " & IDUsuarioGrupo
        End If
    Else
        ShowErrorMessage "Classes.UsuarioGrupo.AddNew", "Error al Agregar el Grupo de Usuarios." & vbCr & vbCr & "IDUsuarioGrupo: " & IDUsuarioGrupo
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
    
    mrecData("Nombre").Value = Nombre
    mrecData("Notas").Value = IIf(Trim(Notas) = "", Null, Notas)
    mrecData("Activo").Value = Activo
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If RefreshList Then
        RefreshList_RefreshUsuarioGrupo IDUsuarioGrupo
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe un Grupo de Usuarios con el mismo Nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.UsuarioGrupo.Update", "Error al Actualizar el Grupo de Usuarios." & vbCr & vbCr & "IDUsuarioGrupo: " & IDUsuarioGrupo
        End If
    Else
        ShowErrorMessage "Classes.UsuarioGrupo.Update", "Error al Actualizar el Grupo de Usuarios." & vbCr & vbCr & "IDUsuarioGrupo: " & IDUsuarioGrupo
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
        RefreshList_RefreshUsuarioGrupo 0
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede Eliminar el Grupo de Usuarios debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.UsuarioGrupo.Delete", "Error al Eliminar el Grupo de Usuarios." & vbCr & vbCr & "IDUsuarioGrupo: " & IDUsuarioGrupo
        End If
    Else
        ShowErrorMessage "Classes.UsuarioGrupo.Delete", "Error al Eliminar el Grupo de Usuarios." & vbCr & vbCr & "IDUsuarioGrupo: " & IDUsuarioGrupo
    End If
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Private Sub InitializeValues()
    Nombre = ""
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

