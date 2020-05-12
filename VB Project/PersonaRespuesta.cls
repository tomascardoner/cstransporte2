VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PersonaRespuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IDPersona As Long
Public FechaHora As Date
Public Respuesta As String
Public Activo As Boolean
Public FechaHoraCreacion As Date
Public IDUsuarioCreacion As Integer
Public FechaHoraModificacion As Date
Public IDUsuarioModificacion As String

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
    cmdData.CommandText = "sp_PersonaRespuesta_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , IDPersona)
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , FechaHora)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.PersonaRespuesta.OpenRecordset", "Error al abrir la tabla de Respuestas de las Personas."
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
        
        Respuesta = mrecData("Respuesta").Value
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
    ShowErrorMessage "Classes.PersonaRespuesta.Load", "Error al obtener los datos de la Respuesta de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "Fecha/Hora: " & Format(FechaHora, "Short Date") & " " & Format(FechaHora, "Short Time")
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
    mrecData("IDPersona").Value = IDPersona
    mrecData("FechaHora").Value = FechaHora
    mrecData("Respuesta").Value = Respuesta
    mrecData("Activo").Value = Activo
    mrecData("FechaHoraCreacion").Value = Now
    mrecData("IDUsuarioCreacion").Value = pUsuario.IDUsuario
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If RefreshList Then
        RefreshList_RefreshPersonaRespuesta IDPersona, FechaHora
    End If
    
    Screen.MousePointer = vbDefault
    AddNew = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe una Respuesta de la Persona con la misma Fecha/Hora.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.PersonaRespuesta.AddNew", "Error al Agregar la Respuesta de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "Fecha/Hora: " & Format(FechaHora, "Short Date") & " " & Format(FechaHora, "Short Time")
        End If
    Else
        ShowErrorMessage "Classes.PersonaRespuesta.AddNew", "Error al Agregar la Respuesta de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "Fecha/Hora: " & Format(FechaHora, "Short Date") & " " & Format(FechaHora, "Short Time")
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
    
    mrecData("FechaHora").Value = FechaHora
    mrecData("Respuesta").Value = Respuesta
    mrecData("Activo").Value = Activo
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If RefreshList Then
        RefreshList_RefreshPersonaRespuesta IDPersona, FechaHora
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe una Respuesta de la Persona con la misma Fecha/Hora.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.PersonaRespuesta.Update", "Error al Actualizar la Respuesta de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "Fecha/Hora: " & Format(FechaHora, "Short Date") & " " & Format(FechaHora, "Short Time")
        End If
    Else
        ShowErrorMessage "Classes.PersonaRespuesta.Update", "Error al Actualizar laRespuesta de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "Fecha/Hora: " & Format(FechaHora, "Short Date") & " " & Format(FechaHora, "Short Time")
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
        RefreshList_RefreshPersonaRespuesta IDPersona, DATE_TIME_FIELD_NULL_VALUE
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.PersonaRespuesta.Delete", "Error al Eliminar laRespuesta de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "Fecha/Hora: " & Format(FechaHora, "Short Date") & " " & Format(FechaHora, "Short Time")
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Public Function ActivateAll(ByVal Value As Boolean) As Boolean
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_PersonaRespuesta_ActivateAll"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , IDPersona)
    cmdData.Parameters.Append cmdData.CreateParameter("Valor", adBoolean, adParamInput, , Value)

    cmdData.Execute
    
    Set cmdData = Nothing

    If RefreshList Then
        RefreshList_RefreshPersonaRespuesta IDPersona, DATE_TIME_FIELD_NULL_VALUE
    End If
    
    Screen.MousePointer = vbDefault
    ActivateAll = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.PersonaRespuesta.ActivateAll", "Error al Activar o Desactivar las Respuestas de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona
End Function

Private Sub InitializeValues()
    FechaHora = DATE_TIME_FIELD_NULL_VALUE
    Respuesta = ""
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

