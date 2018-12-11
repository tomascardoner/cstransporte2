VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Localidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IDProvincia As String
Public IDLocalidad As Long
Public Nombre As String
Public CodigoPostal As Integer

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
    cmdData.CommandText = "sp_Localidad_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDProvincia_FILTER", adChar, adParamInput, 1, IDProvincia)
    cmdData.Parameters.Append cmdData.CreateParameter("IDLocalidad_FILTER", adInteger, adParamInput, , IDLocalidad)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Localidad.OpenRecordset", "Error al abrir la tabla de Localidades."
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
        CodigoPostal = Val(mrecData("CodigoPostal").Value & "")
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Localidad.Load", "Error al obtener los datos de la Localidad." & vbCr & vbCr & "IDProvincia: " & IDProvincia & vbCr & "IDLocalidad: " & IDLocalidad
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
    mrecData("IDProvincia").Value = IDProvincia
    mrecData("IDLocalidad").Value = IDLocalidad
    mrecData("Nombre").Value = Nombre
    mrecData("CodigoPostal").Value = IIf(CodigoPostal = 0, Null, CodigoPostal)
    mrecData.Update
    
    Screen.MousePointer = vbDefault
    AddNew = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe una Localidad con el mismo Nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Localidad.AddNew", "Error al Agregar la Localidad." & vbCr & vbCr & "IDProvincia: " & IDProvincia & vbCr & "IDLocalidad: " & IDLocalidad
        End If
    Else
        ShowErrorMessage "Classes.Localidad.AddNew", "Error al Agregar la Localidad." & vbCr & vbCr & "IDProvincia: " & IDProvincia & vbCr & "IDLocalidad: " & IDLocalidad
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
    mrecData("CodigoPostal").Value = CodigoPostal
    mrecData.Update
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe una Localidad con el mismo Nombre.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Localidad.Update", "Error al Actualizar la Localidad." & vbCr & vbCr & "IDProvincia: " & IDProvincia & vbCr & "IDLocalidad: " & IDLocalidad
        End If
    Else
        ShowErrorMessage "Classes.Localidad.Update", "Error al Actualizar la Localidad." & vbCr & vbCr & "IDProvincia: " & IDProvincia & vbCr & "IDLocalidad: " & IDLocalidad
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
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede Eliminar la Localidad debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Localidad.Delete", "Error al Eliminar la Localidad." & vbCr & vbCr & "IDProvincia: " & IDProvincia & vbCr & "IDLocalidad: " & IDLocalidad
        End If
    Else
        ShowErrorMessage "Classes.Localidad.Delete", "Error al Eliminar la Localidad." & vbCr & vbCr & "IDProvincia: " & IDProvincia & vbCr & "IDLocalidad: " & IDLocalidad
    End If
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Private Sub InitializeValues()
    Nombre = ""
    CodigoPostal = 0
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

