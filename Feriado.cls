VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Feriado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Fecha As Date
Public Nombre As String
Public FechaHoraCreacion As Date
Public IDUsuarioCreacion As Integer
Public FechaHoraModificacion As Date
Public IDUsuarioModificacion As Integer

Public NoMatchRaiseError As Boolean
Public NoMatch As Boolean

Public RefreshList As Boolean

Private mrecData As ADODB.Recordset

Public Property Get Fecha_Formatted() As String
    Fecha_Formatted = Format(Fecha, "Short Date")
End Property

Private Function OpenRecordset() As Boolean
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Feriado_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("Fecha_FILTER", adChar, adParamInput, 10, Format(Fecha, "yyyy/mm/dd"))
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Feriado.OpenRecordset", "Error al abrir la tabla de Feriados."
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
        
        Nombre = mrecData("Nombre").Value & ""
        FechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        IDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        FechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        IDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Feriado.Load", "Error al obtener los datos del Feriado." & vbCr & vbCr & "Fecha: " & Fecha_Formatted
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
    mrecData("Fecha").Value = Fecha
    mrecData("Nombre").Value = IIf(Trim(Nombre) = "", Null, Nombre)
    mrecData("FechaHoraCreacion").Value = Now
    mrecData("IDUsuarioCreacion").Value = pUsuario.IDUsuario
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If RefreshList Then
        RefreshList_RefreshFeriado Fecha
    End If
    
    Screen.MousePointer = vbDefault
    AddNew = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe un Feriado con la misma Fecha.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Feriado.AddNew", "Error al Agregar el Feriado." & vbCr & vbCr & "Fecha: " & Fecha_Formatted
        End If
    Else
        ShowErrorMessage "Classes.Feriado.AddNew", "Error al Agregar el Feriado." & vbCr & vbCr & "Fecha: " & Fecha_Formatted
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
    
    mrecData("Fecha").Value = Fecha
    mrecData("Nombre").Value = IIf(Trim(Nombre) = "", Null, Nombre)
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    If RefreshList Then
        RefreshList_RefreshFeriado Fecha
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe un Feriado con la misma Fecha.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Feriado.Update", "Error al Actualizar el Feriado." & vbCr & vbCr & "Fecha: " & Fecha_Formatted
        End If
    Else
        ShowErrorMessage "Classes.Feriado.Update", "Error al Actualizar el Feriado." & vbCr & vbCr & "Fecha: " & Fecha_Formatted
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
        RefreshList_RefreshFeriado DATE_TIME_FIELD_NULL_VALUE
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_RELATED_RECORDS Then
            Screen.MousePointer = vbDefault
            MsgBox "No se puede Eliminar el Feriado debido a que tiene datos relacionados.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.Feriado.Delete", "Error al Eliminar el Feriado." & vbCr & vbCr & "Fecha: " & Fecha_Formatted
        End If
    Else
        ShowErrorMessage "Classes.Feriado.Delete", "Error al Eliminar el Feriado." & vbCr & vbCr & "Fecha: " & Fecha_Formatted
    End If
    If mrecData.EditMode = adEditInProgress Then
        mrecData.CancelUpdate
    End If
End Function

Private Sub InitializeValues()
    Fecha = DATE_TIME_FIELD_NULL_VALUE
    Nombre = ""
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

Public Sub VerificarReservasDelPasajero(ByVal IDPersona As Long)
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    Dim MessageBox As String
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_Feriado_ViajeDetalle_List"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , IDPersona)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenStatic, adLockOptimistic
    Set cmdData = Nothing
    
    If Not recData.EOF Then
        recData.MoveLast
        recData.MoveFirst
        If recData.RecordCount = 1 Then
            MessageBox = "Este Pasajero tiene una Reserva para un Día Feriado." & vbCr
        Else
            MessageBox = "Este Pasajero tiene " & recData.RecordCount & " Reservas para uno o más Días Feriados." & vbCr
        End If
        Do While Not recData.EOF
            MessageBox = MessageBox & vbCr & "Fecha/Hora: " & Format(recData("FechaHora").Value, "Short Date") & " " & Format(recData("FechaHora").Value, "Short Time") & " - Ruta: " & RTrim(recData("IDRuta").Value) & IIf(IsNull(recData("Nombre").Value), "", " - Feriado: " & recData("Nombre").Value)
            recData.MoveNext
        Loop
        MsgBox MessageBox, vbExclamation, App.Title
    End If
    
    recData.Close
    Set recData = Nothing

    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.Feriado.VerificarReservasDelPasajero", "Error al verificar si el Pasajero tiene Reservas un Día Feriado."
End Sub
