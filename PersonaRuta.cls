VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PersonaRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IDPersona As Long
Public IDRuta As String
Public IDOrigen As Long
Public Sube As String
Public IDDestino As Long
Public Baja As String
Public IDListaPrecio As Long
Public FechaHoraCreacion As Date
Public IDUsuarioCreacion As Integer
Public FechaHoraModificacion As Date
Public IDUsuarioModificacion As Integer

Public NoMatchRaiseError As Boolean
Public NoMatch As Boolean

Public RefreshList As Boolean

Private mrecData As ADODB.Recordset

Private mIDOrigenOriginal As Long
Private mSubeOriginal As String
Private mIDDestinoOriginal As Long
Private mBajaOriginal As String

Private Function OpenRecordset() As Boolean
    Dim cmdData As ADODB.command

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
        
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_PersonaRuta_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , IDPersona)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, IDRuta)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").Value = adCriteriaKey
    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.PersonaRuta.OpenRecordset", "Error al abrir la tabla de Rutas de Personas."
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
        
        IDOrigen = mrecData("IDOrigen").Value
        Sube = mrecData("Sube").Value & ""
        IDDestino = mrecData("IDDestino").Value
        Baja = mrecData("Baja").Value & ""
        IDListaPrecio = mrecData("IDListaPrecio").Value
        FechaHoraCreacion = mrecData("FechaHoraCreacion").Value
        IDUsuarioCreacion = mrecData("IDUsuarioCreacion").Value
        FechaHoraModificacion = mrecData("FechaHoraModificacion").Value
        IDUsuarioModificacion = mrecData("IDUsuarioModificacion").Value
    End If
    mIDOrigenOriginal = IDOrigen
    mSubeOriginal = Sube
    mIDDestinoOriginal = IDDestino
    mBajaOriginal = Baja
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.PersonaRuta.Load", "Error al obtener los datos de la Ruta de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "IDRuta: " & IDRuta
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
    mrecData("IDRuta").Value = IDRuta
    mrecData("IDOrigen").Value = IDOrigen
    mrecData("Sube").Value = IIf(Trim(Sube) = "", Null, Sube)
    mrecData("IDDestino").Value = IDDestino
    mrecData("Baja").Value = IIf(Trim(Baja) = "", Null, Baja)
    mrecData("IDListaPrecio").Value = IDListaPrecio
    mrecData("FechaHoraCreacion").Value = Now
    mrecData("IDUsuarioCreacion").Value = pUsuario.IDUsuario
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    Call ActualizarReservasFuturas
    
    If RefreshList Then
        RefreshList_RefreshPersonaRuta IDPersona, IDRuta
        RefreshList_RefreshViajeDetalle Date, "", 0, True
    End If
    
    Screen.MousePointer = vbDefault
    AddNew = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe la Ruta para la Persona.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.PersonaRuta.AddNew", "Error al Agregar la Ruta de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "IDRuta: " & IDRuta
        End If
    Else
        ShowErrorMessage "Classes.PersonaRuta.AddNew", "Error al Agregar la Ruta de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "IDRuta: " & IDRuta
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
    
    mrecData("IDPersona").Value = IDPersona
    mrecData("IDRuta").Value = IDRuta
    mrecData("IDOrigen").Value = IDOrigen
    mrecData("Sube").Value = IIf(Trim(Sube) = "", Null, Sube)
    mrecData("IDDestino").Value = IDDestino
    mrecData("Baja").Value = IIf(Trim(Baja) = "", Null, Baja)
    mrecData("IDListaPrecio").Value = IDListaPrecio
    mrecData("FechaHoraModificacion").Value = Now
    mrecData("IDUsuarioModificacion").Value = pUsuario.IDUsuario
    mrecData.Update
    
    Call ActualizarReservasFuturas
    
    If RefreshList Then
        RefreshList_RefreshPersonaRuta IDPersona, IDRuta
        RefreshList_RefreshViajeDetalle Date, "", 0, True
    End If
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = pDatabase.ERROR_DUPLICATE_ALTERNATEKEY Then
            Screen.MousePointer = vbDefault
            MsgBox "Ya existe la Ruta para la Persona.", vbExclamation, App.Title
        Else
            ShowErrorMessage "Classes.PersonaRuta.Update", "Error al Actualizar la Ruta de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "IDRuta: " & IDRuta
        End If
    Else
        ShowErrorMessage "Classes.PersonaRuta.Update", "Error al Actualizar la Ruta de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "IDRuta: " & IDRuta
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
        RefreshList_RefreshPersonaRuta IDPersona, ""
    End If
    
    Screen.MousePointer = vbDefault
    Delete = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.PersonaRuta.Delete", "Error al Eliminar la Ruta de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "IDRuta: " & IDRuta
    mrecData.CancelUpdate
End Function

Private Sub InitializeValues()
    IDOrigen = 0
    Sube = ""
    IDDestino = 0
    Baja = ""
    IDListaPrecio = 0
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

Public Property Get Persona() As Persona
    Dim PersonaObject As Persona
    
    Set PersonaObject = New Persona
    PersonaObject.IDPersona = IDPersona
    Call PersonaObject.Load
    Set Persona = PersonaObject
End Property

'============================================================================
' Actualiza las Reservas a Futuro de un Pasajero, dada una Ruta
'============================================================================
Public Function ActualizarReservasFuturas() As Boolean
    Dim Viaje As Viaje
    Dim ViajeDetalle As ViajeDetalle

    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    Dim Ruta As Ruta

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ViajeDetalle_ByPersonaRuta"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDPersona_FILTER", adInteger, adParamInput, , IDPersona)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, IDRuta)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing

    Do While Not recData.EOF
        'Si no se ha modificado el Origen y el Destino, la Actualizo
        Set Ruta = New Ruta
        Ruta.IDRuta = IDRuta
        If Ruta.Load() Then
            If (recData("IDOrigen").Value = Ruta.IDOrigen _
                And IsNull(recData("Sube").Value) _
                And recData("IDDestino").Value = Ruta.IDDestino _
                And IsNull(recData("Baja").Value)) _
                Or (recData("IDOrigen").Value = mIDOrigenOriginal _
                And recData("Sube").Value & "" = mSubeOriginal _
                And recData("IDDestino").Value = mIDDestinoOriginal _
                And recData("Baja").Value & "" = mBajaOriginal) Then
                
                Set ViajeDetalle = New ViajeDetalle
                ViajeDetalle.RefreshListSkip = True
        
                ViajeDetalle.FechaHora = recData("FechaHora").Value
                ViajeDetalle.IDRuta = IDRuta
                ViajeDetalle.Indice = recData("Indice").Value
                If ViajeDetalle.Load Then
                    Set Viaje = New Viaje
                    Viaje.FechaHora = recData("FechaHora").Value
                    Viaje.IDRuta = IDRuta
                    If Viaje.Load() Then
                        If ViajeDetalle.GetNewValues(Viaje.DiaSemanaBase) Then
                            Call ViajeDetalle.MakeDirty
                            Call ViajeDetalle.Update
                        End If
                    End If
                    Set Viaje = Nothing
                End If
                Set ViajeDetalle = Nothing
            End If
        End If
        Set Ruta = Nothing
        
        recData.MoveNext
    Loop

    recData.Close
    Set recData = Nothing

    Set ViajeDetalle = Nothing

    Screen.MousePointer = vbDefault
    ActualizarReservasFuturas = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Classes.PersonaRuta.ActualizarReservasFijas", "Error al Actualizar las Reservas Futuras del Pasajero para la Ruta." & vbCr & vbCr & "IDPersona: " & IDPersona & vbCr & "IDRuta: " & IDRuta
End Function
