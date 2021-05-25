VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListaPrecioDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IDListaPrecio As Long
Public OcupanteTipo As String
Public IDRuta As String
Public IDLugarGrupoOrigen As Long
Public IDLugarGrupoDestino As Long
Public Importe As Currency
Public ImporteWeb As Currency
Public FechaHoraCreacion As Date
Public IDUsuarioCreacion As Integer
Public FechaHoraModificacion As Date
Public IDUsuarioModificacion As Integer

Public IDOrigen As Long
Public IDDestino As Long

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
    cmdData.CommandText = "sp_ListaPrecioDetalle_Data"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDListaPrecio_FILTER", adInteger, adParamInput, , IDListaPrecio)
    cmdData.Parameters.Append cmdData.CreateParameter("OcupanteTipo_FILTER", adChar, adParamInput, 2, OcupanteTipo)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, IDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("IDLugarGrupoOrigen_FILTER", adInteger, adParamInput, , IDLugarGrupoOrigen)
    cmdData.Parameters.Append cmdData.CreateParameter("IDLugarGrupoDestino_FILTER", adInteger, adParamInput, , IDLugarGrupoDestino)
    Set mrecData = New ADODB.Recordset
    mrecData.Open cmdData, , adOpenKeyset, adLockOptimistic
    mrecData.Properties("Update Criteria").value = adCriteriaKey
    Set cmdData = Nothing

    Screen.MousePointer = vbDefault
    OpenRecordset = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.ListaPrecioDetalle.OpenRecordset", "Error al abrir la tabla de Detalles de Listas de Precios."
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
        
    Else
        NoMatch = False
        
        Importe = mrecData("Importe").value
        ImporteWeb = IIf(IsNull(mrecData("ImporteWeb").value), -1, mrecData("ImporteWeb").value)
        FechaHoraCreacion = mrecData("FechaHoraCreacion").value
        IDUsuarioCreacion = mrecData("IDUsuarioCreacion").value
        FechaHoraModificacion = mrecData("FechaHoraModificacion").value
        IDUsuarioModificacion = mrecData("IDUsuarioModificacion").value
    End If
    
    Screen.MousePointer = vbDefault
    Load = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.ListaPrecioDetalle.Load", "Error al obtener los datos del Detalle de la Lista de Precios." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio & vbCr & "Tipo: " & OcupanteTipo & vbCr & "IDRuta: " & IDRuta & vbCr & "IDOrigen: " & IDOrigen & vbCr & "IDDestino: " & IDDestino
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
    mrecData("IDListaPrecio").value = IDListaPrecio
    mrecData("OcupanteTipo").value = OcupanteTipo
    mrecData("IDRuta").value = IDRuta
    mrecData("IDLugarGrupoOrigen").value = IDLugarGrupoOrigen
    mrecData("IDLugarGrupoDestino").value = IDLugarGrupoDestino
    mrecData("Importe").value = Importe
    mrecData("ImporteWeb").value = IIf(ImporteWeb = -1, Null, ImporteWeb)
    mrecData("FechaHoraCreacion").value = Now
    mrecData("IDUsuarioCreacion").value = pUsuario.IDUsuario
    mrecData("FechaHoraModificacion").value = Now
    mrecData("IDUsuarioModificacion").value = pUsuario.IDUsuario
    mrecData.Update
    
    Screen.MousePointer = vbDefault
    AddNew = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.ListaPrecioDetalle.AddNew", "Error al Agregar el Detalle de la Lista de Precios." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio & vbCr & "Tipo: " & OcupanteTipo & vbCr & "IDRuta: " & IDRuta & vbCr & "IDOrigen: " & IDOrigen & vbCr & "IDDestino: " & IDDestino
    mrecData.CancelUpdate
End Function

Public Function Update() As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    mrecData("Importe").value = Importe
    mrecData("ImporteWeb").value = IIf(ImporteWeb = -1, Null, ImporteWeb)
    mrecData("FechaHoraModificacion").value = Now
    mrecData("IDUsuarioModificacion").value = pUsuario.IDUsuario
    mrecData.Update
    
    Screen.MousePointer = vbDefault
    Update = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.ListaPrecioDetalle.Update", "Error al Actualizar el Detalle de la Lista de Precios." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio & vbCr & "Tipo: " & OcupanteTipo & vbCr & "IDRuta: " & IDRuta & vbCr & "IDOrigen: " & IDOrigen & vbCr & "IDDestino: " & IDDestino
    mrecData.CancelUpdate
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
    ShowErrorMessage "Classes.ListaPrecioDetalle.Delete", "Error al Eliminar el Detalle de la Lista de Precios." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio & vbCr & "Tipo: " & OcupanteTipo & vbCr & "IDRuta: " & IDRuta & vbCr & "IDOrigen: " & IDOrigen & vbCr & "IDDestino: " & IDDestino
    mrecData.CancelUpdate
End Function

Private Sub InitializeValues()
    Importe = -1
    ImporteWeb = -1
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

'////////////////////////////////////////////////////////////////////////////
'OBTENGO EL IMPORTE EN BASE EL LUGAR DE ORIGEN Y DE DESTINO
'////////////////////////////////////////////////////////////////////////////
Public Function GetImporteByLugar() As Boolean
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ListaPrecioDetalle_Importe"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDListaPrecio_FILTER", adInteger, adParamInput, , IDListaPrecio)
    cmdData.Parameters.Append cmdData.CreateParameter("OcupanteTipo_FILTER", adChar, adParamInput, 2, OcupanteTipo)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, IDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("IDOrigen_FILTER", adInteger, adParamInput, , IDOrigen)
    cmdData.Parameters.Append cmdData.CreateParameter("IDDestino_FILTER", adInteger, adParamInput, , IDDestino)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenKeyset, adLockOptimistic
    Set cmdData = Nothing
    
    If recData.EOF Then
        Importe = -1
        ImporteWeb = -1
    Else
        Importe = recData("Importe").value
        ImporteWeb = IIf(IsNull(recData("ImporteWeb").value), -1, recData("ImporteWeb").value)
    End If
    
    recData.Close
    Set recData = Nothing
    
    Screen.MousePointer = vbDefault
    GetImporteByLugar = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.ListaPrecioDetalle.GetImporteByLugar", "Error al obtener el Importe del Detalle de la Lista de Precios." & vbCr & vbCr & "IDListaPrecio: " & IDListaPrecio & vbCr & "Tipo: " & OcupanteTipo & vbCr & "IDRuta: " & IDRuta & vbCr & "IDOrigen: " & IDOrigen & vbCr & "IDDestino: " & IDDestino
End Function
