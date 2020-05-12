VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Telephony"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'TYPE
Public TelephonyType As String

'OBJETCS
Public TAPI As TAPI
Public Address As ITAddress
Public LocationInfo As ITLocationInfo

'DEVICES
Public DeviceGUID As String

'STATUS
Public Initialized As Boolean
Public CallerIDSupported As Boolean

'LOCATION INFO
Public LocationPulse As Boolean
Public LocationCountryCode As Long
Public LocationCityCode As String
Public LocationLocalAccessCode As String
Public LocationLongDistanceAccessCode As String

'CALLER ID
Public CallerID_Identificacion_IDSucursalTelefono As Integer
Public CallerIDServer_ValorTimer As Long
Public CallerIDEnabled As Boolean
Public CallerIDModemInitializationString As String
Public CallerIDModemCOMPort As Long

Private mRegistrationInstance As Long

Public Sub Initialize()
    Select Case TelephonyType
        Case "SERVER"
            Call SERVER_Initialize
        Case "TAPI"
            Call TAPI_Initialize
        Case "COMM"
            Call COMM_Initialize
        Case "NONE"
            Call NONE_Initialize
    End Select
    
    frmMDI.picCallerID.Visible = CallerIDSupported
End Sub

Private Sub SERVER_Initialize()
    CallerID_Identificacion_IDSucursalTelefono = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "CallerIDIdentificacionIDSucursalTelefono", "", csrdtNumberInteger)

    'LOCATION INFO
    LocationPulse = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationPulse", False, csrdtBoolean)
    LocationCountryCode = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationCountryCode", "54", csrdtString)
    LocationCityCode = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationCityCode", "02227", csrdtString)
    LocationLocalAccessCode = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationLocalAccessCode", "", csrdtString)
    LocationLongDistanceAccessCode = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationLongDistanceAccessCode", "", csrdtString)
    
    On Error Resume Next
    Err.Clear
    Set precSemaforoLlamada = New ADODB.Recordset
    precSemaforoLlamada.Source = "SELECT IDSucursalTelefono, ValorTimer, TelefonoNumero FROM SemaforoLlamada"
    precSemaforoLlamada.Open , pDatabase.Connection, adOpenKeyset, adLockOptimistic, adCmdText
    If Err.Number <> 0 Then
        ShowErrorMessage "Classes.Telephony.SERVER_Initialize", "Error al Abrir la Tabla de Semáforos de Llamadas."
        Exit Sub
    End If
    precSemaforoLlamada.Properties("Update Criteria").Value = adCriteriaKey
    
    Initialized = True
    CallerIDSupported = True
End Sub

Public Sub SERVER_CheckForCall()
    On Error GoTo ErrorHandler
    
    precSemaforoLlamada.Resync
    precSemaforoLlamada.Filter = "IDSucursalTelefono = " & CallerID_Identificacion_IDSucursalTelefono
    If Not (precSemaforoLlamada.BOF And precSemaforoLlamada.EOF) Then
        If precSemaforoLlamada("ValorTimer").Value <> CallerIDServer_ValorTimer Then
            If CallerIDServer_ValorTimer <> 0 Then
                frmMDI.txtCallerID.Text = precSemaforoLlamada("TelefonoNumero").Value & ""
                frmMDI.CallerID_BuscarPersonas
                frmMDI.CallerID_BuscarVehiculos
            End If
            CallerIDServer_ValorTimer = precSemaforoLlamada("ValorTimer").Value
        End If
    End If
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub SERVER_ShutDown()
    If Not precSemaforoLlamada Is Nothing Then
        If precSemaforoLlamada.State = adStateOpen Then
            precSemaforoLlamada.Close
        End If
        Set precSemaforoLlamada = Nothing
    End If
End Sub

Private Sub TAPI_Initialize()
    Screen.MousePointer = vbHourglass
    
    DeviceGUID = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "DeviceGUID", "", csrdtString)
    CallerIDEnabled = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "CallerIDEnabled", True, csrdtBoolean)
    
    WriteLogEvent "TAPI Initialization: INIT", vbLogEventTypeInformation
    
    On Error GoTo ErrorHandler
    
    Set TAPI = New TAPI
    TAPI.Initialize
    TAPI.EventFilter = TE_CALLMEDIA Or TE_CALLNOTIFICATION Or TE_CALLSTATE Or TE_ADDRESS Or TE_CALLINFOCHANGE
    
    Initialized = True
    
    WriteLogEvent "TAPI Initialization: DONE", vbLogEventTypeInformation

    TAPI_SelectAddress DeviceGUID
    
    frmMDI.SetTAPIObject
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Classes.Telephony.TAPI_Initialize", "Error al Inicializar la Interface de Telefonía."
End Sub

Private Function TAPI_SelectAddress(ByVal DeviceGUID As String) As Boolean
    Dim AddressCurrent As ITAddress
    Dim AddressCapabilities As ITAddressCapabilities
    Dim AddressTranslation As ITAddressTranslation
    
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErrorHandler
    
    Initialized = False
    
    'UNREGISTER EVENTS
    If mRegistrationInstance <> 0 Then
        Call TAPI.UnregisterNotifications(mRegistrationInstance)
    End If
    
    If DeviceGUID = "" Then
        If TAPI.Addresses.Count > 0 Then
            Set Address = TAPI.Addresses(1)
            WriteLogEvent "TAPI: Select Address Not Specifed, Default Selected - " & Address.AddressName, vbLogEventTypeInformation
        End If
    Else
        For Each AddressCurrent In TAPI.Addresses
            If AddressCurrent.QueryMediaType(TAPIMEDIATYPE_DATAMODEM) Then
                Set AddressCapabilities = AddressCurrent
                If AddressCapabilities.AddressCapabilityString(ACS_PERMANENTDEVICEGUID) = DeviceGUID Then
                    Set Address = AddressCurrent
                    WriteLogEvent "TAPI: Select Address Succesfully - " & Address.AddressName, vbLogEventTypeInformation
                End If
                Set AddressCapabilities = Nothing
            End If
        Next AddressCurrent
        If Address Is Nothing Then
            If TAPI.Addresses.Count > 0 Then
                Set Address = TAPI.Addresses(1)
                WriteLogEvent "TAPI: Select Address Not Found, Default Selected - " & Address.AddressName, vbLogEventTypeInformation
                MsgBox "El Dispotivio de Telefonía especificado en la Estación de Trabajo, no existe. Se ha seleccionado el primer Dispositivo disponible.", vbInformation, App.Title
            Else
                WriteLogEvent "TAPI: Select Address Not Found, Default Selected - " & Address.AddressName, vbLogEventTypeInformation
            End If
        End If
    End If
    
    If Not Address Is Nothing Then
        'REGISTER EVENTS
        mRegistrationInstance = TAPI.RegisterCallNotifications(Address, True, True, TAPIMEDIATYPE_AUDIO, 0)
        If mRegistrationInstance <> 0 Then
            Initialized = True
            CallerIDSupported = True
            
            'LOCATION INFO
            Set AddressTranslation = Address
            Set LocationInfo = AddressTranslation.Locations(1)
            LocationPulse = CBool(LocationInfo.Options)
            LocationCountryCode = LocationInfo.CountryCode
            LocationCityCode = LocationInfo.CityCode
            LocationLocalAccessCode = LocationInfo.LocalAccessCode
            LocationLongDistanceAccessCode = LocationInfo.LongDistanceAccessCode
        End If
    End If
    
    TAPI_SelectAddress = True
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.Telephony.TAPI_SelectAddress", "Error al Seleccionar el Dispositivo de Telefonía."
End Function

Private Sub TAPI_ShutDown()
    Set Address = Nothing
    If Initialized Then
        If mRegistrationInstance <> 0 Then
            Call TAPI.UnregisterNotifications(mRegistrationInstance)
        End If
        TAPI.Shutdown
    End If
    Set TAPI = Nothing
End Sub

Public Sub COMM_Initialize()
    Screen.MousePointer = vbHourglass
    
    CallerIDEnabled = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "CallerIDEnabled", True, csrdtBoolean)
    CallerIDModemInitializationString = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "CallerIDModemInitializationString", "AT#CID=1", csrdtString)
    CallerIDModemCOMPort = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "CallerIDModemCOMPort", 1, csrdtNumberInteger)

    CallerID_Identificacion_IDSucursalTelefono = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "CallerIDIdentificacionIDSucursalTelefono", "", csrdtNumberInteger)

    'LOCATION INFO
    LocationPulse = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationPulse", False, csrdtBoolean)
    LocationCountryCode = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationCountryCode", "54", csrdtString)
    LocationCityCode = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationCityCode", "02227", csrdtString)
    LocationLocalAccessCode = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationLocalAccessCode", "", csrdtString)
    LocationLongDistanceAccessCode = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationLongDistanceAccessCode", "", csrdtString)

    With frmMDI.comTelephony
        'MODEM INIT
        WriteLogEvent "COMM Initialization: INIT", vbLogEventTypeInformation
        
        On Error GoTo ErrorHandlerOpenPort
        
        If .PortOpen Then
            .PortOpen = False
        End If
        
        .CommPort = CallerIDModemCOMPort
        .PortOpen = True
        
        Initialized = True
        
        WriteLogEvent "COMM Initialization: DONE", vbLogEventTypeInformation
        
        'CALLER ID
        If CallerIDEnabled Then
            WriteLogEvent "COMM Caller ID: INIT", vbLogEventTypeInformation
            
            On Error GoTo ErrorHandlerCallerID
            
            .Output = CallerIDModemInitializationString & vbCr
            
            CallerIDSupported = True
            
            WriteLogEvent "COMM Caller ID: DONE", vbLogEventTypeInformation
        End If
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandlerOpenPort:
    ShowErrorMessage "Classes.Telephony.COMM_Initialize", "Error al Inicializar el Módem a través del Puerto de Comunicaciones."
    Exit Sub

ErrorHandlerCallerID:
    ShowErrorMessage "Classes.Telephony.COMM_Initialize", "Error al Activar la Identificación de Llamadas en el Módem a través del Puerto de Comunicaciones."
End Sub

Private Sub COMM_ShutDown()
    If frmMDI.comTelephony.PortOpen Then
        frmMDI.comTelephony.PortOpen = False
    End If
End Sub

Private Sub NONE_Initialize()
    'LOCATION INFO
    LocationPulse = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationPulse", False, csrdtBoolean)
    LocationCountryCode = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationCountryCode", "54", csrdtString)
    LocationCityCode = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationCityCode", "02227", csrdtString)
    LocationLocalAccessCode = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationLocalAccessCode", "", csrdtString)
    LocationLongDistanceAccessCode = CSM_Registry.GetValue_FromApplication_CurrentUser("Options\Telephony", "LocationLongDistanceAccessCode", "", csrdtString)
    
    Initialized = True
    CallerIDSupported = False
End Sub

Public Sub CallerID_Check()
    Dim BufferCurrent As String
    Dim Buffer As String
    Dim StartString As Long
    Dim EndString As Long
    Dim NumberIndex As Long
    Dim FechaTemp As String
    Dim RegistroLlamada As RegistroLlamada
    Dim SemaforoLlamada As SemaforoLlamada
    Dim CallerID As String
    Dim TelefonoNumero As String
    
    'Get data from buffer
    BufferCurrent = frmMDI.comTelephony.Input
    If BufferCurrent = "" Then
        Exit Sub
    End If
    
    Buffer = Buffer & BufferCurrent
    
    NumberIndex = InStr(1, Buffer, "NMBR = ")
    If NumberIndex > 0 Then
        Do
            DoEvents
            Buffer = Buffer & frmMDI.comTelephony.Input
        Loop Until InStr(NumberIndex, Buffer, vbCrLf) > 0
        
        'DATE
        StartString = InStr(1, Buffer, "DATE = ") + 7
        FechaTemp = Year(Date) & "/" & Mid(Buffer, StartString, 2) & "/" & Mid(Buffer, StartString + 2, 2)
        
        'TIME
        StartString = InStr(1, Buffer, "TIME = ") + 7
        FechaTemp = FechaTemp & " " & Mid(Buffer, StartString, 2) & ":" & Mid(Buffer, StartString + 2, 2)
        
        'NUMBER
        StartString = InStr(1, Buffer, "NMBR = ") + 7
        EndString = InStr(StartString, Buffer, vbCrLf)
        CallerID = Trim(Mid(Buffer, StartString, EndString - StartString))
        
        If pParametro.RegistroLlamada_Save Then
            Set RegistroLlamada = New RegistroLlamada
            RegistroLlamada.FechaHora = Now
            RegistroLlamada.IDSucursalTelefono = pTelephony.CallerID_Identificacion_IDSucursalTelefono
        End If
        Select Case CallerID
            Case ""
                frmMDI.txtCallerID.Text = CALLERID_NOTAVAILABLE
                If pParametro.RegistroLlamada_Save Then
                    TelefonoNumero = CALLERID_NOTAVAILABLE
                End If
            Case "P"
                frmMDI.txtCallerID.Text = CALLERID_PRIVATE
                If pParametro.RegistroLlamada_Save Then
                    TelefonoNumero = CALLERID_PRIVATE
                End If
            Case Else
                frmMDI.txtCallerID.Text = CallerID
                If pParametro.RegistroLlamada_Save Then
                    TelefonoNumero = IIf(Left(CallerID, 1) <> "0", pTelephony.LocationCityCode & CallerID, CallerID)
                End If
        End Select
        DoEvents
        If pParametro.RegistroLlamada_Save Then
            RegistroLlamada.TelefonoNumero = TelefonoNumero
            RegistroLlamada.Update
            Set RegistroLlamada = Nothing
            
            Set SemaforoLlamada = New SemaforoLlamada
            SemaforoLlamada.IDSucursalTelefono = pTelephony.CallerID_Identificacion_IDSucursalTelefono
            SemaforoLlamada.NoMatchRaiseError = False
            Call SemaforoLlamada.Load
            SemaforoLlamada.ValorTimer = timeGetTime()
            SemaforoLlamada.TelefonoNumero = TelefonoNumero
            SemaforoLlamada.Update
            Set SemaforoLlamada = Nothing
        End If
        
        frmMDI.CallerID_BuscarPersonas
        frmMDI.CallerID_BuscarVehiculos
        
        Buffer = ""
    End If
End Sub

Public Sub DialNumber(ByVal AreaCode As String, ByVal PhoneNumber As String)
    Dim DialableString As String
    Dim DisplayableString As String
    
    Dim AddressTranslation As ITAddressTranslation
    Dim AddressTranslationInfo As ITAddressTranslationInfo
    
    Select Case pParametro.Telephony_Type
        Case "TAPI"
            Set AddressTranslation = Address
            Set AddressTranslationInfo = AddressTranslation.TranslateAddress(IIf(AreaCode = "", "", "+" & LocationCountryCode & " (" & AreaCode & ") ") & PhoneNumber, 0, 0)
            Set AddressTranslation = Nothing
            DialableString = AddressTranslationInfo.DialableString
            DisplayableString = AddressTranslationInfo.DisplayableString
            Set AddressTranslationInfo = Nothing
            
            Address.CreateCall DialableString, LINEADDRESSTYPE_PHONENUMBER, TAPIMEDIATYPE_AUDIO
        Case "COMM"
            DialableString = "ATD"
            DialableString = DialableString & IIf(LocationPulse, "P", "T")
            DialableString = DialableString & IIf(AreaCode = LocationCityCode, LocationLocalAccessCode, LocationLongDistanceAccessCode)
            DialableString = DialableString & IIf(AreaCode = LocationCityCode, "", AreaCode)
            DialableString = DialableString & PhoneNumber
            
            DisplayableString = IIf(AreaCode = LocationCityCode, LocationLocalAccessCode, LocationLongDistanceAccessCode)
            DisplayableString = DisplayableString & IIf(DisplayableString = "", "", " ") & IIf(AreaCode = LocationCityCode, "", "(" & AreaCode & ") ")
            DisplayableString = DisplayableString & PhoneNumber
            
            frmMDI.comTelephony.Output = DialableString
        Case "NONE"
            DialableString = ""
            DisplayableString = ""
    End Select
    
    If DisplayableString <> "" Then
        MsgBox "Llamando al Número: " & DisplayableString, vbInformation, App.Title
    End If
End Sub

Private Sub Class_Terminate()
    Select Case TelephonyType
        Case "TAPI"
            Call TAPI_ShutDown
        Case "COMM"
            Call COMM_ShutDown
        Case "NONE"
    End Select
End Sub
