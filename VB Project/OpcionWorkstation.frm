VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOpcionWorkstation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones de la Estación de Trabajo"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OpcionWorkstation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   8070
   Begin VB.CheckBox chkSyncDateTimeWithDBServer_Enabled 
      Caption         =   "Sincronizar Fecha/Hora con Servidor de BD"
      Height          =   210
      Left            =   4140
      TabIndex        =   35
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Frame fraSucursal 
      Height          =   675
      Left            =   4140
      TabIndex        =   5
      Top             =   600
      Width           =   3795
      Begin MSDataListLib.DataCombo datcboSucursal 
         Height          =   330
         Left            =   900
         TabIndex        =   7
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblSucursal 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal:"
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   690
      End
   End
   Begin VB.Frame fraRefreshList 
      Height          =   1395
      Left            =   120
      TabIndex        =   30
      Top             =   4200
      Width           =   3915
      Begin VB.TextBox txtRefreshList_Slower_CheckInterval 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   36
         Top             =   960
         Width           =   570
      End
      Begin VB.CheckBox chkRefreshList_Enabled 
         Caption         =   "Actualizar Listas al registrar cambios"
         Height          =   210
         Left            =   180
         TabIndex        =   31
         Top             =   240
         Width           =   3315
      End
      Begin VB.TextBox txtRefreshList_Fastest_CheckInterval 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   33
         Top             =   540
         Width           =   570
      End
      Begin VB.Label lblRefreshList_Slowest_CheckInterval 
         AutoSize        =   -1  'True
         Caption         =   "Lento:                     segs."
         Height          =   210
         Left            =   480
         TabIndex        =   34
         Top             =   1020
         Width           =   1800
      End
      Begin VB.Label lblRefreshList_Fastest_CheckInterval 
         AutoSize        =   -1  'True
         Caption         =   "Rápido:                   segs."
         Height          =   210
         Left            =   480
         TabIndex        =   32
         Top             =   600
         Width           =   1800
      End
   End
   Begin VB.Frame fraTelephonyLocation 
      Caption         =   "Telefonía: Ubicación"
      Height          =   2655
      Left            =   4140
      TabIndex        =   18
      Top             =   1500
      Width           =   3795
      Begin VB.OptionButton optTelephonyLocationPulse 
         Caption         =   "Pulso"
         Height          =   210
         Left            =   2760
         TabIndex        =   29
         Top             =   2280
         Width           =   855
      End
      Begin VB.OptionButton optTelephonyLocationTone 
         Caption         =   "Tono"
         Height          =   210
         Left            =   1800
         TabIndex        =   28
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtTelephonyLocationLongDistanceAccessCode 
         Height          =   315
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   26
         Top             =   1740
         Width           =   615
      End
      Begin VB.TextBox txtTelephonyLocationLocalAccessCode 
         Height          =   315
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   24
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTelephonyLocationCityCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   22
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtTelephonyLocationCountryCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   20
         Top             =   300
         Width           =   435
      End
      Begin VB.Label lblTelephonyLocationDialMode 
         AutoSize        =   -1  'True
         Caption         =   "Modo de Discado:"
         Height          =   210
         Left            =   180
         TabIndex        =   27
         Top             =   2280
         Width           =   1290
      End
      Begin VB.Label lblTelephonyLocationLongDistanceAccessCode 
         Caption         =   "Código de acceso a línea externa para llamadas de larga distancia:"
         Height          =   450
         Left            =   180
         TabIndex        =   25
         Top             =   1680
         Width           =   2565
      End
      Begin VB.Label lblTelephonyLocationLocalAccessCode 
         Caption         =   "Código de acceso a línea externa para llamadas locales:"
         Height          =   450
         Left            =   180
         TabIndex        =   23
         Top             =   1140
         Width           =   2565
      End
      Begin VB.Label lblTelephonyLocationCityCode 
         AutoSize        =   -1  'True
         Caption         =   "Código de Area:"
         Height          =   210
         Left            =   180
         TabIndex        =   21
         Top             =   780
         Width           =   1170
      End
      Begin VB.Label lblTelephonyLocationCountryCode 
         AutoSize        =   -1  'True
         Caption         =   "Código de País:"
         Height          =   210
         Left            =   180
         TabIndex        =   19
         Top             =   360
         Width           =   1110
      End
   End
   Begin VB.Frame fraTelephonyTAPI 
      Height          =   675
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3915
      Begin VB.ComboBox cboTelephonyAddress 
         Height          =   330
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2970
      End
      Begin VB.Label lblTelephonyAddress 
         AutoSize        =   -1  'True
         Caption         =   "Módem:"
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.ComboBox cboTelephonyType 
      Height          =   330
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6720
      TabIndex        =   38
      Top             =   5220
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   37
      Top             =   5220
      Width           =   1215
   End
   Begin VB.Frame fraTelephonyCallerID 
      Caption         =   "Identificación de Llamadas"
      Height          =   2655
      Left            =   120
      TabIndex        =   8
      Top             =   1500
      Width           =   3915
      Begin VB.ComboBox cboTelephonyCallerIDIdentificacion 
         Height          =   330
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2160
         Width           =   2475
      End
      Begin VB.CheckBox chkRegistroLlamada_Save 
         Caption         =   "Guardar Registro de Llamadas"
         Height          =   210
         Left            =   180
         TabIndex        =   15
         Top             =   1860
         Width           =   2835
      End
      Begin VB.Frame fraTelephonyCOMMCallerID 
         Height          =   1155
         Left            =   180
         TabIndex        =   10
         Top             =   540
         Width           =   3615
         Begin VB.ComboBox cboTelephonyCallerIDModemPort 
            Height          =   330
            Left            =   780
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   240
            Width           =   2730
         End
         Begin VB.ComboBox cboTelephonyCallerIDModemInitializationString 
            Height          =   330
            Left            =   2220
            TabIndex        =   14
            Top             =   660
            Width           =   1290
         End
         Begin VB.Label lblTelephonyCallerIDModemPort 
            AutoSize        =   -1  'True
            Caption         =   "Módem:"
            Height          =   210
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblTelephonyCallerIDModemInitializationString 
            AutoSize        =   -1  'True
            Caption         =   "Commando de Inicialización:"
            Height          =   210
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   2010
         End
      End
      Begin VB.CheckBox chkTelephonyCallerIDEnabled 
         Caption         =   "Habilitada"
         Height          =   210
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblTelephonyCallerIDIdentificacion 
         AutoSize        =   -1  'True
         Caption         =   "Identificación:"
         Height          =   210
         Left            =   180
         TabIndex        =   16
         Top             =   2220
         Width           =   990
      End
   End
   Begin VB.Label lblTelephonyType 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Telefonía:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmOpcionWorkstation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCTelephonyDeviceGUID As Collection

Private Sub chkRefreshList_Enabled_Click()
    lblRefreshList_Fastest_CheckInterval.Visible = (chkRefreshList_Enabled.Value = vbChecked)
    txtRefreshList_Fastest_CheckInterval.Visible = (chkRefreshList_Enabled.Value = vbChecked)
    
    lblRefreshList_Slowest_CheckInterval.Visible = (chkRefreshList_Enabled.Value = vbChecked)
    txtRefreshList_Slower_CheckInterval.Visible = (chkRefreshList_Enabled.Value = vbChecked)
End Sub

Private Sub chkRegistroLlamada_Save_Click()
    EnableControls
End Sub

Private Sub txtRefreshList_Fastest_CheckInterval_GotFocus()
    CSM_Control_TextBox.SelAllText txtRefreshList_Fastest_CheckInterval
End Sub

Private Sub txtRefreshList_Fastest_CheckInterval_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRefreshList_Fastest_CheckInterval_LostFocus()
    txtRefreshList_Fastest_CheckInterval.Text = Val(txtRefreshList_Fastest_CheckInterval.Text)
    If txtRefreshList_Fastest_CheckInterval.Text = 0 Then
        txtRefreshList_Fastest_CheckInterval.Text = ""
    End If
End Sub

Private Sub txtRefreshList_Slower_CheckInterval_GotFocus()
    CSM_Control_TextBox.SelAllText txtRefreshList_Slower_CheckInterval
End Sub

Private Sub txtRefreshList_Slower_CheckInterval_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRefreshList_Slower_CheckInterval_LostFocus()
    txtRefreshList_Slower_CheckInterval.Text = Val(txtRefreshList_Slower_CheckInterval.Text)
    If txtRefreshList_Slower_CheckInterval.Text = 0 Then
        txtRefreshList_Slower_CheckInterval.Text = ""
    End If
End Sub

Private Sub cboTelephonyType_Click()
    Dim Address As ITAddress
    Dim AddressCapabilities As ITAddressCapabilities
    Dim ListIndex As Long
    Dim Telephony As Telephony
    
    Set pTelephony = New Telephony
    pTelephony.TelephonyType = cboTelephonyType.Text
    pTelephony.Initialize
    If pTelephony.Initialized Then
        Select Case cboTelephonyType.Text
            Case "NONE"
                txtTelephonyLocationCityCode.Text = pTelephony.LocationCityCode
            Case "SERVER"
                cboTelephonyCallerIDIdentificacion.ListIndex = CSM_Control_ComboBox.GetListIndexByItemData(cboTelephonyCallerIDIdentificacion, pTelephony.CallerID_Identificacion_IDSucursalTelefono)
            
                txtTelephonyLocationCountryCode.Text = pTelephony.LocationCountryCode
                txtTelephonyLocationCityCode.Text = pTelephony.LocationCityCode
                txtTelephonyLocationLocalAccessCode.Text = pTelephony.LocationLocalAccessCode
                txtTelephonyLocationLongDistanceAccessCode.Text = pTelephony.LocationLongDistanceAccessCode
                optTelephonyLocationTone.Value = Not pTelephony.LocationPulse
                optTelephonyLocationPulse.Value = pTelephony.LocationPulse
            Case "COMM"
                chkTelephonyCallerIDEnabled.Value = IIf(pTelephony.CallerIDEnabled, vbChecked, vbUnchecked)
                chkTelephonyCallerIDEnabled_Click
                
                cboTelephonyCallerIDModemPort.ListIndex = CSM_Control_ComboBox.GetListIndexByItemData(cboTelephonyCallerIDModemPort, pTelephony.CallerIDModemCOMPort)
                cboTelephonyCallerIDModemInitializationString.Text = pTelephony.CallerIDModemInitializationString
            
                chkRegistroLlamada_Save.Value = IIf(pParametro.RegistroLlamada_Save, vbChecked, vbUnchecked)
                cboTelephonyCallerIDIdentificacion.ListIndex = CSM_Control_ComboBox.GetListIndexByItemData(cboTelephonyCallerIDIdentificacion, pTelephony.CallerID_Identificacion_IDSucursalTelefono)
            
                txtTelephonyLocationCountryCode.Text = pTelephony.LocationCountryCode
                txtTelephonyLocationCityCode.Text = pTelephony.LocationCityCode
                txtTelephonyLocationLocalAccessCode.Text = pTelephony.LocationLocalAccessCode
                txtTelephonyLocationLongDistanceAccessCode.Text = pTelephony.LocationLongDistanceAccessCode
                optTelephonyLocationTone.Value = Not pTelephony.LocationPulse
                optTelephonyLocationPulse.Value = pTelephony.LocationPulse
            Case "TAPI"
                cboTelephonyAddress.Clear
                Set mCTelephonyDeviceGUID = New Collection
                For Each Address In pTelephony.TAPI.Addresses
                    If Address.QueryMediaType(TAPIMEDIATYPE_DATAMODEM) Then
                        cboTelephonyAddress.AddItem Address.AddressName
                        Set AddressCapabilities = Address
                        mCTelephonyDeviceGUID.Add AddressCapabilities.AddressCapabilityString(ACS_PERMANENTDEVICEGUID)
                        If AddressCapabilities.AddressCapabilityString(ACS_PERMANENTDEVICEGUID) = pTelephony.DeviceGUID Then
                            cboTelephonyAddress.ListIndex = cboTelephonyAddress.NewIndex
                        End If
                        Set AddressCapabilities = Nothing
                    End If
                Next Address
                For ListIndex = 1 To cboTelephonyAddress.ListCount
                    If mCTelephonyDeviceGUID(ListIndex) = pTelephony.DeviceGUID Then
                        cboTelephonyAddress.ListIndex = ListIndex - 1
                        Exit For
                    End If
                Next ListIndex
                
                chkTelephonyCallerIDEnabled.Value = IIf(pTelephony.CallerIDEnabled, vbChecked, vbUnchecked)
                chkTelephonyCallerIDEnabled_Click
                
                chkRegistroLlamada_Save.Value = IIf(pParametro.RegistroLlamada_Save, vbChecked, vbUnchecked)
        End Select
    End If
    
    EnableControls
End Sub

Private Sub chkTelephonyCallerIDEnabled_Click()
    EnableControls
End Sub

Private Sub cmdCancel_Click()
    If pTelephony.TelephonyType <> pParametro.Telephony_Type Then
        Set pTelephony = New Telephony
        pTelephony.TelephonyType = pParametro.Telephony_Type
        pTelephony.Initialize
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Select Case cboTelephonyType.Text
        Case "NONE"
            If Val(txtTelephonyLocationCityCode.Text) = 0 Then
                MsgBox "Debe especificar el Código Telefónico de Area.", vbInformation, App.Title
                txtTelephonyLocationCityCode.SetFocus
                Exit Sub
            End If

            pParametro.Telephony_Type = cboTelephonyType.Text
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "Type", pParametro.Telephony_Type)
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "LocationCityCode", txtTelephonyLocationCityCode.Text)
        Case "SERVER"
            If cboTelephonyCallerIDIdentificacion.ListIndex = -1 Then
                MsgBox "Debe especificar la Identificación.", vbInformation, App.Title
                cboTelephonyCallerIDIdentificacion.SetFocus
                Exit Sub
            End If

            If Val(txtTelephonyLocationCityCode.Text) = 0 Then
                MsgBox "Debe especificar el Código Telefónico de Area.", vbInformation, App.Title
                txtTelephonyLocationCityCode.SetFocus
                Exit Sub
            End If

            pParametro.Telephony_Type = cboTelephonyType.Text
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "Type", pParametro.Telephony_Type)
            
            'LOCATION
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "LocationCountryCode", txtTelephonyLocationCountryCode.Text)
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "LocationCityCode", txtTelephonyLocationCityCode.Text)
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "LocationLocalAccessCode", txtTelephonyLocationLocalAccessCode.Text)
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "LocationLongDistanceAccessCode", txtTelephonyLocationLongDistanceAccessCode.Text)
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "LocationPulse", optTelephonyLocationPulse.Value)
            
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "CallerIDIdentificacionIDSucursalTelefono", cboTelephonyCallerIDIdentificacion.ItemData(cboTelephonyCallerIDIdentificacion.ListIndex))
        Case "COMM"
            If chkTelephonyCallerIDEnabled.Value = vbChecked Then
                If cboTelephonyCallerIDModemPort.ListIndex = -1 Then
                    MsgBox "Debe seleccionar el Módem.", vbInformation, App.Title
                    cboTelephonyCallerIDModemPort.SetFocus
                    Exit Sub
                End If
                If cboTelephonyCallerIDModemInitializationString.Text = "" Then
                    MsgBox "Debe ingresar el Comando de Inicialización del Módem.", vbInformation, App.Title
                    cboTelephonyCallerIDModemInitializationString.SetFocus
                    Exit Sub
                End If
                If chkRegistroLlamada_Save.Value = vbChecked Then
                    If cboTelephonyCallerIDIdentificacion.ListIndex = -1 Then
                        MsgBox "Debe especificar la Identificación.", vbInformation, App.Title
                        cboTelephonyCallerIDIdentificacion.SetFocus
                        Exit Sub
                    End If
                End If
            End If
            
            'LOCATION
            If Val(txtTelephonyLocationCountryCode.Text) = 0 Then
                MsgBox "Debe especificar el Código Telefónico de País.", vbInformation, App.Title
                txtTelephonyLocationCountryCode.SetFocus
                Exit Sub
            End If
            If Val(txtTelephonyLocationCityCode.Text) = 0 Then
                MsgBox "Debe especificar el Código Telefónico de Area.", vbInformation, App.Title
                txtTelephonyLocationCityCode.SetFocus
                Exit Sub
            End If
            If optTelephonyLocationTone.Value = False And optTelephonyLocationPulse.Value = False Then
                MsgBox "Debe especificar el Modo de Discado Telefónico.", vbInformation, App.Title
                Exit Sub
            End If
            
            pParametro.Telephony_Type = cboTelephonyType.Text
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "Type", pParametro.Telephony_Type)
            
            'CALLER ID
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "CallerIDEnabled", (chkTelephonyCallerIDEnabled.Value = vbChecked))
            If chkTelephonyCallerIDEnabled.Value = vbChecked Then
                Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "CallerIDModemInitializationString", cboTelephonyCallerIDModemInitializationString.Text)
                Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "CallerIDModemCOMPort", cboTelephonyCallerIDModemPort.ItemData(cboTelephonyCallerIDModemPort.ListIndex))
                Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "CallerIDSave", (chkRegistroLlamada_Save.Value = vbChecked))
                If chkRegistroLlamada_Save.Value = vbChecked Then
                    Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "CallerIDIdentificacionIDSucursalTelefono", cboTelephonyCallerIDIdentificacion.ItemData(cboTelephonyCallerIDIdentificacion.ListIndex))
                End If
            End If

            'LOCATION
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "LocationCountryCode", txtTelephonyLocationCountryCode.Text)
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "LocationCityCode", txtTelephonyLocationCityCode.Text)
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "LocationLocalAccessCode", txtTelephonyLocationLocalAccessCode.Text)
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "LocationLongDistanceAccessCode", txtTelephonyLocationLongDistanceAccessCode.Text)
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "LocationPulse", optTelephonyLocationPulse.Value)
            
        Case "TAPI"
            If cboTelephonyAddress.ListIndex = -1 Then
                MsgBox "Debe seleccionar el Módem.", vbInformation, App.Title
                cboTelephonyAddress.SetFocus
                Exit Sub
            End If
            
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "DeviceGUID", mCTelephonyDeviceGUID(cboTelephonyAddress.ListIndex + 1))
            Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "CallerIDEnabled", (chkTelephonyCallerIDEnabled.Value = vbChecked))
            If chkTelephonyCallerIDEnabled.Value = vbChecked Then
                Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Telephony", "CallerIDSave", (chkRegistroLlamada_Save.Value = vbChecked))
            End If
    End Select
    
    If datcboSucursal.BoundText = "" Then
        MsgBox "Debe especificar la Sucursal.", vbInformation, App.Title
        datcboSucursal.SetFocus
        Exit Sub
    End If
    
    If chkRefreshList_Enabled.Value = vbChecked Then
        If txtRefreshList_Fastest_CheckInterval.Text = "" Then
            MsgBox "Debe ingresar el Intervalo de Verificación Rápido para la Actualización de Listas.", vbInformation, App.Title
            txtRefreshList_Fastest_CheckInterval.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtRefreshList_Fastest_CheckInterval.Text) Then
            MsgBox "El Intervalo de Verificación Rápido para la Actualización de Listas debe ser un valor numérico.", vbInformation, App.Title
            txtRefreshList_Fastest_CheckInterval.SetFocus
            txtRefreshList_Fastest_CheckInterval_GotFocus
            Exit Sub
        End If
        
        If txtRefreshList_Slower_CheckInterval.Text = "" Then
            MsgBox "Debe ingresar el Intervalo de Verificación Lento para la Actualización de Listas.", vbInformation, App.Title
            txtRefreshList_Slower_CheckInterval.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtRefreshList_Slower_CheckInterval.Text) Then
            MsgBox "El Intervalo de Verificación Lneto para la Actualización de Listas debe ser un valor numérico.", vbInformation, App.Title
            txtRefreshList_Slower_CheckInterval.SetFocus
            txtRefreshList_Slower_CheckInterval_GotFocus
            Exit Sub
        End If
    End If
    
    Set pTelephony = New Telephony
    pTelephony.TelephonyType = pParametro.Telephony_Type
    pTelephony.Initialize
    
    pParametro.RegistroLlamada_Save = (chkRegistroLlamada_Save.Value = vbChecked)
    
    pParametro.IDSucursal = datcboSucursal.BoundText
    Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\Misc", "IDSucursal", pParametro.IDSucursal)
    
    pParametro.RefreshList_Enabled = (chkRefreshList_Enabled.Value = vbChecked)
    Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\RefreshLists", "Enabled", pParametro.RefreshList_Enabled)
    
    pParametro.RefreshList_Fastest_CheckInterval_Seconds = CLng(txtRefreshList_Fastest_CheckInterval.Text)
    Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\RefreshLists", "Fastest_CheckIntervalSeconds", pParametro.RefreshList_Fastest_CheckInterval_Seconds)
    
    pParametro.RefreshList_Slowest_CheckInterval_Seconds = CLng(txtRefreshList_Slower_CheckInterval.Text)
    Call CSM_Registry.SetValue_ToApplication_CurrentUser("Options\RefreshLists", "Slowest_CheckIntervalSeconds", pParametro.RefreshList_Slowest_CheckInterval_Seconds)
    
    pParametro.Workstation_SyncDateTimeWithDBServer_Enabled = (chkSyncDateTimeWithDBServer_Enabled.Value = vbChecked)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Options", "SyncDateTimeWithDBServer", (chkSyncDateTimeWithDBServer_Enabled.Value = vbChecked))
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim CSC_TelephonyCOM As CSC_TelephonyCOM
    Dim COMIndex As Integer
    Dim ModemIndex As Integer

    cboTelephonyType.AddItem "NONE"
    cboTelephonyType.AddItem "COMM"
    cboTelephonyType.AddItem "SERVER"
    
    Set CSC_TelephonyCOM = New CSC_TelephonyCOM
    Call CSC_TelephonyCOM.LoadPortsList
    For COMIndex = 1 To 8
        For ModemIndex = 1 To CSC_TelephonyCOM.CModemCOMPorts.Count
            If Val(Mid(CSC_TelephonyCOM.CModemCOMPorts(ModemIndex), 4)) = COMIndex Then
                cboTelephonyCallerIDModemPort.AddItem CSC_TelephonyCOM.CModemCOMPorts(ModemIndex) & " - " & CSC_TelephonyCOM.CModemNames(ModemIndex)
                cboTelephonyCallerIDModemPort.ItemData(cboTelephonyCallerIDModemPort.NewIndex) = COMIndex
                Exit For
            End If
        Next ModemIndex
        If cboTelephonyCallerIDModemPort.ListCount < COMIndex Then
            cboTelephonyCallerIDModemPort.AddItem "COM" & COMIndex
            cboTelephonyCallerIDModemPort.ItemData(cboTelephonyCallerIDModemPort.NewIndex) = COMIndex
        End If
    Next COMIndex
    
    Set CSC_TelephonyCOM = Nothing
    
    cboTelephonyCallerIDModemInitializationString.AddItem "AT#CID=1"
    cboTelephonyCallerIDModemInitializationString.AddItem "AT#CID=2"
    
    Call CSM_Control_ComboBox.FillFromSQL(cboTelephonyCallerIDIdentificacion, "SELECT SucursalTelefono.IDSucursalTelefono, Sucursal.Nombre + ' - ' + RTRIM(SucursalTelefono.TelefonoNumero) AS Display FROM Sucursal INNER JOIN SucursalTelefono ON Sucursal.IDSucursal = SucursalTelefono.IDSucursal WHERE Sucursal.Activo = 1 AND SucursalTelefono.Activo = 1 ORDER BY Sucursal.Nombre, SucursalTelefono.TelefonoNumero", "IDSucursalTelefono", "Display", "Sucursales y Teléfonos", cscpNone)
    
    cboTelephonyType.ListIndex = Switch(pParametro.Telephony_Type = "NONE", 0, pParametro.Telephony_Type = "COMM", 1, pParametro.Telephony_Type = "TAPI", 0, pParametro.Telephony_Type = "SERVER", 2)
    
    Call CSM_Control_DataCombo.FillFromSQL(datcboSucursal, "SELECT IDSucursal, Nombre FROM Sucursal WHERE Activo = 1 ORDER BY Nombre", "IDSucursal", "Nombre", "Sucursal", cscpItemOrNone, pParametro.IDSucursal)
        
    chkRefreshList_Enabled.Value = IIf(pParametro.RefreshList_Enabled, vbChecked, vbUnchecked)
    chkRefreshList_Enabled_Click
    txtRefreshList_Fastest_CheckInterval.Text = pParametro.RefreshList_Fastest_CheckInterval_Seconds
    txtRefreshList_Slower_CheckInterval.Text = pParametro.RefreshList_Slowest_CheckInterval_Seconds
    
    chkSyncDateTimeWithDBServer_Enabled.Value = IIf(pParametro.Workstation_SyncDateTimeWithDBServer_Enabled, vbChecked, vbUnchecked)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOpcionWorkstation = Nothing
End Sub

Private Sub EnableControls()
    fraTelephonyCallerID.Visible = (cboTelephonyType.Text <> "NONE")
    chkTelephonyCallerIDEnabled.Visible = (cboTelephonyType.Text = "TAPI" Or cboTelephonyType.Text = "COMM")
    
    fraTelephonyTAPI.Visible = (cboTelephonyType.Text = "TAPI")
    
    fraTelephonyCOMMCallerID.Visible = (cboTelephonyType.Text = "COMM" And chkTelephonyCallerIDEnabled.Value = vbChecked)
    
    chkRegistroLlamada_Save.Visible = ((cboTelephonyType.Text = "TAPI" Or cboTelephonyType.Text = "COMM") And chkTelephonyCallerIDEnabled.Value = vbChecked)
    
    lblTelephonyCallerIDIdentificacion.Visible = (cboTelephonyType.Text = "COMM" And chkTelephonyCallerIDEnabled.Value And chkRegistroLlamada_Save.Value = vbChecked) Or (cboTelephonyType.Text = "SERVER")
    cboTelephonyCallerIDIdentificacion.Visible = (cboTelephonyType.Text = "COMM" And chkTelephonyCallerIDEnabled.Value And chkRegistroLlamada_Save.Value = vbChecked) Or (cboTelephonyType.Text = "SERVER")
    
    fraTelephonyLocation.Visible = (cboTelephonyType.Text <> "TAPI")
    lblTelephonyLocationCountryCode.Visible = (cboTelephonyType.Text = "COMM" Or cboTelephonyType.Text = "SERVER")
    txtTelephonyLocationCountryCode.Visible = (cboTelephonyType.Text = "COMM" Or cboTelephonyType.Text = "SERVER")
    lblTelephonyLocationLocalAccessCode.Visible = (cboTelephonyType.Text = "COMM" Or cboTelephonyType.Text = "SERVER")
    txtTelephonyLocationLocalAccessCode.Visible = (cboTelephonyType.Text = "COMM" Or cboTelephonyType.Text = "SERVER")
    lblTelephonyLocationLongDistanceAccessCode.Visible = (cboTelephonyType.Text = "COMM" Or cboTelephonyType.Text = "SERVER")
    txtTelephonyLocationLongDistanceAccessCode.Visible = (cboTelephonyType.Text = "COMM" Or cboTelephonyType.Text = "SERVER")
    lblTelephonyLocationDialMode.Visible = (cboTelephonyType.Text = "COMM" Or cboTelephonyType.Text = "SERVER")
    optTelephonyLocationTone.Visible = (cboTelephonyType.Text = "COMM" Or cboTelephonyType.Text = "SERVER")
    optTelephonyLocationPulse.Visible = (cboTelephonyType.Text = "COMM" Or cboTelephonyType.Text = "SERVER")
End Sub

Private Sub txtTelephonyLocationCountryCode_GotFocus()
    CSM_Control_TextBox.SelAllText txtTelephonyLocationCountryCode
End Sub

Private Sub txtTelephonyLocationCountryCode_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTelephonyLocationCountryCode_LostFocus()
    txtTelephonyLocationCountryCode.Text = CleanNotNumericChars(txtTelephonyLocationCountryCode.Text)
End Sub

Private Sub txtTelephonyLocationCityCode_GotFocus()
    CSM_Control_TextBox.SelAllText txtTelephonyLocationCityCode
End Sub

Private Sub txtTelephonyLocationCityCode_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTelephonyLocationCityCode_LostFocus()
    txtTelephonyLocationCityCode.Text = CleanNotNumericChars(txtTelephonyLocationCityCode.Text)
End Sub

Private Sub txtTelephonyLocationLocalAccessCode_GotFocus()
    CSM_Control_TextBox.SelAllText txtTelephonyLocationLocalAccessCode
End Sub

Private Sub txtTelephonyLocationLocalAccessCode_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc("*") And KeyAscii <> Asc("#") And KeyAscii <> Asc(",") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTelephonyLocationLocalAccessCode_LostFocus()
    txtTelephonyLocationLocalAccessCode.Text = CleanInvalidCharsByAllowed(txtTelephonyLocationLocalAccessCode.Text, "0123456789*#,")
End Sub

Private Sub txtTelephonyLocationLongDistanceAccessCode_GotFocus()
    CSM_Control_TextBox.SelAllText txtTelephonyLocationLongDistanceAccessCode
End Sub

Private Sub txtTelephonyLocationLongDistanceAccessCode_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc("*") And KeyAscii <> Asc("#") And KeyAscii <> Asc(",") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTelephonyLocationLongDistanceAccessCode_LostFocus()
    txtTelephonyLocationLongDistanceAccessCode.Text = CleanInvalidCharsByAllowed(txtTelephonyLocationLongDistanceAccessCode.Text, "0123456789*#,")
End Sub
