VERSION 5.00
Begin VB.Form frmApplicationInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información de la Aplicacion"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ApplicationInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4500
      Width           =   1275
   End
   Begin VB.ListBox lstInfo 
      Height          =   4260
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4875
   End
End
Attribute VB_Name = "frmApplicationInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Private Type Size
        cx As Long
        cy As Long
End Type

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim CategorySeparator As String
    Dim CategorySeparatorSize As Size
    
    Const CATEGORY_SEPARATOR_CHAR = "="
    Const SIZE_SCREEN_PROPORTION = 0.8
    Const CONTROL_SPACE = 60
    
    'RESIZE
    Height = Screen.Height * SIZE_SCREEN_PROPORTION
    Width = Screen.Width * SIZE_SCREEN_PROPORTION
    
    lstInfo.Top = CONTROL_SPACE
    lstInfo.Left = CONTROL_SPACE
    lstInfo.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lstInfo.Height = ScaleHeight - lstInfo.Top - CONTROL_SPACE - cmdOK.Height - CONTROL_SPACE
    cmdOK.Top = ScaleHeight - CONTROL_SPACE - cmdOK.Height
    cmdOK.Left = ScaleWidth - CONTROL_SPACE - cmdOK.Width
    
    'CALCULATE SEPARATOR WIDTH
    Call GetTextExtentPoint32(Me.hdc, CATEGORY_SEPARATOR_CHAR, Len(CATEGORY_SEPARATOR_CHAR), CategorySeparatorSize)
    CategorySeparator = String(((lstInfo.Width / Screen.TwipsPerPixelX) / CategorySeparatorSize.cx) - 1, CATEGORY_SEPARATOR_CHAR)
        
    With lstInfo
        'COMPUTER
        .AddItem "THIS COMPUTER:"
        .AddItem "--------------"
        .AddItem "Name:             " & CSM_Session.GetComputerName()
        .AddItem "Local Username:   " & CSM_Session.GetUserName()
        .AddItem "Network Username: " & CSM_Session.GetNetworkUserName()
        'CONFIG
        .AddItem CategorySeparator
        .AddItem "CONFIGURATION:"
        .AddItem "--------------"
        .AddItem "Current:          " & pParametro.Config_Name
        Select Case pParametro.Config_Type
            Case "REG"
                .AddItem "Type:             Registry"
            Case "INI"
                .AddItem "Type:             INI File"
                .AddItem "Filename:         " & pParametro.Config_INIFile_FullPathAndName
            Case Else
                .AddItem "Type:             " & CSM_Constant.ITEM_START_CHAR & " Unknown " & CSM_Constant.ITEM_END_CHAR
        End Select
        .AddItem "Application File: " & App.Path & "\" & App.EXEName & ".exe"
        'DATABASE INFO
        .AddItem CategorySeparator
        .AddItem "DATABASE:"
        .AddItem "---------"
        .AddItem "Provider:                " & pParametro.Database_Provider
        .AddItem "Connection Timeout:      " & pParametro.Database_ConnectionTimeout
        .AddItem "Command Timeout:         " & pParametro.Database_CommandTimeout
        .AddItem "Packet Size:             " & pParametro.Database_PacketSize
        .AddItem "Data Type Compatibility: " & pParametro.Database_DataTypeCompatibility
        .AddItem "Data Source:             " & pParametro.Database_DataSource
        .AddItem "Failover Partner:        " & pParametro.Database_FailoverPartner
        .AddItem "User ID:                 " & pParametro.Database_UserID
        .AddItem "Database:                " & pParametro.Database_Database
        .AddItem "Database History:        " & pParametro.Database_DatabaseHistory
        'REPORTS
        .AddItem CategorySeparator
        .AddItem "REPORTS:"
        .AddItem "--------"
        .AddItem "Path:     " & pParametro.Report_Path
        'TELEPHONY
        .AddItem CategorySeparator
        .AddItem "TELEPHONY:"
        .AddItem "----------"
        .AddItem "Type:                      " & pTelephony.TelephonyType
        If pTelephony.TelephonyType = "TAPI" Then
            If pTelephony.Initialized Then
                .AddItem "Selected Address Name:     " & pTelephony.Address.AddressName
            End If
        End If
        .AddItem "Dial Mode:                 " & IIf(pTelephony.LocationPulse, "Pulse", "Tone")
        .AddItem "Country Code:              " & pTelephony.LocationCountryCode
        .AddItem "City Code:                 " & pTelephony.LocationCityCode
        .AddItem "Local Access Code:         " & pTelephony.LocationLocalAccessCode
        .AddItem "Long Distance Access Code: " & pTelephony.LocationLongDistanceAccessCode
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmApplicationInfo = Nothing
End Sub
