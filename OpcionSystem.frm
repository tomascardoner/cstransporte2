VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOpcionSystem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones del Sistema"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OpcionSystem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   7605
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6240
      TabIndex        =   65
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   64
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame fraGeneral 
      Height          =   4995
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   7095
      Begin VB.CheckBox chkPermitir_Reservas_Condicionales 
         Caption         =   "Permitir Tomar Reservas Condicionales"
         Height          =   210
         Left            =   180
         TabIndex        =   18
         Top             =   4020
         Width           =   3735
      End
      Begin MSDataListLib.DataCombo datcboProvincia_ID_Predeterminada 
         Height          =   330
         Left            =   3300
         TabIndex        =   3
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datcboListaPrecio_ID_Predeterminada 
         Height          =   330
         Left            =   3300
         TabIndex        =   5
         Top             =   660
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datcboLugarGrupo_ID_Otro 
         Height          =   330
         Left            =   3300
         TabIndex        =   9
         Top             =   1620
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datcboLugar_ID_Otro 
         Height          =   330
         Left            =   3300
         TabIndex        =   11
         Top             =   1980
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datcboRuta_ID_Otra 
         Height          =   330
         Left            =   3300
         TabIndex        =   13
         Top             =   2340
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datcboTelefonoTipo_ID_Otro 
         Height          =   330
         Left            =   3300
         TabIndex        =   15
         Top             =   2700
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datcboUsuarioGrupo_ID_Predeterminado 
         Height          =   330
         Left            =   3300
         TabIndex        =   7
         Top             =   1080
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo datcboVehiculo_TelefonoTipo_ID 
         Height          =   330
         Left            =   3300
         TabIndex        =   17
         Top             =   3240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblListaPrecio_ID_Predeterminada 
         AutoSize        =   -1  'True
         Caption         =   "Lista de Precios Predeterminada:"
         Height          =   210
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   2370
      End
      Begin VB.Label lblLugarGrupo_ID_Otro 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Lugares ""Otro"":"
         Height          =   210
         Left            =   180
         TabIndex        =   8
         Top             =   1680
         Width           =   1845
      End
      Begin VB.Label lblLugar_ID_Otro 
         AutoSize        =   -1  'True
         Caption         =   "Lugar ""Otro"":"
         Height          =   210
         Left            =   180
         TabIndex        =   10
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label lblRuta_ID_Otra 
         AutoSize        =   -1  'True
         Caption         =   "Ruta ""Otra"":"
         Height          =   210
         Left            =   180
         TabIndex        =   12
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblTelefonoTipo_ID_Otro 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Teléfono ""Otro"":"
         Height          =   210
         Left            =   180
         TabIndex        =   14
         Top             =   2760
         Width           =   1725
      End
      Begin VB.Label lblUsuarioGrupo_ID_Predeterminado 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Usuarios Predeterminado:"
         Height          =   210
         Left            =   180
         TabIndex        =   6
         Top             =   1140
         Width           =   2580
      End
      Begin VB.Label lblVehiculo_TelefonoTipo_ID 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Teléfono para los Vehículos:"
         Height          =   210
         Left            =   180
         TabIndex        =   16
         Top             =   3300
         Width           =   2640
      End
      Begin VB.Label lblProvincia_ID_Predeterminada 
         AutoSize        =   -1  'True
         Caption         =   "Provincia Predeterminada:"
         Height          =   210
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   1875
      End
   End
   Begin VB.Frame fraCuentaCorriente 
      Height          =   3075
      Left            =   240
      TabIndex        =   30
      Top             =   480
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox txtCuentaCorriente_MovimientoAnterior_Minutos 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   42
         Top             =   2580
         Width           =   570
      End
      Begin MSDataListLib.DataCombo datcboCuentaCorrienteGrupo_ID_ViajeDebito 
         Height          =   330
         Left            =   3300
         TabIndex        =   32
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datcboCuentaCorrienteGrupo_ID_Transferencia 
         Height          =   330
         Left            =   3300
         TabIndex        =   38
         Top             =   1560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datcboCuentaCorrienteCaja_ID_ViajeDebito 
         Height          =   330
         Left            =   3300
         TabIndex        =   36
         Top             =   1020
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datcboCuentaCorrienteGrupo_ID_ViajeCredito 
         Height          =   330
         Left            =   3300
         TabIndex        =   34
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datcboCuentaCorrienteGrupo_ID_Sueldo 
         Height          =   330
         Left            =   3300
         TabIndex        =   40
         Top             =   2100
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblCuentaCorriente_MovimientoAnterior_Minutos 
         AutoSize        =   -1  'True
         Caption         =   "Intervalo para considerar un movimiento como Anterior:                   minutos."
         Height          =   210
         Left            =   180
         TabIndex        =   41
         Top             =   2640
         Width           =   5445
      End
      Begin VB.Label lblCuentaCorrienteGrupo_ID_Sueldo 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Cta. Cte. para Sueldos:"
         Height          =   210
         Left            =   180
         TabIndex        =   39
         Top             =   2160
         Width           =   2880
      End
      Begin VB.Label lblCuentaCorrienteGrupo_ID_ViajeDebito 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Cta. Cte. para Débito de Viajes:"
         Height          =   210
         Left            =   180
         TabIndex        =   31
         Top             =   300
         Width           =   2970
      End
      Begin VB.Label lblCuentaCorrienteGrupo_ID_Transferencia 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Cta. Cte. para Transferencias:"
         Height          =   210
         Left            =   180
         TabIndex        =   37
         Top             =   1620
         Width           =   2910
      End
      Begin VB.Label lblCuentaCorrienteCaja_ID_ViajeDebito 
         AutoSize        =   -1  'True
         Caption         =   "Caja de Cta. Cte. para Débito de Viajes:"
         Height          =   210
         Left            =   180
         TabIndex        =   35
         Top             =   1080
         Width           =   2835
      End
      Begin VB.Label lblCuentaCorrienteGrupo_ID_ViajeCredito 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Cta. Cte. para Crédito de Viajes:"
         Height          =   210
         Left            =   180
         TabIndex        =   33
         Top             =   660
         Width           =   3030
      End
   End
   Begin VB.Frame fraPlanillaViajeEmail 
      Height          =   4455
      Left            =   240
      TabIndex        =   43
      Top             =   480
      Width           =   7095
      Begin VB.ComboBox cboPlanillaViajeEmail_SendReportFormat 
         Height          =   330
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txtPlanillaViajeEmail_SMTPHost 
         Height          =   315
         Left            =   2820
         MaxLength       =   50
         TabIndex        =   48
         Top             =   1380
         Width           =   4095
      End
      Begin VB.TextBox txtPlanillaViajeEmail_SMTPUserName 
         Height          =   315
         Left            =   2820
         MaxLength       =   50
         TabIndex        =   50
         Top             =   1800
         Width           =   4095
      End
      Begin VB.TextBox txtPlanillaViajeEmail_SMTPPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2820
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   52
         Top             =   2220
         Width           =   4095
      End
      Begin VB.CheckBox chkPlanillaViajeEmail_SendExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "Enviar en Formato Microsoft Excel:"
         Height          =   210
         Left            =   150
         TabIndex        =   46
         Top             =   780
         Width           =   2865
      End
      Begin VB.Label lblPlanillaViajeEmail_SendReportFormat 
         AutoSize        =   -1  'True
         Caption         =   "Enviar Reporte exportado a:"
         Height          =   210
         Left            =   180
         TabIndex        =   44
         Top             =   420
         Width           =   2025
      End
      Begin VB.Label lblPlanillaViajeEmail_SMTPHost 
         AutoSize        =   -1  'True
         Caption         =   "Servidor SMTP:"
         Height          =   210
         Left            =   180
         TabIndex        =   47
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label lblPlanillaViajeEmail_SMTPUserName 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   210
         Left            =   180
         TabIndex        =   49
         Top             =   1860
         Width           =   600
      End
      Begin VB.Label lblPlanillaViajeEmail_SMTPPassword 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   210
         Left            =   180
         TabIndex        =   51
         Top             =   2280
         Width           =   795
      End
   End
   Begin VB.Frame fraInternetDataEmail 
      Height          =   2655
      Left            =   240
      TabIndex        =   53
      Top             =   480
      Width           =   7095
      Begin VB.TextBox txtInternetDataEmail_SenderAddress 
         Height          =   315
         Left            =   2820
         MaxLength       =   100
         TabIndex        =   57
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtInternetDataEmail_SenderDisplayName 
         Height          =   315
         Left            =   2820
         MaxLength       =   100
         TabIndex        =   55
         Top             =   300
         Width           =   4095
      End
      Begin VB.TextBox txtInternetDataEmail_SMTPPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2820
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   63
         Top             =   2160
         Width           =   4095
      End
      Begin VB.TextBox txtInternetDataEmail_SMTPUserName 
         Height          =   315
         Left            =   2820
         MaxLength       =   50
         TabIndex        =   61
         Top             =   1740
         Width           =   4095
      End
      Begin VB.TextBox txtInternetDataEmail_SMTPHost 
         Height          =   315
         Left            =   2820
         MaxLength       =   50
         TabIndex        =   59
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Label lblInternetDataEmail_SenderAddress 
         AutoSize        =   -1  'True
         Caption         =   "E-mail del Remitente de Correo:"
         Height          =   210
         Left            =   180
         TabIndex        =   56
         Top             =   780
         Width           =   2235
      End
      Begin VB.Label lblInternetDataEmail_SenderDisplayName 
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Remitente de Correo:"
         Height          =   210
         Left            =   180
         TabIndex        =   54
         Top             =   360
         Width           =   2370
      End
      Begin VB.Label lblInternetDataEmail_SMTPPassword 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   210
         Left            =   240
         TabIndex        =   62
         Top             =   2220
         Width           =   795
      End
      Begin VB.Label lblInternetDataEmail_SMTPUserName 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   210
         Left            =   240
         TabIndex        =   60
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label lblInternetDataEmail_SMTPHost 
         AutoSize        =   -1  'True
         Caption         =   "Servidor SMTP:"
         Height          =   210
         Left            =   240
         TabIndex        =   58
         Top             =   1380
         Width           =   1110
      End
   End
   Begin VB.Frame fraPersona 
      Height          =   4995
      Left            =   240
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Frame fraPersonaDatosIncompletos 
         Caption         =   "Datos incompletos:"
         Height          =   1815
         Left            =   120
         TabIndex        =   20
         Top             =   180
         Width           =   3375
         Begin VB.Frame fraPersonaDatosIncompletos_Campos 
            Caption         =   "Datos a verificar:"
            Height          =   975
            Left            =   1680
            TabIndex        =   27
            Top             =   660
            Width           =   1515
            Begin VB.CheckBox chkPersona_DatoIncompleto_Domicilio 
               Caption         =   "Domicilio"
               Height          =   210
               Left            =   120
               TabIndex        =   29
               Top             =   600
               Width           =   1275
            End
            Begin VB.CheckBox chkPersona_DatoIncompleto_Documento 
               Caption         =   "Documento"
               Height          =   210
               Left            =   120
               TabIndex        =   28
               Top             =   300
               Width           =   1275
            End
         End
         Begin VB.Frame fraPersonaDatosIncompletos_Registros 
            Caption         =   "Personas:"
            Height          =   975
            Left            =   180
            TabIndex        =   24
            Top             =   660
            Width           =   1335
            Begin VB.OptionButton optPersona_DatoIncompleto_Nuevas 
               Caption         =   "Nuevas"
               Height          =   210
               Left            =   120
               TabIndex        =   26
               Top             =   600
               Width           =   855
            End
            Begin VB.OptionButton optPersona_DatoIncompleto_Todas 
               Caption         =   "Todas"
               Height          =   210
               Left            =   120
               TabIndex        =   25
               Top             =   300
               Width           =   855
            End
         End
         Begin VB.OptionButton optPersona_DatoIncompleto_Avisar 
            Caption         =   "Avisar"
            Height          =   210
            Left            =   1260
            TabIndex        =   22
            Top             =   300
            Width           =   855
         End
         Begin VB.OptionButton optPersona_DatoIncompleto_Exigir 
            Caption         =   "Exigir"
            Height          =   210
            Left            =   2340
            TabIndex        =   23
            Top             =   300
            Width           =   855
         End
         Begin VB.OptionButton optPersona_DatoIncompleto_Ignorar 
            Caption         =   "Ignorar"
            Height          =   210
            Left            =   180
            TabIndex        =   21
            Top             =   300
            Width           =   855
         End
      End
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   5475
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9657
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "GENERAL"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Personas"
            Key             =   "PERSONA"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cuenta Corriente"
            Key             =   "CUENTA_CORRIENTE"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Planillas por E-mail"
            Key             =   "PLANILLAVIAJEEMAIL"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOpcionSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim DES As CSC_Encryption_DES
    
    If datcboProvincia_ID_Predeterminada.BoundText = "" Then
        MsgBox "Debe seleccionar la Provincia Predeterminada.", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("GENERAL")
        datcboProvincia_ID_Predeterminada.SetFocus
        Exit Sub
    End If
    If Val(datcboListaPrecio_ID_Predeterminada.BoundText) = 0 Then
        MsgBox "Debe seleccionar la Lista de Precios Predeterminada.", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("GENERAL")
        datcboListaPrecio_ID_Predeterminada.SetFocus
        Exit Sub
    End If
    If Val(datcboLugarGrupo_ID_Otro.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo de Lugares ""GENERAL"".", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("GENERAL")
        datcboLugarGrupo_ID_Otro.SetFocus
        Exit Sub
    End If
    If Val(datcboLugar_ID_Otro.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Lugar ""GENERAL"".", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("GENERAL")
        datcboLugar_ID_Otro.SetFocus
        Exit Sub
    End If
    If datcboRuta_ID_Otra.BoundText = "" Then
        MsgBox "Debe seleccionar la Ruta ""Otra"".", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("GENERAL")
        datcboRuta_ID_Otra.SetFocus
        Exit Sub
    End If
    If Val(datcboTelefonoTipo_ID_Otro.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Tipo de Teléfono ""GENERAL"".", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("GENERAL")
        datcboTelefonoTipo_ID_Otro.SetFocus
        Exit Sub
    End If
    If Val(datcboVehiculo_TelefonoTipo_ID.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Tipo de Teléfono para los Vehículos.", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("GENERAL")
        datcboVehiculo_TelefonoTipo_ID.SetFocus
        Exit Sub
    End If
    If Val(datcboUsuarioGrupo_ID_Predeterminado.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo de Usuarios Predeterminado.", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("GENERAL")
        datcboUsuarioGrupo_ID_Predeterminado.SetFocus
        Exit Sub
    End If
    
    If Val(datcboCuentaCorrienteGrupo_ID_ViajeDebito.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo de Cuenta Corriente para el Débito de Viajes.", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("CUENTA_CORRIENTE")
        datcboCuentaCorrienteGrupo_ID_ViajeDebito.SetFocus
        Exit Sub
    End If
    If Val(datcboCuentaCorrienteGrupo_ID_ViajeCredito.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo de Cuenta Corriente para el Crédito de Viajes.", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("CUENTA_CORRIENTE")
        datcboCuentaCorrienteGrupo_ID_ViajeCredito.SetFocus
        Exit Sub
    End If
    If Val(datcboCuentaCorrienteCaja_ID_ViajeDebito.BoundText) = 0 Then
        MsgBox "Debe seleccionar la Caja de Cuenta Corriente para Débito de Viajes.", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("CUENTA_CORRIENTE")
        datcboCuentaCorrienteCaja_ID_ViajeDebito.SetFocus
        Exit Sub
    End If
    If Val(datcboCuentaCorrienteGrupo_ID_Transferencia.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo de Cuenta Corriente para Transferencias.", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("CUENTA_CORRIENTE")
        datcboCuentaCorrienteGrupo_ID_Transferencia.SetFocus
        Exit Sub
    End If
    If Val(datcboCuentaCorrienteGrupo_ID_Sueldo.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo de Cuenta Corriente para Sueldos.", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("CUENTA_CORRIENTE")
        datcboCuentaCorrienteGrupo_ID_Sueldo.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtCuentaCorriente_MovimientoAnterior_Minutos.Text) Then
        MsgBox "El Intervalo para considerar un movimiento como Anterior debe ser un valor numérico.", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("CUENTA_CORRIENTE")
        txtCuentaCorriente_MovimientoAnterior_Minutos.SetFocus
        txtCuentaCorriente_MovimientoAnterior_Minutos_GotFocus
        Exit Sub
    End If
    If cboPlanillaViajeEmail_SendReportFormat.ListIndex = -1 Then
        MsgBox "Debe especificar el Formato del Reporte a exportar.", vbInformation, App.Title
        Set tabMain.SelectedItem = tabMain.Tabs("PLANILLAVIAJEEMAIL")
        cboPlanillaViajeEmail_SendReportFormat.SetFocus
        Exit Sub
    End If
    
    'GENERAL
    pParametro.Provincia_ID_Predeterminada = datcboProvincia_ID_Predeterminada.BoundText
    pParametro.ListaPrecio_ID_Predeterminada = Val(datcboListaPrecio_ID_Predeterminada.BoundText)
    pParametro.LugarGrupo_ID_Otro = Val(datcboLugarGrupo_ID_Otro.BoundText)
    pParametro.Lugar_ID_Otro = Val(datcboLugar_ID_Otro.BoundText)
    pParametro.Ruta_ID_Otra = datcboRuta_ID_Otra.BoundText
    pParametro.TelefonoTipo_ID_Otro = Val(datcboTelefonoTipo_ID_Otro.BoundText)
    pParametro.Vehiculo_TelefonoTipo_ID = Val(datcboVehiculo_TelefonoTipo_ID.BoundText)
    pParametro.UsuarioGrupo_ID_Predeterminado = Val(datcboUsuarioGrupo_ID_Predeterminado.BoundText)
    pParametro.Permitir_Reservas_Condicionales = (chkPermitir_Reservas_Condicionales.Value = vbChecked)
    
    'PERSONA
    pParametro.Persona_DatoIncompleto_AvisoTipo = Switch(optPersona_DatoIncompleto_Ignorar.Value, 0, optPersona_DatoIncompleto_Avisar.Value, 1, optPersona_DatoIncompleto_Exigir.Value, 2)
    pParametro.Persona_DatoIncompleto_RegistroTodos = optPersona_DatoIncompleto_Todas.Value
    pParametro.Persona_DatoIncompleto_CampoDocumento = (chkPersona_DatoIncompleto_Documento.Value = vbChecked)
    pParametro.Persona_DatoIncompleto_CampoDomicilio = (chkPersona_DatoIncompleto_Domicilio.Value = vbChecked)
    
    pParametro.CuentaCorrienteGrupo_ID_ViajeDebito = Val(datcboCuentaCorrienteGrupo_ID_ViajeDebito.BoundText)
    pParametro.CuentaCorrienteGrupo_ID_ViajeCredito = Val(datcboCuentaCorrienteGrupo_ID_ViajeCredito.BoundText)
    pParametro.CuentaCorrienteCaja_ID_ViajeDebito = Val(datcboCuentaCorrienteCaja_ID_ViajeDebito.BoundText)
    pParametro.CuentaCorrienteGrupo_ID_Transferencia = Val(datcboCuentaCorrienteGrupo_ID_Transferencia.BoundText)
    pParametro.CuentaCorrienteGrupo_ID_Sueldo = Val(datcboCuentaCorrienteGrupo_ID_Sueldo.BoundText)
    pParametro.CuentaCorriente_MovimientoAnterior_Minutos = CLng(txtCuentaCorriente_MovimientoAnterior_Minutos.Text)
    
    pParametro.PlanillaViajeEmail_SendReportFormat = cboPlanillaViajeEmail_SendReportFormat.ItemData(cboPlanillaViajeEmail_SendReportFormat.ListIndex)
    pParametro.PlanillaViajeEmail_SendExcel = (chkPlanillaViajeEmail_SendExcel.Value = vbChecked)
    pParametro.PlanillaViajeEmail_SMTPHost = txtPlanillaViajeEmail_SMTPHost.Text
    pParametro.PlanillaViajeEmail_SMTPUserName = txtPlanillaViajeEmail_SMTPUserName.Text
    pParametro.PlanillaViajeEmail_SMTPPassword = txtPlanillaViajeEmail_SMTPPassword.Text
    
    pParametro.InternetDataEmail_SenderDisplayName = txtInternetDataEmail_SenderDisplayName.Text
    pParametro.InternetDataEmail_SenderAddress = txtInternetDataEmail_SenderAddress.Text
    pParametro.InternetDataEmail_SMTPHost = txtInternetDataEmail_SMTPHost.Text
    pParametro.InternetDataEmail_SMTPUserName = txtInternetDataEmail_SMTPUserName.Text
    pParametro.InternetDataEmail_SMTPPassword = txtInternetDataEmail_SMTPPassword.Text
    
    'GENERAL
    Call pParametro.SaveSystemParameterText("Provincia_ID_Predeterminada", pParametro.Provincia_ID_Predeterminada)
    Call pParametro.SaveSystemParameterNumberInteger("ListaPrecio_ID_Predeterminada", pParametro.ListaPrecio_ID_Predeterminada)
    Call pParametro.SaveSystemParameterNumberInteger("LugarGrupo_ID_Otro", pParametro.LugarGrupo_ID_Otro)
    Call pParametro.SaveSystemParameterNumberInteger("Lugar_ID_Otro", pParametro.Lugar_ID_Otro)
    Call pParametro.SaveSystemParameterText("Ruta_ID_Otra", pParametro.Ruta_ID_Otra)
    Call pParametro.SaveSystemParameterNumberInteger("TelefonoTipo_ID_Otro", pParametro.TelefonoTipo_ID_Otro)
    Call pParametro.SaveSystemParameterNumberInteger("Vehiculo_TelefonoTipo_ID", pParametro.Vehiculo_TelefonoTipo_ID)
    Call pParametro.SaveSystemParameterNumberInteger("UsuarioGrupo_ID_Predeterminado", pParametro.UsuarioGrupo_ID_Predeterminado)
    Call pParametro.SaveSystemParameterBoolean("Permitir_Reservas_Condicionales", pParametro.Permitir_Reservas_Condicionales)
    
    'PERSONA
    Call pParametro.SaveSystemParameterNumberInteger("Persona_DatoIncompleto_AvisoTipo", pParametro.Persona_DatoIncompleto_AvisoTipo)
    Call pParametro.SaveSystemParameterBoolean("Persona_DatoIncompleto_RegistroTodos", pParametro.Persona_DatoIncompleto_RegistroTodos)
    Call pParametro.SaveSystemParameterBoolean("Persona_DatoIncompleto_CampoDocumento", pParametro.Persona_DatoIncompleto_CampoDocumento)
    Call pParametro.SaveSystemParameterBoolean("Persona_DatoIncompleto_CampoDomicilio", pParametro.Persona_DatoIncompleto_CampoDomicilio)
    
    Call pParametro.SaveSystemParameterNumberInteger("CuentaCorrienteGrupo_ID_ViajeDebito", pParametro.CuentaCorrienteGrupo_ID_ViajeDebito)
    Call pParametro.SaveSystemParameterNumberInteger("CuentaCorrienteGrupo_ID_ViajeCredito", pParametro.CuentaCorrienteGrupo_ID_ViajeCredito)
    Call pParametro.SaveSystemParameterNumberInteger("CuentaCorrienteCaja_ID_ViajeDebito", pParametro.CuentaCorrienteCaja_ID_ViajeDebito)
    Call pParametro.SaveSystemParameterNumberInteger("CuentaCorrienteGrupo_ID_Transferencia", pParametro.CuentaCorrienteGrupo_ID_Transferencia)
    Call pParametro.SaveSystemParameterNumberInteger("CuentaCorrienteGrupo_ID_Sueldo", pParametro.CuentaCorrienteGrupo_ID_Sueldo)
    Call pParametro.SaveSystemParameterNumberInteger("CuentaCorriente_MovimientoAnterior_Minutos", pParametro.CuentaCorriente_MovimientoAnterior_Minutos)
    
    Call pParametro.SaveSystemParameterNumberInteger("PlanillaViajeEmail_SendReportFormat", pParametro.PlanillaViajeEmail_SendReportFormat)
    Call pParametro.SaveSystemParameterBoolean("PlanillaViajeEmail_SendExcel", pParametro.PlanillaViajeEmail_SendExcel)
    Call pParametro.SaveSystemParameterText("PlanillaViajeEmail_SMTPHost", pParametro.PlanillaViajeEmail_SMTPHost)
    Call pParametro.SaveSystemParameterText("PlanillaViajeEmail_SMTPUserName", pParametro.PlanillaViajeEmail_SMTPUserName)
    If pParametro.PlanillaViajeEmail_SMTPUserName = "" Then
        Call pParametro.SaveSystemParameterText("PlanillaViajeEmail_SMTPPassword", "")
    Else
        Set DES = New CSC_Encryption_DES
        Call pParametro.SaveSystemParameterText("PlanillaViajeEmail_SMTPPassword", DES.EncryptString(pParametro.PlanillaViajeEmail_SMTPPassword, PASSWORD_ENCRYPTION_KEY, False))
        Set DES = Nothing
    End If
    
    Call pParametro.SaveSystemParameterText("InternetDataEmail_SenderDisplayName", pParametro.InternetDataEmail_SenderDisplayName)
    Call pParametro.SaveSystemParameterText("InternetDataEmail_SenderAddress", pParametro.InternetDataEmail_SenderAddress)
    Call pParametro.SaveSystemParameterText("InternetDataEmail_SMTPHost", pParametro.InternetDataEmail_SMTPHost)
    Call pParametro.SaveSystemParameterText("InternetDataEmail_SMTPUserName", pParametro.InternetDataEmail_SMTPUserName)
    If pParametro.InternetDataEmail_SMTPUserName = "" Then
        Call pParametro.SaveSystemParameterText("InternetDataEmail_SMTPPassword", "")
    Else
        Set DES = New CSC_Encryption_DES
        Call pParametro.SaveSystemParameterText("InternetDataEmail_SMTPPassword", DES.EncryptString(pParametro.InternetDataEmail_SMTPPassword, PASSWORD_ENCRYPTION_KEY, False))
        Set DES = Nothing
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    'GENERAL
    CSM_Control_DataCombo.FillFromSQL datcboProvincia_ID_Predeterminada, "SELECT IDProvincia, Nombre FROM Provincia ORDER BY Nombre", "IDProvincia", "Nombre", "Provincias", cscpItemOrNone, pParametro.Provincia_ID_Predeterminada
    CSM_Control_DataCombo.FillFromSQL datcboListaPrecio_ID_Predeterminada, "SELECT IDListaPrecio, Nombre FROM ListaPrecio WHERE Activo = 1 ORDER BY Nombre", "IDListaPrecio", "Nombre", "Listas de Precios", cscpItemOrNone, pParametro.ListaPrecio_ID_Predeterminada
    CSM_Control_DataCombo.FillFromSQL datcboLugarGrupo_ID_Otro, "SELECT IDLugarGrupo, Nombre FROM LugarGrupo WHERE Activo = 1 ORDER BY Nombre", "IDLugarGrupo", "Nombre", "Grupos de Lugares", cscpItemOrNone, pParametro.LugarGrupo_ID_Otro
    CSM_Control_DataCombo.FillFromSQL datcboLugar_ID_Otro, "SELECT IDLugar, Nombre FROM Lugar WHERE Activo = 1 ORDER BY Nombre", "IDLugar", "Nombre", "Lugares", cscpItemOrNone, pParametro.Lugar_ID_Otro
    CSM_Control_DataCombo.FillFromSQL datcboRuta_ID_Otra, "SELECT RTRIM(IDRuta) AS IDRuta FROM Ruta WHERE Activo = 1 ORDER BY IDRuta", "IDRuta", "IDRuta", "Rutas", cscpItemOrNone, pParametro.Ruta_ID_Otra
    CSM_Control_DataCombo.FillFromSQL datcboTelefonoTipo_ID_Otro, "SELECT IDTelefonoTipo, Nombre FROM TelefonoTipo WHERE Activo = 1 ORDER BY Nombre", "IDTelefonoTipo", "Nombre", "Tipos de Teléfono", cscpItemOrNone, pParametro.TelefonoTipo_ID_Otro
    CSM_Control_DataCombo.FillFromSQL datcboVehiculo_TelefonoTipo_ID, "SELECT IDTelefonoTipo, Nombre FROM TelefonoTipo WHERE Activo = 1 ORDER BY Nombre", "IDTelefonoTipo", "Nombre", "Tipos de Teléfono", cscpItemOrNone, pParametro.Vehiculo_TelefonoTipo_ID
    CSM_Control_DataCombo.FillFromSQL datcboUsuarioGrupo_ID_Predeterminado, "SELECT IDUsuarioGrupo, Nombre FROM UsuarioGrupo WHERE Activo = 1 OR IDUsuarioGrupo = " & pParametro.UsuarioGrupo_ID_Predeterminado, "IDUsuarioGrupo", "Nombre", "Grupos de Usuarios", cscpItemOrNone, pParametro.UsuarioGrupo_ID_Predeterminado
    chkPermitir_Reservas_Condicionales.Value = IIf(pParametro.Permitir_Reservas_Condicionales, vbChecked, vbUnchecked)
    
    'PERSONA
    optPersona_DatoIncompleto_Ignorar.Value = (pParametro.Persona_DatoIncompleto_AvisoTipo = 0)
    optPersona_DatoIncompleto_Avisar.Value = (pParametro.Persona_DatoIncompleto_AvisoTipo = 1)
    optPersona_DatoIncompleto_Exigir.Value = (pParametro.Persona_DatoIncompleto_AvisoTipo = 2)
    
    optPersona_DatoIncompleto_Todas.Value = pParametro.Persona_DatoIncompleto_RegistroTodos
    optPersona_DatoIncompleto_Nuevas.Value = Not pParametro.Persona_DatoIncompleto_RegistroTodos
    
    chkPersona_DatoIncompleto_Documento.Value = IIf(pParametro.Persona_DatoIncompleto_CampoDocumento, vbChecked, vbUnchecked)
    chkPersona_DatoIncompleto_Domicilio.Value = IIf(pParametro.Persona_DatoIncompleto_CampoDomicilio, vbChecked, vbUnchecked)

    CSM_Control_DataCombo.FillFromSQL datcboCuentaCorrienteGrupo_ID_ViajeDebito, "SELECT IDCuentaCorrienteGrupo, Nombre FROM CuentaCorrienteGrupo WHERE Activo = 1 ORDER BY Nombre", "IDCuentaCorrienteGrupo", "Nombre", "Grupos de Cuenta Corriente", cscpItemOrNone, pParametro.CuentaCorrienteGrupo_ID_ViajeDebito
    CSM_Control_DataCombo.FillFromSQL datcboCuentaCorrienteGrupo_ID_ViajeCredito, "SELECT IDCuentaCorrienteGrupo, Nombre FROM CuentaCorrienteGrupo WHERE Activo = 1 ORDER BY Nombre", "IDCuentaCorrienteGrupo", "Nombre", "Grupos de Cuenta Corriente", cscpItemOrNone, pParametro.CuentaCorrienteGrupo_ID_ViajeCredito
    CSM_Control_DataCombo.FillFromSQL datcboCuentaCorrienteCaja_ID_ViajeDebito, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas de Cuenta Corriente", cscpItemOrNone, pParametro.CuentaCorrienteCaja_ID_ViajeDebito
    CSM_Control_DataCombo.FillFromSQL datcboCuentaCorrienteGrupo_ID_Transferencia, "SELECT IDCuentaCorrienteGrupo, Nombre FROM CuentaCorrienteGrupo WHERE Activo = 1 ORDER BY Nombre", "IDCuentaCorrienteGrupo", "Nombre", "Grupos de Cuenta Corriente", cscpItemOrNone, pParametro.CuentaCorrienteGrupo_ID_Transferencia
    CSM_Control_DataCombo.FillFromSQL datcboCuentaCorrienteGrupo_ID_Sueldo, "SELECT IDCuentaCorrienteGrupo, Nombre FROM CuentaCorrienteGrupo WHERE Activo = 1 ORDER BY Nombre", "IDCuentaCorrienteGrupo", "Nombre", "Grupos de Cuenta Corriente", cscpItemOrNone, pParametro.CuentaCorrienteGrupo_ID_Sueldo
    txtCuentaCorriente_MovimientoAnterior_Minutos.Text = pParametro.CuentaCorriente_MovimientoAnterior_Minutos
        
    'FORMATO DEL REPORTE
    CSM_Control_ComboBox.AddItemWithItemData cboPlanillaViajeEmail_SendReportFormat, "-- No exportar --", crEFTNoFormat
    CSM_Control_ComboBox.AddItemWithItemData cboPlanillaViajeEmail_SendReportFormat, "Adobe PDF", crEFTPortableDocFormat
    CSM_Control_ComboBox.AddItemWithItemData cboPlanillaViajeEmail_SendReportFormat, "Crystal Reports", crEFTCrystalReport
    CSM_Control_ComboBox.AddItemWithItemData cboPlanillaViajeEmail_SendReportFormat, "HTML 4.0", crEFTHTML40
    CSM_Control_ComboBox.AddItemWithItemData cboPlanillaViajeEmail_SendReportFormat, "Microsoft Excel 97", crEFTExcel97
    CSM_Control_ComboBox.AddItemWithItemData cboPlanillaViajeEmail_SendReportFormat, "Microsoft Word", crEFTWordForWindows
    CSM_Control_ComboBox.AddItemWithItemData cboPlanillaViajeEmail_SendReportFormat, "Rich Text Format", crEFTExactRichText
    CSM_Control_ComboBox.AddItemWithItemData cboPlanillaViajeEmail_SendReportFormat, "XML", crEFTXML
    cboPlanillaViajeEmail_SendReportFormat.ListIndex = CSM_Control_ComboBox.GetListIndexByItemData(cboPlanillaViajeEmail_SendReportFormat, pParametro.PlanillaViajeEmail_SendReportFormat, cscpItemOrFirst)
    
    chkPlanillaViajeEmail_SendExcel.Value = IIf(pParametro.PlanillaViajeEmail_SendExcel, vbChecked, vbUnchecked)
    txtPlanillaViajeEmail_SMTPHost.Text = pParametro.PlanillaViajeEmail_SMTPHost
    txtPlanillaViajeEmail_SMTPUserName.Text = pParametro.PlanillaViajeEmail_SMTPUserName
    txtPlanillaViajeEmail_SMTPPassword.Text = pParametro.PlanillaViajeEmail_SMTPPassword

    'INTERNET DATA POR MAIL
    txtInternetDataEmail_SenderDisplayName.Text = pParametro.InternetDataEmail_SenderDisplayName
    txtInternetDataEmail_SenderAddress.Text = pParametro.InternetDataEmail_SenderAddress
    txtInternetDataEmail_SMTPHost.Text = pParametro.InternetDataEmail_SMTPHost
    txtInternetDataEmail_SMTPUserName.Text = pParametro.InternetDataEmail_SMTPUserName
    txtInternetDataEmail_SMTPPassword.Text = pParametro.InternetDataEmail_SMTPPassword
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOpcionSystem = Nothing
End Sub

Private Sub tabMain_Click()
    fraGeneral.Visible = (tabMain.SelectedItem.Key = "GENERAL")
    fraPersona.Visible = (tabMain.SelectedItem.Key = "PERSONA")
    fraCuentaCorriente.Visible = (tabMain.SelectedItem.Key = "CUENTA_CORRIENTE")
    fraPlanillaViajeEmail.Visible = (tabMain.SelectedItem.Key = "PLANILLAVIAJEEMAIL")
    fraInternetDataEmail.Visible = (tabMain.SelectedItem.Key = "INTERNETDATAEMAIL")
End Sub

Private Sub txtInternetDataEmail_SenderAddress_GotFocus()
    CSM_Control_TextBox.SelAllText txtInternetDataEmail_SenderAddress
End Sub

Private Sub txtInternetDataEmail_SenderDisplayName_GotFocus()
    CSM_Control_TextBox.SelAllText txtInternetDataEmail_SenderDisplayName
End Sub

Private Sub txtInternetDataEmail_SMTPHost_GotFocus()
    CSM_Control_TextBox.SelAllText txtInternetDataEmail_SMTPHost
End Sub

Private Sub txtInternetDataEmail_SMTPPassword_GotFocus()
    CSM_Control_TextBox.SelAllText txtInternetDataEmail_SMTPPassword
End Sub

Private Sub txtInternetDataEmail_SMTPUserName_GotFocus()
    CSM_Control_TextBox.SelAllText txtInternetDataEmail_SMTPUserName
End Sub

Private Sub txtPlanillaViajeEmail_SMTPHost_GotFocus()
    CSM_Control_TextBox.SelAllText txtPlanillaViajeEmail_SMTPHost
End Sub

Private Sub txtPlanillaViajeEmail_SMTPPassword_GotFocus()
    CSM_Control_TextBox.SelAllText txtPlanillaViajeEmail_SMTPPassword
End Sub

Private Sub txtPlanillaViajeEmail_SMTPUserName_GotFocus()
    CSM_Control_TextBox.SelAllText txtPlanillaViajeEmail_SMTPUserName
End Sub

Private Sub txtCuentaCorriente_MovimientoAnterior_Minutos_GotFocus()
    CSM_Control_TextBox.SelAllText txtCuentaCorriente_MovimientoAnterior_Minutos
End Sub

Private Sub txtCuentaCorriente_MovimientoAnterior_Minutos_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCuentaCorriente_MovimientoAnterior_Minutos_LostFocus()
    txtCuentaCorriente_MovimientoAnterior_Minutos.Text = Val(txtCuentaCorriente_MovimientoAnterior_Minutos.Text)
    If txtCuentaCorriente_MovimientoAnterior_Minutos.Text = 0 Then
        txtCuentaCorriente_MovimientoAnterior_Minutos.Text = ""
    End If
End Sub
