VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPersonaPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PersonaPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   9870
   Begin VB.TextBox txtPersonaACargo 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5760
      Width           =   2715
   End
   Begin VB.CommandButton cmdPersonaACargo 
      Height          =   315
      Left            =   3840
      Picture         =   "PersonaPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Buscar..."
      Top             =   5760
      Width           =   315
   End
   Begin VB.CommandButton cmdPersonaACargoClear 
      Height          =   315
      Left            =   4200
      Picture         =   "PersonaPropiedad.frx":0AD4
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Borrar"
      Top             =   5760
      Width           =   315
   End
   Begin VB.CheckBox chkHabilitadoViajar 
      Alignment       =   1  'Right Justify
      Caption         =   "Habilitado a viajar:"
      Height          =   225
      Left            =   6120
      TabIndex        =   82
      Top             =   5760
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   8760
      Picture         =   "PersonaPropiedad.frx":105E
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   60
      Width           =   975
   End
   Begin VB.CheckBox chkListaPasajero 
      Alignment       =   1  'Right Justify
      Caption         =   "Lista de Pasajeros:"
      Height          =   390
      Left            =   4800
      TabIndex        =   84
      Top             =   6120
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.ComboBox cboSueldoDia 
      Height          =   330
      Left            =   8460
      Style           =   2  'Dropdown List
      TabIndex        =   73
      Top             =   3180
      Width           =   930
   End
   Begin VB.TextBox txtSueldoImporte 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5700
      MaxLength       =   20
      TabIndex        =   71
      Top             =   3180
      Width           =   1455
   End
   Begin VB.CheckBox chkHabilitadoInternet 
      Alignment       =   1  'Right Justify
      Caption         =   "Habilitado internet:"
      Height          =   210
      Left            =   8040
      TabIndex        =   83
      Top             =   5760
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cmdEmail 
      Height          =   315
      Left            =   4200
      Picture         =   "PersonaPropiedad.frx":1688
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Enviar Mail"
      Top             =   5220
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txtTelefonoArea 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   7680
      MaxLength       =   5
      TabIndex        =   63
      Top             =   2340
      Width           =   615
   End
   Begin VB.TextBox txtTelefonoNumero 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   8280
      MaxLength       =   16
      TabIndex        =   64
      Top             =   2340
      Width           =   1095
   End
   Begin VB.TextBox txtTelefonoArea 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   7680
      MaxLength       =   5
      TabIndex        =   57
      Top             =   1980
      Width           =   615
   End
   Begin VB.TextBox txtTelefonoNumero 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   8280
      MaxLength       =   16
      TabIndex        =   58
      Top             =   1980
      Width           =   1095
   End
   Begin VB.TextBox txtTelefonoArea 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   7680
      MaxLength       =   5
      TabIndex        =   51
      Top             =   1620
      Width           =   615
   End
   Begin VB.TextBox txtTelefonoNumero 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   8280
      MaxLength       =   16
      TabIndex        =   52
      Top             =   1620
      Width           =   1095
   End
   Begin VB.TextBox txtTelefonoArea 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   7680
      MaxLength       =   5
      TabIndex        =   45
      Top             =   1260
      Width           =   615
   End
   Begin VB.TextBox txtTelefonoNumero 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   8280
      MaxLength       =   16
      TabIndex        =   46
      Top             =   1260
      Width           =   1095
   End
   Begin VB.TextBox txtTelefonoArea 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   7680
      MaxLength       =   5
      TabIndex        =   39
      Top             =   900
      Width           =   615
   End
   Begin VB.TextBox txtTelefonoNumero 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   8280
      MaxLength       =   16
      TabIndex        =   40
      Top             =   900
      Width           =   1095
   End
   Begin VB.TextBox txtEmail 
      Height          =   315
      Left            =   900
      MaxLength       =   100
      TabIndex        =   30
      Top             =   5220
      Width           =   3255
   End
   Begin VB.CheckBox chkPermiteViajarSinPagar 
      Alignment       =   1  'Right Justify
      Caption         =   "Autorizado a Viajar sin Pagar al Contado:"
      Height          =   210
      Left            =   4800
      TabIndex        =   78
      Top             =   4140
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CommandButton cmdTelefonoDial 
      Height          =   315
      Index           =   5
      Left            =   9420
      Picture         =   "PersonaPropiedad.frx":1C9A
      Style           =   1  'Graphical
      TabIndex        =   65
      TabStop         =   0   'False
      ToolTipText     =   "Llamar"
      Top             =   2340
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton cmdTelefonoDial 
      Height          =   315
      Index           =   4
      Left            =   9420
      Picture         =   "PersonaPropiedad.frx":22AC
      Style           =   1  'Graphical
      TabIndex        =   59
      TabStop         =   0   'False
      ToolTipText     =   "Llamar"
      Top             =   1980
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton cmdTelefonoDial 
      Height          =   315
      Index           =   3
      Left            =   9420
      Picture         =   "PersonaPropiedad.frx":28BE
      Style           =   1  'Graphical
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   "Llamar"
      Top             =   1620
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton cmdTelefonoDial 
      Height          =   315
      Index           =   2
      Left            =   9420
      Picture         =   "PersonaPropiedad.frx":2ED0
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      ToolTipText     =   "Llamar"
      Top             =   1260
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton cmdTelefonoDial 
      Height          =   315
      Index           =   1
      Left            =   9420
      Picture         =   "PersonaPropiedad.frx":34E2
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Llamar"
      Top             =   900
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txtDocumentoNumero 
      Height          =   315
      Left            =   1980
      MaxLength       =   15
      TabIndex        =   24
      Top             =   4140
      Width           =   1635
   End
   Begin VB.CommandButton cmdPersonaCuentaCorrienteClear 
      Height          =   315
      Left            =   9420
      Picture         =   "PersonaPropiedad.frx":3AF4
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "Borrar"
      Top             =   3720
      Width           =   315
   End
   Begin VB.CommandButton cmdPersonaCuentaCorriente 
      Height          =   315
      Left            =   9060
      Picture         =   "PersonaPropiedad.frx":407E
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Buscar..."
      Top             =   3720
      Width           =   315
   End
   Begin VB.TextBox txtPersonaCuentaCorriente 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox txtTelefonoTipoOtro 
      Height          =   315
      Index           =   5
      Left            =   6900
      MaxLength       =   15
      TabIndex        =   62
      Top             =   2340
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTelefonoTipoOtro 
      Height          =   315
      Index           =   4
      Left            =   6900
      MaxLength       =   15
      TabIndex        =   56
      Top             =   1980
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTelefonoTipoOtro 
      Height          =   315
      Index           =   3
      Left            =   6900
      MaxLength       =   15
      TabIndex        =   50
      Top             =   1620
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTelefonoTipoOtro 
      Height          =   315
      Index           =   2
      Left            =   6900
      MaxLength       =   15
      TabIndex        =   44
      Top             =   1260
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTelefonoTipoOtro 
      Height          =   315
      Index           =   1
      Left            =   6900
      MaxLength       =   15
      TabIndex        =   38
      Top             =   900
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtDomicilioCalle3 
      Height          =   315
      Left            =   900
      MaxLength       =   50
      TabIndex        =   15
      Top             =   2700
      Width           =   3615
   End
   Begin VB.TextBox txtDomicilioCalle2 
      Height          =   315
      Left            =   900
      MaxLength       =   50
      TabIndex        =   13
      Top             =   2340
      Width           =   3615
   End
   Begin VB.OptionButton optTipoAdministrativo 
      Caption         =   "Administrativo"
      Height          =   195
      Left            =   8100
      TabIndex        =   69
      Top             =   2880
      Width           =   1335
   End
   Begin VB.OptionButton optTipoConductor 
      Caption         =   "Conductor"
      Height          =   195
      Left            =   6780
      TabIndex        =   68
      Top             =   2880
      Width           =   1095
   End
   Begin VB.OptionButton optTipoCliente 
      Caption         =   "Cliente"
      Height          =   195
      Left            =   5700
      TabIndex        =   67
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtNotas 
      Height          =   1065
      Left            =   5700
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   80
      Top             =   4560
      Width           =   4035
   End
   Begin VB.TextBox txtCodigoPostal 
      Height          =   315
      Left            =   900
      MaxLength       =   8
      TabIndex        =   21
      Top             =   3780
      Width           =   1035
   End
   Begin VB.TextBox txtDomicilioDepartamento 
      Height          =   315
      Left            =   3900
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1980
      Width           =   615
   End
   Begin VB.TextBox txtDomicilioPiso 
      Height          =   315
      Left            =   2580
      MaxLength       =   10
      TabIndex        =   9
      Top             =   1980
      Width           =   615
   End
   Begin VB.TextBox txtDomicilioNumero 
      Height          =   315
      Left            =   900
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1980
      Width           =   1035
   End
   Begin VB.TextBox txtDomicilioCalle1 
      Height          =   315
      Left            =   900
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1620
      Width           =   3615
   End
   Begin VB.TextBox txtApellido 
      Height          =   315
      Left            =   900
      MaxLength       =   100
      TabIndex        =   1
      Top             =   900
      Width           =   3615
   End
   Begin VB.CheckBox chkActivo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Activo"
      Height          =   210
      Left            =   4800
      TabIndex        =   81
      Top             =   5760
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   900
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1260
      Width           =   3615
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   88
      Top             =   660
      Width           =   9615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8460
      TabIndex        =   86
      Top             =   6120
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   7080
      TabIndex        =   85
      Top             =   6120
      Width           =   1275
   End
   Begin MSDataListLib.DataCombo datcboProvincia 
      Height          =   330
      Left            =   900
      TabIndex        =   17
      Top             =   3060
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
   Begin MSDataListLib.DataCombo datcboDocumentoTipo 
      Height          =   330
      Left            =   900
      TabIndex        =   23
      Top             =   4140
      Width           =   1035
      _ExtentX        =   1826
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
   Begin MSDataListLib.DataCombo datcboTelefonoTipo 
      Height          =   330
      Index           =   1
      Left            =   5700
      TabIndex        =   37
      Top             =   900
      Width           =   1155
      _ExtentX        =   2037
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
   Begin MSDataListLib.DataCombo datcboTelefonoTipo 
      Height          =   330
      Index           =   2
      Left            =   5700
      TabIndex        =   43
      Top             =   1260
      Width           =   1155
      _ExtentX        =   2037
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
   Begin MSDataListLib.DataCombo datcboTelefonoTipo 
      Height          =   330
      Index           =   3
      Left            =   5700
      TabIndex        =   49
      Top             =   1620
      Width           =   1155
      _ExtentX        =   2037
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
   Begin MSDataListLib.DataCombo datcboTelefonoTipo 
      Height          =   330
      Index           =   4
      Left            =   5700
      TabIndex        =   55
      Top             =   1980
      Width           =   1155
      _ExtentX        =   2037
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
   Begin MSDataListLib.DataCombo datcboTelefonoTipo 
      Height          =   330
      Index           =   5
      Left            =   5700
      TabIndex        =   61
      Top             =   2340
      Width           =   1155
      _ExtentX        =   2037
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
   Begin MSDataListLib.DataCombo datcboLocalidad 
      Height          =   330
      Left            =   900
      TabIndex        =   19
      Top             =   3420
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
   Begin MSComCtl2.DTPicker dtpFechaNacimiento 
      Height          =   315
      Left            =   900
      TabIndex        =   28
      Top             =   4860
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      Format          =   59768833
      CurrentDate     =   36950
   End
   Begin MSDataListLib.DataCombo datcboCondicionIVA 
      Height          =   330
      Left            =   900
      TabIndex        =   26
      Top             =   4500
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
   Begin VB.Label lblPersonaACargo 
      AutoSize        =   -1  'True
      Caption         =   "A cargo de:"
      Height          =   210
      Left            =   120
      TabIndex        =   32
      Top             =   5820
      Width           =   855
   End
   Begin VB.Label lblSueldoDia 
      AutoSize        =   -1  'True
      Caption         =   "Día del Mes:"
      Height          =   210
      Left            =   7440
      TabIndex        =   72
      Top             =   3240
      Width           =   870
   End
   Begin VB.Line Line5 
      X1              =   4680
      X2              =   9750
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblSueldoImporte 
      AutoSize        =   -1  'True
      Caption         =   "Sueldo:"
      Height          =   210
      Left            =   4800
      TabIndex        =   70
      Top             =   3240
      Width           =   540
   End
   Begin VB.Line Line4 
      X1              =   4650
      X2              =   9720
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4650
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line2 
      X1              =   4650
      X2              =   9720
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblCondicionIVA 
      AutoSize        =   -1  'True
      Caption         =   "I.V.A.:"
      Height          =   210
      Left            =   120
      TabIndex        =   25
      Top             =   4560
      Width           =   450
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "E-mail:"
      Height          =   210
      Left            =   120
      TabIndex        =   29
      Top             =   5280
      Width           =   465
   End
   Begin VB.Label lblFechaNacimiento 
      AutoSize        =   -1  'True
      Caption         =   "F. Nacim.:"
      Height          =   210
      Left            =   120
      TabIndex        =   27
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label lblDocumento 
      AutoSize        =   -1  'True
      Caption         =   "Docum.:"
      Height          =   210
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   585
   End
   Begin VB.Label lblPersonaCuentaCorriente 
      AutoSize        =   -1  'True
      Caption         =   "Debitar a:"
      Height          =   210
      Left            =   4800
      TabIndex        =   74
      Top             =   3780
      Width           =   690
   End
   Begin VB.Label lblTelefono 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono 5:"
      Height          =   210
      Index           =   5
      Left            =   4800
      TabIndex        =   60
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label lblTelefono 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono 4:"
      Height          =   210
      Index           =   4
      Left            =   4800
      TabIndex        =   54
      Top             =   2040
      Width           =   810
   End
   Begin VB.Label lblTelefono 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono 3:"
      Height          =   210
      Index           =   3
      Left            =   4800
      TabIndex        =   48
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label lblTelefono 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono 2:"
      Height          =   210
      Index           =   2
      Left            =   4800
      TabIndex        =   42
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label lblTelefono 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono 1:"
      Height          =   210
      Index           =   1
      Left            =   4800
      TabIndex        =   36
      Top             =   960
      Width           =   810
   End
   Begin VB.Label lblLocalidad 
      AutoSize        =   -1  'True
      Caption         =   "&Localidad:"
      Height          =   210
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblDomicilioCalle3 
      AutoSize        =   -1  'True
      Caption         =   "&Calle 3:"
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   525
   End
   Begin VB.Label lblDomicilioCalle2 
      AutoSize        =   -1  'True
      Caption         =   "&Calle 2:"
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   525
   End
   Begin VB.Line Line1 
      X1              =   4650
      X2              =   4650
      Y1              =   900
      Y2              =   6480
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      Caption         =   "&Tipo:"
      Height          =   210
      Left            =   4800
      TabIndex        =   66
      Top             =   2880
      Width           =   345
   End
   Begin VB.Label lblNotas 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
      Height          =   210
      Left            =   4800
      TabIndex        =   79
      Top             =   4620
      Width           =   465
   End
   Begin VB.Label lblProvincia 
      AutoSize        =   -1  'True
      Caption         =   "P&rovincia:"
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   705
   End
   Begin VB.Label lblCodigoPostal 
      AutoSize        =   -1  'True
      Caption         =   "C.P.:"
      Height          =   210
      Left            =   120
      TabIndex        =   20
      Top             =   3840
      Width           =   330
   End
   Begin VB.Label lblDomicilioDepartamento 
      AutoSize        =   -1  'True
      Caption         =   "&Dpto.:"
      Height          =   210
      Left            =   3360
      TabIndex        =   10
      Top             =   2040
      Width           =   420
   End
   Begin VB.Label lblDomicilioPiso 
      AutoSize        =   -1  'True
      Caption         =   "&Piso:"
      Height          =   210
      Left            =   2100
      TabIndex        =   8
      Top             =   2040
      Width           =   345
   End
   Begin VB.Label lblDomicilioNumero 
      AutoSize        =   -1  'True
      Caption         =   "N&úmero:"
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   600
   End
   Begin VB.Label lblDomicilioCalle1 
      AutoSize        =   -1  'True
      Caption         =   "&Calle:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   390
   End
   Begin VB.Label lblApellido 
      AutoSize        =   -1  'True
      Caption         =   "Ap&ellido:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   600
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos de la Persona"
      Height          =   210
      Left            =   780
      TabIndex        =   87
      Top             =   240
      Width           =   2640
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "PersonaPropiedad.frx":4608
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmPersonaPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPersona As Persona
Private mNew As Boolean
Private mKeyDecimal As Boolean

Private mFormWaitingForSelectSave As String
Private mSelectTypeFilterSave As String
Private mLoading As Boolean

Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mPersona
End Sub

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef Persona As Persona)
    Dim PersonaACargo As Persona
    Dim PersonaCuentaCorriente As Persona
    Dim Feriado As Feriado

    Set mPersona = Persona
    Set Persona = Nothing
    mNew = (mPersona.IDPersona = 0)
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    mLoading = True
    
    With mPersona
        txtApellido.Text = .Apellido
        txtNombre.Text = .Nombre
        txtDomicilioCalle1.Text = .DomicilioCalle1
        txtDomicilioNumero.Text = .DomicilioNumero
        txtDomicilioPiso.Text = .DomicilioPiso
        txtDomicilioDepartamento.Text = .DomicilioDepartamento
        txtDomicilioCalle2.Text = .DomicilioCalle2
        txtDomicilioCalle3.Text = .DomicilioCalle3
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboProvincia, "(SELECT ' ' AS IDProvincia, '----------' AS Nombre) UNION (SELECT IDProvincia, Nombre FROM Provincia) ORDER BY Nombre", "IDProvincia", "Nombre", "Provincias", cscpItemOrFirst, .IDProvincia) Then
            Unload Me
            Exit Sub
        End If
        datcboLocalidad.BoundText = .IDLocalidad
        txtCodigoPostal.Text = .CodigoPostal
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboDocumentoTipo, "(SELECT 0 AS IDDocumentoTipo, '------' AS Nombre, 1 AS Orden FROM DocumentoTipo) UNION (SELECT IDDocumentoTipo, Nombre, 2 AS Orden FROM DocumentoTipo WHERE Activo = 1) ORDER BY Orden, Nombre", "IDDocumentoTipo", "Nombre", "Tipos de Documento", cscpItemOrFirst, .IDDocumentoTipo) Then
            Unload Me
            Exit Sub
        End If
        txtDocumentoNumero.Text = .DocumentoNumero
        If Not CSM_Control_DataCombo.FillFromSQL(datcboCondicionIVA, "(SELECT 0 AS IDCondicionIVA, '------------------------' AS Nombre, 1 AS Orden FROM CondicionIVA) UNION (SELECT IDCondicionIVA, Nombre, 2 AS Orden FROM CondicionIVA WHERE Activo = 1) ORDER BY Orden, Nombre", "IDCondicionIVA", "Nombre", "Condiciones de IVA", cscpItemOrFirst, .IDCondicionIVA) Then
            Unload Me
            Exit Sub
        End If
        
        If .FechaNacimiento = DATE_TIME_FIELD_NULL_VALUE Then
            dtpFechaNacimiento.value = Date
            dtpFechaNacimiento.value = Null
        Else
            dtpFechaNacimiento.value = .FechaNacimiento
        End If
        txtEmail.Text = .Email
        
        txtPersonaACargo.Tag = .IDPersonaACargo
        If .IDPersonaACargo > 0 Then
            Set PersonaACargo = New Persona
            PersonaACargo.IDPersona = .IDPersonaACargo
            If PersonaACargo.Load() Then
                txtPersonaACargo.Text = PersonaACargo.ApellidoNombre
            End If
            Set PersonaACargo = Nothing
        Else
            txtPersonaACargo.Text = ""
        End If
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboTelefonoTipo(1), "(SELECT 0 AS IDTelefonoTipo, '--------' AS Nombre, 1 AS Orden FROM TelefonoTipo) UNION (SELECT IDTelefonoTipo, Nombre, 2 AS Orden FROM TelefonoTipo WHERE Activo = 1) ORDER BY Orden, IDTelefonoTipo", "IDTelefonoTipo", "Nombre", "Tipos de Teléfonos", cscpItemOrFirst, .IDTelefono1Tipo) Then
            Unload Me
            Exit Sub
        End If
        txtTelefonoTipoOtro(1).Text = .Telefono1TipoOtro
        txtTelefonoArea(1).Text = .Telefono1Area
        txtTelefonoNumero(1).Text = .Telefono1Numero
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboTelefonoTipo(2), "(SELECT 0 AS IDTelefonoTipo, '--------' AS Nombre, 1 AS Orden FROM TelefonoTipo) UNION (SELECT IDTelefonoTipo, Nombre, 2 AS Orden FROM TelefonoTipo WHERE Activo = 1) ORDER BY Orden, IDTelefonoTipo", "IDTelefonoTipo", "Nombre", "Tipos de Teléfonos", cscpItemOrFirst, .IDTelefono2Tipo) Then
            Unload Me
            Exit Sub
        End If
        txtTelefonoTipoOtro(2).Text = .Telefono2TipoOtro
        txtTelefonoArea(2).Text = .Telefono2Area
        txtTelefonoNumero(2).Text = .Telefono2Numero
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboTelefonoTipo(3), "(SELECT 0 AS IDTelefonoTipo, '--------' AS Nombre, 1 AS Orden FROM TelefonoTipo) UNION (SELECT IDTelefonoTipo, Nombre, 2 AS Orden FROM TelefonoTipo WHERE Activo = 1) ORDER BY Orden, IDTelefonoTipo", "IDTelefonoTipo", "Nombre", "Tipos de Teléfonos", cscpItemOrFirst, .IDTelefono3Tipo) Then
            Unload Me
            Exit Sub
        End If
        txtTelefonoTipoOtro(3).Text = .Telefono3TipoOtro
        txtTelefonoArea(3).Text = .Telefono3Area
        txtTelefonoNumero(3).Text = .Telefono3Numero
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboTelefonoTipo(4), "(SELECT 0 AS IDTelefonoTipo, '--------' AS Nombre, 1 AS Orden FROM TelefonoTipo) UNION (SELECT IDTelefonoTipo, Nombre, 2 AS Orden FROM TelefonoTipo WHERE Activo = 1) ORDER BY Orden, IDTelefonoTipo", "IDTelefonoTipo", "Nombre", "Tipos de Teléfonos", cscpItemOrFirst, .IDTelefono4Tipo) Then
            Unload Me
            Exit Sub
        End If
        txtTelefonoTipoOtro(4).Text = .Telefono4TipoOtro
        txtTelefonoArea(4).Text = .Telefono4Area
        txtTelefonoNumero(4).Text = .Telefono4Numero
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboTelefonoTipo(5), "(SELECT 0 AS IDTelefonoTipo, '--------' AS Nombre, 1 AS Orden FROM TelefonoTipo) UNION (SELECT IDTelefonoTipo, Nombre, 2 AS Orden FROM TelefonoTipo WHERE Activo = 1) ORDER BY Orden, IDTelefonoTipo", "IDTelefonoTipo", "Nombre", "Tipos de Teléfonos", cscpItemOrFirst, .IDTelefono5Tipo) Then
            Unload Me
            Exit Sub
        End If
        txtTelefonoTipoOtro(5).Text = .Telefono5TipoOtro
        txtTelefonoArea(5).Text = .Telefono5Area
        txtTelefonoNumero(5).Text = .Telefono5Numero
        
        optTipoCliente.value = (.EntidadTipo = ENTIDAD_TIPO_PERSONA_CLIENTE)
        optTipoConductor.value = (.EntidadTipo = ENTIDAD_TIPO_PERSONA_CONDUCTOR)
        optTipoAdministrativo.value = (.EntidadTipo = ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO)
        If .EntidadTipo = "" Then
            optTipoCliente.value = True
        End If
        
        txtSueldoImporte.Text = .SueldoImporte
        txtSueldoImporte_LostFocus
        If .SueldoDia = 99 Then
            cboSueldoDia.ListIndex = 29
        Else
            cboSueldoDia.ListIndex = .SueldoDia
        End If

        txtPersonaCuentaCorriente.Tag = .IDPersonaCuentaCorriente
        If .IDPersonaCuentaCorriente > 0 Then
            Set PersonaCuentaCorriente = New Persona
            PersonaCuentaCorriente.IDPersona = .IDPersonaCuentaCorriente
            If PersonaCuentaCorriente.Load() Then
                txtPersonaCuentaCorriente.Text = PersonaCuentaCorriente.ApellidoNombre
            End If
            Set PersonaCuentaCorriente = Nothing
        Else
            txtPersonaCuentaCorriente.Text = ""
        End If
        
        chkPermiteViajarSinPagar.value = IIf(.PermiteViajarSinPagar, vbChecked, vbUnchecked)
        chkHabilitadoViajar.value = IIf(.HabilitadoViajar, vbChecked, vbUnchecked)
        chkHabilitadoInternet.value = IIf(.HabilitadoInternet, vbChecked, vbUnchecked)
        txtNotas.Text = .Notas
        chkActivo.value = IIf(.Activo, vbChecked, vbUnchecked)
        chkListaPasajero.value = IIf(.ListaPasajero, vbChecked, vbUnchecked)
        
        If Not mNew Then
            SetLastPersona .IDPersona, .ApellidoNombre
        End If
    
        Set Feriado = New Feriado
        Feriado.VerificarReservasDelPasajero .IDPersona
        Set Feriado = Nothing
    End With
    
    optTipoCliente.Enabled = (mNew And pCPermiso.GotPermission(PERMISO_PERSONA_ADD_ALLTYPE, False))
    optTipoConductor.Enabled = optTipoCliente.Enabled
    optTipoAdministrativo.Enabled = optTipoCliente.Enabled
    
    mLoading = False
        
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEmail_Click()
    CSM_Instance.Execute "mailto:" & txtEmail.Text, , , , , Me.hwnd
End Sub

Private Sub cmdPersonaACargo_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        frmPersona.FormKeepOpenOnSelect = True
        Call frmPersona.FindAndShowItem(Val(txtPersonaACargo.Tag), UCase(Left(txtPersonaACargo.Text, 1)), Me.Name, "", "PC")
    End If
End Sub

Private Sub cmdPersonaACargoClear_Click()
    txtPersonaACargo.Tag = 0
    txtPersonaACargo.Text = ""
    
    On Error Resume Next
    datcboTelefonoTipo(0).SetFocus
End Sub

Private Sub cmdInternetDataSend_Click()
    Dim EmailMessage As EmailMessage
    
    Const LINE_FEED As String = vbCrLf
    
    If chkHabilitadoInternet.value = vbUnchecked Then
        MsgBox "Debe habilitar la Persona para que opere por Internet.", vbInformation, App.Title
        chkHabilitadoInternet.SetFocus
        Exit Sub
    End If
    If (datcboDocumentoTipo.BoundText = "------" Or Trim(txtDocumentoNumero.Text) = "") Then
        MsgBox "Debe completar el Número de Documento para que la Persona opere por Internet.", vbInformation, App.Title
        datcboDocumentoTipo.SetFocus
        Exit Sub
    End If
    If Trim(txtEmail.Text) = "" Then
        MsgBox "Debe completar el E-mail para que la Persona opere por Internet.", vbInformation, App.Title
        txtEmail.SetFocus
        Exit Sub
    End If
    If InStr(1, txtEmail.Text, "@") = 0 Then
        MsgBox "La Dirección de E-mail de la Persona es incorrecta porque no contiene la arroba '@'.", vbInformation, App.Title
        txtEmail.SetFocus
        Exit Sub
    End If
    
    Set EmailMessage = New EmailMessage
    With EmailMessage
        .DateTime = Now
        .SenderDisplayName = pParametro.CompanyName
        .SenderAddress = pParametro.InternetDataEmail_SenderAddress
        .RecipientToDisplayName = Trim(txtApellido.Text) & ", " & Trim(txtNombre.Text)
        .RecipientToAddress = txtEmail.Text
        .Subject = CSTransporte_SDK.TextReplaceSystemVariables(pParametro.WebSite_Persona_EmailBienvenida_Asunto, mPersona)
        .Body = CSTransporte_SDK.TextReplaceSystemVariables(pParametro.WebSite_Persona_EmailBienvenida_Texto, mPersona)
        
        Call .Add
    End With
    Set EmailMessage = Nothing
End Sub

Private Sub cmdOK_Click()
    If Trim(txtApellido.Text) = "" Then
        MsgBox "Debe ingresar el Apellido de la Persona.", vbInformation, App.Title
        txtApellido.SetFocus
        Exit Sub
    End If
    
    'DOMICILIO
    If Trim(txtDomicilioCalle1.Text) = "" And pParametro.Persona_DatoIncompleto_CampoDomicilio And (mNew Or pParametro.Persona_DatoIncompleto_RegistroTodos) Then
        Select Case pParametro.Persona_DatoIncompleto_AvisoTipo
            Case 0
            Case 1
                If MsgBox("Esta persona no tiene ingresado el Domicilio." & vbCr & vbCr & "¿Desea ingresarlo ahora?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    txtDomicilioCalle1.SetFocus
                    Exit Sub
                End If
            Case 2
                MsgBox "Es obligatorio ingresar el Domicilio", vbExclamation, App.Title
                txtDomicilioCalle1.SetFocus
                Exit Sub
        End Select
    End If
    
    'DOCUMENTO
    If Val(datcboDocumentoTipo.BoundText) > 0 And Trim(txtDocumentoNumero.Text) = "" Then
        MsgBox "Si selecciona el Tipo de Documento, deberá completar el Número de Documento.", vbInformation, App.Title
        txtDocumentoNumero.SetFocus
        Exit Sub
    End If
    If Val(datcboDocumentoTipo.BoundText) = 0 And Trim(txtDocumentoNumero.Text) <> "" Then
        MsgBox "Si ingresa el Número de Documento, deberá especificar el Tipo de Documento.", vbInformation, App.Title
        datcboDocumentoTipo.SetFocus
        Exit Sub
    End If
    If Trim(txtDocumentoNumero.Text) = "" And pParametro.Persona_DatoIncompleto_CampoDocumento And (mNew Or pParametro.Persona_DatoIncompleto_RegistroTodos) Then
        Select Case pParametro.Persona_DatoIncompleto_AvisoTipo
            Case 0
            Case 1
                If MsgBox("Esta persona no tiene ingresado el Tipo y Número de Documento." & vbCr & vbCr & "¿Desea ingresarlos ahora?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    datcboDocumentoTipo.SetFocus
                    Exit Sub
                End If
            Case 2
                MsgBox "Es obligatorio ingresar el Tipo y Número de Documento", vbExclamation, App.Title
                datcboDocumentoTipo.SetFocus
                Exit Sub
        End Select
    End If
    
    If Trim(txtEmail.Text) <> "" Then
        If InStr(1, txtEmail.Text, "@") = 0 Then
            MsgBox "La Dirección de E-mail de la Persona es incorrecta porque no contiene la arroba '@'.", vbInformation, App.Title
            txtEmail.SetFocus
            Exit Sub
        End If
    End If
    If chkHabilitadoInternet.value = vbChecked Then
        If (datcboDocumentoTipo.BoundText = "------" Or Trim(txtDocumentoNumero.Text) = "") Then
            If MsgBox("Esta Persona no podrá operar en Internet hasta que no complete su Tipo y Número de Documento." & vbCr & vbCr & "¿Desea completar los Datos?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                datcboDocumentoTipo.SetFocus
                Exit Sub
            End If
        End If
        If Trim(txtEmail.Text) = "" Then
            If MsgBox("Esta Persona no podrá operar en Internet hasta que no complete su dirección de E-mail." & vbCr & vbCr & "¿Desea completar los Datos?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                txtEmail.SetFocus
                Exit Sub
            End If
        End If
    End If
    If optTipoAdministrativo.value Then
        If Not IsNumeric(txtSueldoImporte.Text) Then
            MsgBox "El Sueldo ingresado es incorrecto.", vbInformation, App.Title
            txtSueldoImporte.SetFocus
            Exit Sub
        End If
        If CCur(txtSueldoImporte.Text) < 0 Then
            MsgBox "El Sueldo debe ser mayor o igual a cero.", vbInformation, App.Title
            txtSueldoImporte.SetFocus
            Exit Sub
        End If
    End If
    
    With mPersona
        .Apellido = txtApellido.Text
        .Nombre = txtNombre.Text
        .DomicilioCalle1 = txtDomicilioCalle1.Text
        .DomicilioNumero = txtDomicilioNumero.Text
        .DomicilioPiso = txtDomicilioPiso.Text
        .DomicilioDepartamento = txtDomicilioDepartamento.Text
        .DomicilioCalle2 = txtDomicilioCalle2.Text
        .DomicilioCalle3 = txtDomicilioCalle3.Text
        .CodigoPostal = txtCodigoPostal.Text
        .IDLocalidad = Val(datcboLocalidad.BoundText)
        .IDProvincia = datcboProvincia.BoundText
        .IDDocumentoTipo = IIf(datcboDocumentoTipo.BoundText = "------", "", datcboDocumentoTipo.BoundText)
        .DocumentoNumero = txtDocumentoNumero.Text
        .IDCondicionIVA = IIf(datcboCondicionIVA.BoundText = "------", "", datcboCondicionIVA.BoundText)
        
        .FechaNacimiento = IIf(IsNull(dtpFechaNacimiento.value), DATE_TIME_FIELD_NULL_VALUE, dtpFechaNacimiento.value)
        .Email = txtEmail.Text
        
        .IDPersonaACargo = Val(txtPersonaACargo.Tag)
    
        .IDTelefono1Tipo = Val(datcboTelefonoTipo(1).BoundText)
        .Telefono1TipoOtro = txtTelefonoTipoOtro(1).Text
        .Telefono1Area = txtTelefonoArea(1).Text
        .Telefono1Numero = txtTelefonoNumero(1).Text
    
        .IDTelefono2Tipo = Val(datcboTelefonoTipo(2).BoundText)
        .Telefono2TipoOtro = txtTelefonoTipoOtro(2).Text
        .Telefono2Area = txtTelefonoArea(2).Text
        .Telefono2Numero = txtTelefonoNumero(2).Text
    
        .IDTelefono3Tipo = Val(datcboTelefonoTipo(3).BoundText)
        .Telefono3TipoOtro = txtTelefonoTipoOtro(3).Text
        .Telefono3Area = txtTelefonoArea(3).Text
        .Telefono3Numero = txtTelefonoNumero(3).Text
    
        .IDTelefono4Tipo = Val(datcboTelefonoTipo(4).BoundText)
        .Telefono4TipoOtro = txtTelefonoTipoOtro(4).Text
        .Telefono4Area = txtTelefonoArea(4).Text
        .Telefono4Numero = txtTelefonoNumero(4).Text
    
        .IDTelefono5Tipo = Val(datcboTelefonoTipo(5).BoundText)
        .Telefono5TipoOtro = txtTelefonoTipoOtro(5).Text
        .Telefono5Area = txtTelefonoArea(5).Text
        .Telefono5Numero = txtTelefonoNumero(5).Text
        
        If optTipoAdministrativo.value Then
            .SueldoImporte = CCur(txtSueldoImporte.Text)
            If cboSueldoDia.ListIndex = 29 Then
                .SueldoDia = 99
            Else
                .SueldoDia = cboSueldoDia.ListIndex
            End If
        End If
            
        .EntidadTipo = IIf(optTipoCliente.value, ENTIDAD_TIPO_PERSONA_CLIENTE, IIf(optTipoConductor.value, ENTIDAD_TIPO_PERSONA_CONDUCTOR, ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO))
        .IDPersonaCuentaCorriente = Val(txtPersonaCuentaCorriente.Tag)
        .PermiteViajarSinPagar = (chkPermiteViajarSinPagar.value = vbChecked)
        .HabilitadoViajar = (chkHabilitadoViajar.value = vbChecked)
        .HabilitadoInternet = (chkHabilitadoInternet.value = vbChecked)
        .Notas = txtNotas.Text
        .Activo = (chkActivo.value = vbChecked)
        .ListaPasajero = (chkListaPasajero.value = vbChecked)
        
        If mNew Then
            If Not .AddNew() Then
                Exit Sub
            End If
        Else
            If Not .Update() Then
                Exit Sub
            End If
        End If
        
        SetLastPersona .IDPersona, .ApellidoNombre
    End With

    Unload Me
End Sub

Private Sub cmdPersonaCuentaCorriente_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        frmPersona.FormKeepOpenOnSelect = True
        Call frmPersona.FindAndShowItem(Val(txtPersonaCuentaCorriente.Tag), UCase(Left(txtPersonaCuentaCorriente.Text, 1)), Me.Name, "", "PF")
    End If
End Sub

Private Sub cmdPersonaCuentaCorrienteClear_Click()
    txtPersonaCuentaCorriente.Tag = 0
    txtPersonaCuentaCorriente.Text = ""
    
    On Error Resume Next
    txtNotas.SetFocus
End Sub

Private Sub datcboLocalidad_Change()
    Dim Localidad As Localidad
    
    If Val(datcboLocalidad.BoundText) = 0 Then
        txtCodigoPostal.Text = ""
    Else
        Set Localidad = New Localidad
        
        Localidad.IDProvincia = datcboProvincia.BoundText
        Localidad.IDLocalidad = Val(datcboLocalidad.BoundText)
        If Localidad.Load() Then
            txtCodigoPostal.Text = IIf(Localidad.CodigoPostal = 0, "", Localidad.CodigoPostal)
        End If
        
        Set Localidad = Nothing
    End If
End Sub

Private Sub datcboProvincia_Change()
    datcboLocalidad.BoundText = ""
    Call CSM_Control_DataCombo.FillFromSQL(datcboLocalidad, "SELECT IDLocalidad, Nombre FROM Localidad WHERE IDProvincia = '" & datcboProvincia.BoundText & "' ORDER BY Nombre", "IDLocalidad", "Nombre", "Localidades", cscpFirstIfUnique)
End Sub

Private Sub Form_Load()
    Dim Dia As Byte
    
    cboSueldoDia.AddItem "---"
    For Dia = 1 To 28
        cboSueldoDia.AddItem Dia
    Next Dia
    cboSueldoDia.AddItem "Ultimo"
    
    chkHabilitadoViajar.Visible = pCPermiso.GotPermission(PERMISO_PERSONA_HABILITACION_VIAJAR_ESTABLECER, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mPersona = Nothing
    Set frmPersonaPropiedad = Nothing
End Sub

Private Sub optTipoAdministrativo_Click()
    ShowControls
End Sub

Private Sub optTipoCliente_Click()
    ShowControls
End Sub

Private Sub optTipoConductor_Click()
    ShowControls
End Sub

Private Sub txtApellido_Change()
    Caption = "Propiedades" & IIf(Trim(txtApellido.Text) = "", IIf(Trim(txtNombre.Text) = "", "", " de " & txtNombre.Text), " de " & txtApellido.Text & IIf(Trim(txtNombre.Text) = "", "", ", " & txtNombre.Text))
End Sub

Private Sub txtApellido_GotFocus()
    CSM_Control_TextBox.SelAllText txtApellido
End Sub

Private Sub txtApellido_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtApellido_LostFocus()
    txtApellido.Text = UCase(txtApellido.Text)
    txtApellido.Text = CleanInvalidSpaces(txtApellido.Text)
End Sub

Private Sub txtCodigoPostal_GotFocus()
    CSM_Control_TextBox.SelAllText txtCodigoPostal
End Sub

Private Sub txtCodigoPostal_LostFocus()
    txtCodigoPostal.Text = CleanInvalidSpaces(txtCodigoPostal.Text)
End Sub

Private Sub txtDocumentoNumero_GotFocus()
    CSM_Control_TextBox.SelAllText txtDocumentoNumero
End Sub

Private Sub txtDocumentoNumero_LostFocus()
    txtDocumentoNumero.Text = CleanInvalidSpaces(txtDocumentoNumero.Text)
End Sub

Private Sub txtDomicilioCalle1_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioCalle1
End Sub

Private Sub txtDomicilioCalle1_LostFocus()
    txtDomicilioCalle1.Text = CleanInvalidSpaces(txtDomicilioCalle1.Text)
End Sub

Private Sub txtDomicilioCalle2_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioCalle2
End Sub

Private Sub txtDomicilioCalle2_LostFocus()
    txtDomicilioCalle2.Text = CleanInvalidSpaces(txtDomicilioCalle2.Text)
End Sub

Private Sub txtDomicilioCalle3_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioCalle3
End Sub

Private Sub txtDomicilioCalle3_LostFocus()
    txtDomicilioCalle3.Text = CleanInvalidSpaces(txtDomicilioCalle3.Text)
End Sub

Private Sub txtDomicilioDepartamento_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioDepartamento
End Sub

Private Sub txtDomicilioDepartamento_LostFocus()
    txtDomicilioDepartamento.Text = CleanInvalidSpaces(txtDomicilioDepartamento.Text)
End Sub

Private Sub txtDomicilioNumero_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioNumero
End Sub

Private Sub txtDomicilioNumero_LostFocus()
    txtDomicilioNumero.Text = CleanInvalidSpaces(txtDomicilioNumero.Text)
End Sub

Private Sub txtDomicilioPiso_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioPiso
End Sub

Private Sub txtDomicilioPiso_LostFocus()
    txtDomicilioPiso.Text = CleanInvalidSpaces(txtDomicilioPiso.Text)
End Sub

Private Sub txtEmail_Change()
    cmdEmail.Visible = (Trim(txtEmail.Text) <> "")
End Sub

Private Sub txtEmail_GotFocus()
    CSM_Control_TextBox.SelAllText txtEmail
End Sub

Private Sub txtEmail_LostFocus()
    txtEmail.Text = CleanInvalidSpaces(txtEmail.Text)
End Sub

Private Sub txtNombre_Change()
    Caption = "Propiedades" & IIf(Trim(txtApellido.Text) = "", IIf(Trim(txtNombre.Text) = "", "", " de " & txtNombre.Text), " de " & txtApellido.Text & IIf(Trim(txtNombre.Text) = "", "", ", " & txtNombre.Text))
End Sub

Private Sub txtNombre_GotFocus()
    CSM_Control_TextBox.SelAllText txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombre_LostFocus()
    txtNombre.Text = UCase(txtNombre.Text)
    txtNombre.Text = CleanInvalidSpaces(txtNombre.Text)
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
    txtNotas.Text = CleanInvalidSpaces(txtNotas.Text)
End Sub

Private Sub datcboTelefonoTipo_Change(Index As Integer)
    txtTelefonoTipoOtro(Index).Visible = (Val(datcboTelefonoTipo(Index).BoundText) = pParametro.TelefonoTipo_ID_Otro)
    txtTelefonoArea(Index).Visible = (Val(datcboTelefonoTipo(Index).BoundText) > 0)
    txtTelefonoNumero(Index).Visible = (Val(datcboTelefonoTipo(Index).BoundText) > 0)
End Sub

Private Sub txtTelefonoTipoOtro_GotFocus(Index As Integer)
    CSM_Control_TextBox.SelAllText txtTelefonoTipoOtro(Index)
End Sub

Private Sub txtTelefonoArea_GotFocus(Index As Integer)
    CSM_Control_TextBox.SelAllText txtTelefonoArea(Index)
End Sub

Private Sub txtTelefonoArea_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTelefonoArea_LostFocus(Index As Integer)
    txtTelefonoArea(Index).Text = CleanNotNumericChars(txtTelefonoArea(Index).Text)
End Sub

Private Sub txtTelefonoNumero_GotFocus(Index As Integer)
    CSM_Control_TextBox.SelAllText txtTelefonoNumero(Index)
End Sub

Private Sub txtTelefonoNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTelefonoNumero_Change(Index As Integer)
    If pTelephony.TelephonyType <> "NONE" Then
        cmdTelefonoDial(Index).Visible = (Trim(txtTelefonoNumero(Index).Text) <> "" And pTelephony.Initialized)
    Else
        cmdTelefonoDial(Index).Visible = False
    End If
    If pTelephony.Initialized And Not mLoading Then
        If Trim(txtTelefonoNumero(Index).Text) = "" Then
            If Trim(txtTelefonoArea(Index).Text) = pTelephony.LocationCityCode Then
                txtTelefonoArea(Index).Text = ""
            End If
        Else
            If Trim(txtTelefonoArea(Index).Text) = "" Then
                txtTelefonoArea(Index).Text = pTelephony.LocationCityCode
            End If
        End If
    End If
End Sub

Private Sub txtTelefonoNumero_LostFocus(Index As Integer)
    txtTelefonoNumero(Index).Text = CleanNotNumericChars(txtTelefonoNumero(Index).Text)
End Sub

Private Sub cmdTelefonoDial_Click(Index As Integer)
    Dim TelefonoTipo As TelefonoTipo
    
    If pTelephony.TelephonyType <> "NONE" And pTelephony.Initialized Then
        If Val(datcboTelefonoTipo(Index).BoundText) > 0 Then
            Set TelefonoTipo = New TelefonoTipo
            TelefonoTipo.IDTelefonoTipo = Val(datcboTelefonoTipo(Index).BoundText)
            If TelefonoTipo.Load() Then
                Call pTelephony.DialNumber(txtTelefonoArea(Index).Text, TelefonoTipo.DiscadoPrefijo & txtTelefonoNumero(Index).Text & TelefonoTipo.DiscadoSufijo)
            End If
            Set TelefonoTipo = Nothing
        Else
            Call pTelephony.DialNumber(txtTelefonoArea(Index).Text, txtTelefonoNumero(Index).Text)
        End If
    End If
End Sub

Private Sub txtSueldoImporte_GotFocus()
    CSM_Control_TextBox.SelAllText txtSueldoImporte
End Sub

Private Sub txtSueldoImporte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtSueldoImporte_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, txtSueldoImporte.Text, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, txtSueldoImporte.Text, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtSueldoImporte_LostFocus()
    If Not IsNumeric(txtSueldoImporte.Text) Then
        txtSueldoImporte.Text = Val(txtSueldoImporte.Text)
    End If
    txtSueldoImporte.Text = Format(CCur(txtSueldoImporte.Text), "Currency")
End Sub

Public Sub PersonaSelected(ByVal IDPersona As Long, ByVal Tag As String)
    Dim Persona As Persona
    
    frmPersona.FormWaitingForSelect = mFormWaitingForSelectSave
    frmPersona.SelectTypeFilter = mSelectTypeFilterSave
    
    Select Case Tag
        Case "PF"
            If IDPersona = mPersona.IDPersona Then
                MsgBox "No se puede seleccionar la misma Persona para Facturar.", vbExclamation, App.Title
            Else
                Set Persona = New Persona
                Persona.IDPersona = IDPersona
                If Persona.Load() Then
                    If Persona.IDPersonaCuentaCorriente > 0 Then
                        MsgBox "No se puede seleccionar a una Persona que ya tiene especificado debitar a otra Persona.", vbExclamation, App.Title
                    Else
                        txtPersonaCuentaCorriente.Tag = IDPersona
                        txtPersonaCuentaCorriente.Text = Persona.ApellidoNombre
                    End If
                End If
                Set Persona = Nothing
            End If
        Case "PC"
            If IDPersona = mPersona.IDPersona Then
                MsgBox "No se puede seleccionar la misma Persona para estar A Cargo.", vbExclamation, App.Title
            Else
                Set Persona = New Persona
                Persona.IDPersona = IDPersona
                If Persona.Load() Then
                    If Persona.IDPersonaACargo > 0 Then
                        MsgBox "No se puede seleccionar a una Persona que tiene especificada una Persona A Cargo.", vbExclamation, App.Title
                    Else
                        txtPersonaACargo.Tag = IDPersona
                        txtPersonaACargo.Text = Persona.ApellidoNombre
                    End If
                End If
                Set Persona = Nothing
            End If
    End Select
    
    On Error Resume Next
    txtNotas.SetFocus
End Sub

Private Sub ShowControls()
    lblSueldoImporte.Visible = (optTipoAdministrativo.value Or optTipoConductor.value)
    txtSueldoImporte.Visible = (optTipoAdministrativo.value Or optTipoConductor.value)
    lblSueldoDia.Visible = (optTipoAdministrativo.value Or optTipoConductor.value)
    cboSueldoDia.Visible = (optTipoAdministrativo.value Or optTipoConductor.value)
End Sub

Public Sub FillComboBoxDocumentoTipo()
    Dim KeySave As String
    Dim recData As ADODB.Recordset
    
    KeySave = datcboDocumentoTipo.BoundText
    Set recData = datcboDocumentoTipo.RowSource
    recData.Requery
    Set recData = Nothing
    datcboDocumentoTipo.BoundText = KeySave
    If datcboDocumentoTipo.BoundText = "" Then
        datcboDocumentoTipo.BoundText = "------"
    End If
End Sub

Public Sub FillComboBoxTelefonoTipo()
    Dim KeySave(1 To 5) As Long
    Dim recData As ADODB.Recordset
    Dim Index As Integer
    
    For Index = 1 To 5
        KeySave(Index) = Val(datcboTelefonoTipo(Index).BoundText)
        Set recData = datcboTelefonoTipo(Index).RowSource
        recData.Requery
        Set recData = Nothing
        datcboTelefonoTipo(Index).BoundText = KeySave(Index)
        If Val(datcboTelefonoTipo(Index).BoundText) = 0 Then
            datcboTelefonoTipo(Index).BoundText = 0
        End If
    Next Index
End Sub
