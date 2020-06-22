VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmContactoPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ContactoPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   5895
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   4800
      Picture         =   "ContactoPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   140
      Top             =   60
      Width           =   975
   End
   Begin VB.Frame fraGeneral 
      Height          =   4635
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   5415
      Begin VB.CommandButton cmdContactoGrupo 
         Caption         =   "..."
         Height          =   315
         Left            =   5040
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Grupos"
         Top             =   4200
         Width           =   255
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   6
         Top             =   960
         Width           =   4035
      End
      Begin VB.TextBox txtApellido 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   4
         Top             =   600
         Width           =   4035
      End
      Begin VB.TextBox txtCompania 
         Height          =   315
         Left            =   1260
         MaxLength       =   100
         TabIndex        =   8
         Top             =   1320
         Width           =   4035
      End
      Begin VB.TextBox txtTituloLaboral 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1680
         Width           =   4035
      End
      Begin VB.TextBox txtTelefonoArea 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   38
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox txtTelefonoNumero 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   3840
         MaxLength       =   16
         TabIndex        =   39
         Top             =   3660
         Width           =   1095
      End
      Begin VB.TextBox txtTelefonoArea 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   32
         Top             =   3300
         Width           =   615
      End
      Begin VB.TextBox txtTelefonoNumero 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   3840
         MaxLength       =   16
         TabIndex        =   33
         Top             =   3300
         Width           =   1095
      End
      Begin VB.TextBox txtTelefonoArea 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   26
         Top             =   2940
         Width           =   615
      End
      Begin VB.TextBox txtTelefonoNumero 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   3840
         MaxLength       =   16
         TabIndex        =   27
         Top             =   2940
         Width           =   1095
      End
      Begin VB.TextBox txtTelefonoArea 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   20
         Top             =   2580
         Width           =   615
      End
      Begin VB.TextBox txtTelefonoNumero 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   3840
         MaxLength       =   16
         TabIndex        =   21
         Top             =   2580
         Width           =   1095
      End
      Begin VB.TextBox txtTelefonoArea 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   14
         Top             =   2220
         Width           =   615
      End
      Begin VB.TextBox txtTelefonoNumero 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   3840
         MaxLength       =   16
         TabIndex        =   15
         Top             =   2220
         Width           =   1095
      End
      Begin VB.CommandButton cmdTelefonoDial 
         Height          =   315
         Index           =   5
         Left            =   4980
         Picture         =   "ContactoPropiedad.frx":0B74
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Llamar"
         Top             =   3660
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdTelefonoDial 
         Height          =   315
         Index           =   4
         Left            =   4980
         Picture         =   "ContactoPropiedad.frx":1186
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Llamar"
         Top             =   3300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdTelefonoDial 
         Height          =   315
         Index           =   3
         Left            =   4980
         Picture         =   "ContactoPropiedad.frx":1798
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Llamar"
         Top             =   2940
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdTelefonoDial 
         Height          =   315
         Index           =   2
         Left            =   4980
         Picture         =   "ContactoPropiedad.frx":1DAA
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Llamar"
         Top             =   2580
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdTelefonoDial 
         Height          =   315
         Index           =   1
         Left            =   4980
         Picture         =   "ContactoPropiedad.frx":23BC
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Llamar"
         Top             =   2220
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtTelefonoTipoOtro 
         Height          =   315
         Index           =   5
         Left            =   2460
         MaxLength       =   15
         TabIndex        =   37
         Top             =   3660
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtTelefonoTipoOtro 
         Height          =   315
         Index           =   4
         Left            =   2460
         MaxLength       =   15
         TabIndex        =   31
         Top             =   3300
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtTelefonoTipoOtro 
         Height          =   315
         Index           =   3
         Left            =   2460
         MaxLength       =   15
         TabIndex        =   25
         Top             =   2940
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtTelefonoTipoOtro 
         Height          =   315
         Index           =   2
         Left            =   2460
         MaxLength       =   15
         TabIndex        =   19
         Top             =   2580
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtTelefonoTipoOtro 
         Height          =   315
         Index           =   1
         Left            =   2460
         MaxLength       =   15
         TabIndex        =   13
         Top             =   2220
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDataListLib.DataCombo datcboTelefonoTipo 
         Height          =   330
         Index           =   1
         Left            =   1260
         TabIndex        =   12
         Top             =   2220
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
         Left            =   1260
         TabIndex        =   18
         Top             =   2580
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
         Left            =   1260
         TabIndex        =   24
         Top             =   2940
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
         Left            =   1260
         TabIndex        =   30
         Top             =   3300
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
         Left            =   1260
         TabIndex        =   36
         Top             =   3660
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
      Begin MSDataListLib.DataCombo datcboContactoGrupo 
         Height          =   330
         Left            =   1260
         TabIndex        =   42
         Top             =   4200
         Width           =   3735
         _ExtentX        =   6588
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
      Begin MSDataListLib.DataCombo datcboTitulo 
         Height          =   330
         Left            =   1260
         TabIndex        =   2
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VB.Line Line3 
         X1              =   120
         X2              =   5280
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   5280
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label lblContactoGrupo 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   210
         Left            =   120
         TabIndex        =   41
         Top             =   4260
         Width           =   495
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         Caption         =   "&Nombre:"
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label lblApellido 
         AutoSize        =   -1  'True
         Caption         =   "Ap&ellido:"
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   615
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Título:"
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   420
      End
      Begin VB.Label lblCompania 
         AutoSize        =   -1  'True
         Caption         =   "Compañía:"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   1380
         Width           =   750
      End
      Begin VB.Label lblTituloLaboral 
         AutoSize        =   -1  'True
         Caption         =   "Título Laboral:"
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   1740
         Width           =   1005
      End
      Begin VB.Label lblTelefono 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono 5:"
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   35
         Top             =   3720
         Width           =   810
      End
      Begin VB.Label lblTelefono 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono 4:"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   3360
         Width           =   810
      End
      Begin VB.Label lblTelefono 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono 3:"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   810
      End
      Begin VB.Label lblTelefono 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono 2:"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   810
      End
      Begin VB.Label lblTelefono 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono 1:"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   810
      End
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   139
      Top             =   780
      Width           =   5715
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4500
      TabIndex        =   136
      Top             =   6600
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3180
      TabIndex        =   135
      Top             =   6600
      Width           =   1275
   End
   Begin VB.Frame fraDatosAdicionales 
      Height          =   3195
      Left            =   240
      TabIndex        =   125
      Top             =   1680
      Width           =   5415
      Begin VB.TextBox txtSobreNombre 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   127
         Top             =   240
         Width           =   4035
      End
      Begin VB.TextBox txtAsistente 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   131
         Top             =   1080
         Width           =   4035
      End
      Begin VB.TextBox txtNotas 
         Height          =   1065
         Left            =   1260
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   133
         Top             =   1620
         Width           =   4035
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "&Activo"
         Height          =   210
         Left            =   120
         TabIndex        =   134
         Top             =   2820
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFechaNacimiento 
         Height          =   315
         Left            =   1260
         TabIndex        =   129
         Top             =   660
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
         Format          =   108593153
         CurrentDate     =   36950
      End
      Begin VB.Label lblSobreNombre 
         AutoSize        =   -1  'True
         Caption         =   "Sobrenombre:"
         Height          =   210
         Left            =   120
         TabIndex        =   126
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label lblFechaNacimiento 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Nacim.:"
         Height          =   210
         Left            =   120
         TabIndex        =   128
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label lblAsistente 
         AutoSize        =   -1  'True
         Caption         =   "Secretaria/o:"
         Height          =   210
         Left            =   120
         TabIndex        =   130
         Top             =   1140
         Width           =   930
      End
      Begin VB.Label lblNotas 
         AutoSize        =   -1  'True
         Caption         =   "Notas:"
         Height          =   210
         Left            =   120
         TabIndex        =   132
         Top             =   1680
         Width           =   465
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5280
         Y1              =   1500
         Y2              =   1500
      End
   End
   Begin VB.Frame fraInternet 
      Height          =   3015
      Left            =   240
      TabIndex        =   106
      Top             =   1680
      Width           =   5415
      Begin VB.CommandButton cmdPaginaWeb 
         Height          =   315
         Left            =   4980
         Picture         =   "ContactoPropiedad.frx":29CE
         Style           =   1  'Graphical
         TabIndex        =   124
         TabStop         =   0   'False
         ToolTipText     =   "Abrir Página"
         Top             =   2580
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdEmail 
         Height          =   315
         Index           =   3
         Left            =   4980
         Picture         =   "ContactoPropiedad.frx":2FE0
         Style           =   1  'Graphical
         TabIndex        =   119
         TabStop         =   0   'False
         ToolTipText     =   "Enviar Mail"
         Top             =   1800
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdEmail 
         Height          =   315
         Index           =   2
         Left            =   4980
         Picture         =   "ContactoPropiedad.frx":35F2
         Style           =   1  'Graphical
         TabIndex        =   114
         TabStop         =   0   'False
         ToolTipText     =   "Enviar Mail"
         Top             =   1020
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdEmail 
         Height          =   315
         Index           =   1
         Left            =   4980
         Picture         =   "ContactoPropiedad.frx":3C04
         Style           =   1  'Graphical
         TabIndex        =   109
         TabStop         =   0   'False
         ToolTipText     =   "Enviar Mail"
         Top             =   240
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtPaginaWeb 
         Height          =   315
         Left            =   1260
         MaxLength       =   100
         TabIndex        =   123
         Top             =   2580
         Width           =   3675
      End
      Begin VB.TextBox txtEmailNombre 
         Height          =   315
         Index           =   3
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   121
         Top             =   2160
         Width           =   4035
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Index           =   3
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   118
         Top             =   1800
         Width           =   3675
      End
      Begin VB.TextBox txtEmailNombre 
         Height          =   315
         Index           =   2
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   116
         Top             =   1380
         Width           =   4035
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Index           =   2
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   113
         Top             =   1020
         Width           =   3675
      End
      Begin VB.TextBox txtEmailNombre 
         Height          =   315
         Index           =   1
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   111
         Top             =   600
         Width           =   4035
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Index           =   1
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   108
         Top             =   240
         Width           =   3675
      End
      Begin VB.Label lblPaginaWeb 
         AutoSize        =   -1  'True
         Caption         =   "Página Web:"
         Height          =   210
         Left            =   120
         TabIndex        =   122
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label lblEmailNombre 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   120
         Top             =   2220
         Width           =   900
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "E-mail 3:"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   117
         Top             =   1860
         Width           =   600
      End
      Begin VB.Label lblEmailNombre 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   115
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "E-mail 2:"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   112
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblEmailNombre 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   110
         Top             =   660
         Width           =   900
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "E-mail 1:"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   107
         Top             =   300
         Width           =   600
      End
   End
   Begin VB.Frame fraDomicilioOtro 
      Height          =   3495
      Left            =   240
      TabIndex        =   84
      Top             =   1680
      Width           =   5415
      Begin VB.CheckBox chkDomicilioOtroMailing 
         Caption         =   "Enviar correspondencia a este Domicilio"
         Height          =   210
         Left            =   1260
         TabIndex        =   105
         Top             =   3180
         Width           =   3795
      End
      Begin VB.TextBox txtDomicilioOtroCalle1 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   88
         Top             =   600
         Width           =   4035
      End
      Begin VB.TextBox txtDomicilioOtroNumero 
         Height          =   315
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   90
         Top             =   960
         Width           =   1035
      End
      Begin VB.TextBox txtDomicilioOtroPiso 
         Height          =   315
         Left            =   2940
         MaxLength       =   10
         TabIndex        =   92
         Top             =   960
         Width           =   795
      End
      Begin VB.TextBox txtDomicilioOtroDepartamento 
         Height          =   315
         Left            =   4500
         MaxLength       =   10
         TabIndex        =   94
         Top             =   960
         Width           =   795
      End
      Begin VB.TextBox txtDomicilioOtroCodigoPostal 
         Height          =   315
         Left            =   1260
         MaxLength       =   8
         TabIndex        =   104
         Top             =   2760
         Width           =   1035
      End
      Begin VB.TextBox txtDomicilioOtroCalle2 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   96
         Top             =   1320
         Width           =   4035
      End
      Begin VB.TextBox txtDomicilioOtroCalle3 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   98
         Top             =   1680
         Width           =   4035
      End
      Begin VB.TextBox txtDomicilioOtroNombre 
         Height          =   315
         Left            =   1260
         MaxLength       =   15
         TabIndex        =   86
         Top             =   240
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo datcboDomicilioOtroProvincia 
         Height          =   330
         Left            =   1260
         TabIndex        =   100
         Top             =   2040
         Width           =   4035
         _ExtentX        =   7117
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
      Begin MSDataListLib.DataCombo datcboDomicilioOtroLocalidad 
         Height          =   330
         Left            =   1260
         TabIndex        =   102
         Top             =   2400
         Width           =   4035
         _ExtentX        =   7117
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
      Begin VB.Label lblDomicilioOtroCalle1 
         AutoSize        =   -1  'True
         Caption         =   "&Calle:"
         Height          =   210
         Left            =   120
         TabIndex        =   87
         Top             =   660
         Width           =   390
      End
      Begin VB.Label lblDomicilioOtroNumero 
         AutoSize        =   -1  'True
         Caption         =   "N&úmero:"
         Height          =   210
         Left            =   120
         TabIndex        =   89
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label lblDomicilioOtroPiso 
         AutoSize        =   -1  'True
         Caption         =   "&Piso:"
         Height          =   210
         Left            =   2460
         TabIndex        =   91
         Top             =   1020
         Width           =   345
      End
      Begin VB.Label lblDomicilioOtroDepartamento 
         AutoSize        =   -1  'True
         Caption         =   "&Dpto.:"
         Height          =   210
         Left            =   4020
         TabIndex        =   93
         Top             =   1020
         Width           =   420
      End
      Begin VB.Label lblDomicilioOtroCodigoPostal 
         AutoSize        =   -1  'True
         Caption         =   "C.P.:"
         Height          =   210
         Left            =   120
         TabIndex        =   103
         Top             =   2820
         Width           =   330
      End
      Begin VB.Label lblDomicilioOtroProvincia 
         AutoSize        =   -1  'True
         Caption         =   "P&rovincia:"
         Height          =   210
         Left            =   120
         TabIndex        =   99
         Top             =   2100
         Width           =   705
      End
      Begin VB.Label lblDomicilioOtroCalle2 
         AutoSize        =   -1  'True
         Caption         =   "&Calle 2:"
         Height          =   210
         Left            =   120
         TabIndex        =   95
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblDomicilioOtroCalle3 
         AutoSize        =   -1  'True
         Caption         =   "&Calle 3:"
         Height          =   210
         Left            =   120
         TabIndex        =   97
         Top             =   1740
         Width           =   525
      End
      Begin VB.Label lblDomicilioOtroLocalidad 
         AutoSize        =   -1  'True
         Caption         =   "&Localidad:"
         Height          =   210
         Left            =   120
         TabIndex        =   101
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label lblDomicilioOtroNombre 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   210
         Left            =   120
         TabIndex        =   85
         Top             =   300
         Width           =   345
      End
   End
   Begin VB.Frame fraDomicilioParticular 
      Height          =   3135
      Left            =   240
      TabIndex        =   64
      Top             =   1680
      Width           =   5415
      Begin VB.CheckBox chkDomicilioParticularMailing 
         Caption         =   "Enviar correspondencia a este Domicilio"
         Height          =   210
         Left            =   1260
         TabIndex        =   83
         Top             =   2820
         Width           =   3795
      End
      Begin VB.TextBox txtDomicilioParticularCalle1 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   66
         Top             =   240
         Width           =   4035
      End
      Begin VB.TextBox txtDomicilioParticularNumero 
         Height          =   315
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   68
         Top             =   600
         Width           =   1035
      End
      Begin VB.TextBox txtDomicilioParticularPiso 
         Height          =   315
         Left            =   2940
         MaxLength       =   10
         TabIndex        =   70
         Top             =   600
         Width           =   795
      End
      Begin VB.TextBox txtDomicilioParticularDepartamento 
         Height          =   315
         Left            =   4500
         MaxLength       =   10
         TabIndex        =   72
         Top             =   600
         Width           =   795
      End
      Begin VB.TextBox txtDomicilioParticularCodigoPostal 
         Height          =   315
         Left            =   1260
         MaxLength       =   8
         TabIndex        =   82
         Top             =   2400
         Width           =   1035
      End
      Begin VB.TextBox txtDomicilioParticularCalle2 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   74
         Top             =   960
         Width           =   4035
      End
      Begin VB.TextBox txtDomicilioParticularCalle3 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   76
         Top             =   1320
         Width           =   4035
      End
      Begin MSDataListLib.DataCombo datcboDomicilioParticularProvincia 
         Height          =   330
         Left            =   1260
         TabIndex        =   78
         Top             =   1680
         Width           =   4035
         _ExtentX        =   7117
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
      Begin MSDataListLib.DataCombo datcboDomicilioParticularLocalidad 
         Height          =   330
         Left            =   1260
         TabIndex        =   80
         Top             =   2040
         Width           =   4035
         _ExtentX        =   7117
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
      Begin VB.Label lblDomicilioParticularCalle1 
         AutoSize        =   -1  'True
         Caption         =   "&Calle:"
         Height          =   210
         Left            =   120
         TabIndex        =   65
         Top             =   300
         Width           =   390
      End
      Begin VB.Label lblDomicilioParticularNumero 
         AutoSize        =   -1  'True
         Caption         =   "N&úmero:"
         Height          =   210
         Left            =   120
         TabIndex        =   67
         Top             =   660
         Width           =   600
      End
      Begin VB.Label lblDomicilioParticularPiso 
         AutoSize        =   -1  'True
         Caption         =   "&Piso:"
         Height          =   210
         Left            =   2460
         TabIndex        =   69
         Top             =   660
         Width           =   345
      End
      Begin VB.Label lblDomicilioParticularDepartamento 
         AutoSize        =   -1  'True
         Caption         =   "&Dpto.:"
         Height          =   210
         Left            =   4020
         TabIndex        =   71
         Top             =   660
         Width           =   420
      End
      Begin VB.Label lblDomicilioParticularCodigoPostal 
         AutoSize        =   -1  'True
         Caption         =   "C.P.:"
         Height          =   210
         Left            =   120
         TabIndex        =   81
         Top             =   2460
         Width           =   330
      End
      Begin VB.Label lblDomicilioParticularProvincia 
         AutoSize        =   -1  'True
         Caption         =   "P&rovincia:"
         Height          =   210
         Left            =   120
         TabIndex        =   77
         Top             =   1740
         Width           =   705
      End
      Begin VB.Label lblDomicilioParticularCalle2 
         AutoSize        =   -1  'True
         Caption         =   "&Calle 2:"
         Height          =   210
         Left            =   120
         TabIndex        =   73
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label lblDomicilioParticularCalle3 
         AutoSize        =   -1  'True
         Caption         =   "&Calle 3:"
         Height          =   210
         Left            =   120
         TabIndex        =   75
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblDomicilioParticularLocalidad 
         AutoSize        =   -1  'True
         Caption         =   "&Localidad:"
         Height          =   210
         Left            =   120
         TabIndex        =   79
         Top             =   2100
         Width           =   735
      End
   End
   Begin VB.Frame fraDomicilioLaboral 
      Height          =   3135
      Left            =   240
      TabIndex        =   44
      Top             =   1680
      Width           =   5415
      Begin VB.CheckBox chkDomicilioLaboralMailing 
         Caption         =   "Enviar correspondencia a este Domicilio"
         Height          =   210
         Left            =   1260
         TabIndex        =   63
         Top             =   2820
         Width           =   3795
      End
      Begin VB.TextBox txtDomicilioLaboralCalle1 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   46
         Top             =   240
         Width           =   4035
      End
      Begin VB.TextBox txtDomicilioLaboralNumero 
         Height          =   315
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   48
         Top             =   600
         Width           =   1035
      End
      Begin VB.TextBox txtDomicilioLaboralPiso 
         Height          =   315
         Left            =   2940
         MaxLength       =   10
         TabIndex        =   50
         Top             =   600
         Width           =   795
      End
      Begin VB.TextBox txtDomicilioLaboralDepartamento 
         Height          =   315
         Left            =   4500
         MaxLength       =   10
         TabIndex        =   52
         Top             =   600
         Width           =   795
      End
      Begin VB.TextBox txtDomicilioLaboralCodigoPostal 
         Height          =   315
         Left            =   1260
         MaxLength       =   8
         TabIndex        =   62
         Top             =   2400
         Width           =   1035
      End
      Begin VB.TextBox txtDomicilioLaboralCalle2 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   54
         Top             =   960
         Width           =   4035
      End
      Begin VB.TextBox txtDomicilioLaboralCalle3 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   56
         Top             =   1320
         Width           =   4035
      End
      Begin MSDataListLib.DataCombo datcboDomicilioLaboralProvincia 
         Height          =   330
         Left            =   1260
         TabIndex        =   58
         Top             =   1680
         Width           =   4035
         _ExtentX        =   7117
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
      Begin MSDataListLib.DataCombo datcboDomicilioLaboralLocalidad 
         Height          =   330
         Left            =   1260
         TabIndex        =   60
         Top             =   2040
         Width           =   4035
         _ExtentX        =   7117
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
      Begin VB.Label lblDomicilioLaboralCalle1 
         AutoSize        =   -1  'True
         Caption         =   "&Calle:"
         Height          =   210
         Left            =   120
         TabIndex        =   45
         Top             =   300
         Width           =   390
      End
      Begin VB.Label lblDomicilioLaboralNumero 
         AutoSize        =   -1  'True
         Caption         =   "N&úmero:"
         Height          =   210
         Left            =   120
         TabIndex        =   47
         Top             =   660
         Width           =   600
      End
      Begin VB.Label lblDomicilioLaboralPiso 
         AutoSize        =   -1  'True
         Caption         =   "&Piso:"
         Height          =   210
         Left            =   2460
         TabIndex        =   49
         Top             =   660
         Width           =   345
      End
      Begin VB.Label lblDomicilioLaboralDepartamento 
         AutoSize        =   -1  'True
         Caption         =   "&Dpto.:"
         Height          =   210
         Left            =   4020
         TabIndex        =   51
         Top             =   660
         Width           =   420
      End
      Begin VB.Label lblDomicilioLaboralCodigoPostal 
         AutoSize        =   -1  'True
         Caption         =   "C.P.:"
         Height          =   210
         Left            =   120
         TabIndex        =   61
         Top             =   2460
         Width           =   330
      End
      Begin VB.Label lblDomicilioLaboralProvincia 
         AutoSize        =   -1  'True
         Caption         =   "P&rovincia:"
         Height          =   210
         Left            =   120
         TabIndex        =   57
         Top             =   1740
         Width           =   705
      End
      Begin VB.Label lblDomicilioLaboralCalle2 
         AutoSize        =   -1  'True
         Caption         =   "&Calle 2:"
         Height          =   210
         Left            =   120
         TabIndex        =   53
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label lblDomicilioLaboralCalle3 
         AutoSize        =   -1  'True
         Caption         =   "&Calle 3:"
         Height          =   210
         Left            =   120
         TabIndex        =   55
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblDomicilioLaboralLocalidad 
         AutoSize        =   -1  'True
         Caption         =   "&Localidad:"
         Height          =   210
         Left            =   120
         TabIndex        =   59
         Top             =   2100
         Width           =   735
      End
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   5415
      Left            =   120
      TabIndex        =   138
      Top             =   1020
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9551
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "GENERAL"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Domicilio Laboral"
            Key             =   "DOMICILIO_LABORAL"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Domicilio Particular"
            Key             =   "DOMICILIO_PARTICULAR"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Otro Domicilio"
            Key             =   "DOMICILIO_OTRO"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Internet"
            Key             =   "INTERNET"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Datos Adicionales"
            Key             =   "DATOS_ADICIONALES"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos del Contacto"
      Height          =   210
      Left            =   780
      TabIndex        =   137
      Top             =   300
      Width           =   2550
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "ContactoPropiedad.frx":4216
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmContactoPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mContacto As Contacto
Private mNew As Boolean

Private mLoading As Boolean


Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mContacto
End Sub

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef Contacto As Contacto)
    Set mContacto = Contacto
    Set Contacto = Nothing
    mNew = (mContacto.IDContacto = 0)
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    mLoading = True
    
    With mContacto
        'GENERAL
        If Not CSM_Control_DataCombo.FillFromSQL(datcboTitulo, "SELECT IDTitulo, Nombre FROM Titulo ORDER BY Nombre", "IDTitulo", "Nombre", "Títulos") Then
            Unload Me
            Exit Sub
        End If
        datcboTitulo.Text = .Titulo
        txtApellido.Text = .Apellido
        txtNombre.Text = .Nombre
        txtCompania.Text = .Compania
        txtTituloLaboral.Text = .TituloLaboral
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboTelefonoTipo(1), "(SELECT 0 AS IDTelefonoTipo, '--------' AS Nombre, 1 AS Orden FROM TelefonoTipo) UNION (SELECT IDTelefonoTipo, Nombre, 2 AS Orden FROM TelefonoTipo WHERE Activo = 1) ORDER BY Orden, IDTelefonoTipo", "IDTelefonoTipo", "Nombre", "Tipos de Teléfonos", cscpItemOrfirst, .IDTelefono1Tipo) Then
            Unload Me
            Exit Sub
        End If
        txtTelefonoTipoOtro(1).Text = .Telefono1TipoOtro
        txtTelefonoArea(1).Text = .Telefono1Area
        txtTelefonoNumero(1).Text = .Telefono1Numero
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboTelefonoTipo(2), "(SELECT 0 AS IDTelefonoTipo, '--------' AS Nombre, 1 AS Orden FROM TelefonoTipo) UNION (SELECT IDTelefonoTipo, Nombre, 2 AS Orden FROM TelefonoTipo WHERE Activo = 1) ORDER BY Orden, IDTelefonoTipo", "IDTelefonoTipo", "Nombre", "Tipos de Teléfonos", cscpItemOrfirst, .IDTelefono2Tipo) Then
            Unload Me
            Exit Sub
        End If
        txtTelefonoTipoOtro(2).Text = .Telefono2TipoOtro
        txtTelefonoArea(2).Text = .Telefono2Area
        txtTelefonoNumero(2).Text = .Telefono2Numero
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboTelefonoTipo(3), "(SELECT 0 AS IDTelefonoTipo, '--------' AS Nombre, 1 AS Orden FROM TelefonoTipo) UNION (SELECT IDTelefonoTipo, Nombre, 2 AS Orden FROM TelefonoTipo WHERE Activo = 1) ORDER BY Orden, IDTelefonoTipo", "IDTelefonoTipo", "Nombre", "Tipos de Teléfonos", cscpItemOrfirst, .IDTelefono3Tipo) Then
            Unload Me
            Exit Sub
        End If
        txtTelefonoTipoOtro(3).Text = .Telefono3TipoOtro
        txtTelefonoArea(3).Text = .Telefono3Area
        txtTelefonoNumero(3).Text = .Telefono3Numero
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboTelefonoTipo(4), "(SELECT 0 AS IDTelefonoTipo, '--------' AS Nombre, 1 AS Orden FROM TelefonoTipo) UNION (SELECT IDTelefonoTipo, Nombre, 2 AS Orden FROM TelefonoTipo WHERE Activo = 1) ORDER BY Orden, IDTelefonoTipo", "IDTelefonoTipo", "Nombre", "Tipos de Teléfonos", cscpItemOrfirst, .IDTelefono4Tipo) Then
            Unload Me
            Exit Sub
        End If
        txtTelefonoTipoOtro(4).Text = .Telefono4TipoOtro
        txtTelefonoArea(4).Text = .Telefono4Area
        txtTelefonoNumero(4).Text = .Telefono4Numero
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboTelefonoTipo(5), "(SELECT 0 AS IDTelefonoTipo, '--------' AS Nombre, 1 AS Orden FROM TelefonoTipo) UNION (SELECT IDTelefonoTipo, Nombre, 2 AS Orden FROM TelefonoTipo WHERE Activo = 1) ORDER BY Orden, IDTelefonoTipo", "IDTelefonoTipo", "Nombre", "Tipos de Teléfonos", cscpItemOrfirst, .IDTelefono5Tipo) Then
            Unload Me
            Exit Sub
        End If
        txtTelefonoTipoOtro(5).Text = .Telefono5TipoOtro
        txtTelefonoArea(5).Text = .Telefono5Area
        txtTelefonoNumero(5).Text = .Telefono5Numero
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboContactoGrupo, "SELECT IDContactoGrupo, Nombre FROM ContactoGrupo WHERE Activo = 1 ORDER BY Nombre", "IDContactoGrupo", "Nombre", "Grupos de Contactos", cscpItemOrNone, .IDContactoGrupo) Then
            Unload Me
            Exit Sub
        End If
        
        'DOMICILIO LABORAL
        txtDomicilioLaboralCalle1.Text = .DomicilioLaboralCalle1
        txtDomicilioLaboralNumero.Text = .DomicilioLaboralNumero
        txtDomicilioLaboralPiso.Text = .DomicilioLaboralPiso
        txtDomicilioLaboralDepartamento.Text = .DomicilioLaboralDepartamento
        txtDomicilioLaboralCalle2.Text = .DomicilioLaboralCalle2
        txtDomicilioLaboralCalle3.Text = .DomicilioLaboralCalle3
        If Not CSM_Control_DataCombo.FillFromSQL(datcboDomicilioLaboralProvincia, "(SELECT ' ' AS IDProvincia, '----------' AS Nombre) UNION (SELECT IDProvincia, Nombre FROM Provincia) ORDER BY Nombre", "IDProvincia", "Nombre", "Provincias", cscpItemOrfirst, .DomicilioLaboralIDProvincia) Then
            Unload Me
            Exit Sub
        End If
        datcboDomicilioLaboralLocalidad.BoundText = .DomicilioLaboralIDLocalidad
        txtDomicilioLaboralCodigoPostal.Text = .DomicilioLaboralCodigoPostal
        chkDomicilioLaboralMailing.Value = IIf(.DomicilioLaboralMailing, vbChecked, vbUnchecked)
        
        'DOMICILIO PARTICULAR
        txtDomicilioParticularCalle1.Text = .DomicilioParticularCalle1
        txtDomicilioParticularNumero.Text = .DomicilioParticularNumero
        txtDomicilioParticularPiso.Text = .DomicilioParticularPiso
        txtDomicilioParticularDepartamento.Text = .DomicilioParticularDepartamento
        txtDomicilioParticularCalle2.Text = .DomicilioParticularCalle2
        txtDomicilioParticularCalle3.Text = .DomicilioParticularCalle3
        If Not CSM_Control_DataCombo.FillFromSQL(datcboDomicilioParticularProvincia, "(SELECT ' ' AS IDProvincia, '----------' AS Nombre) UNION (SELECT IDProvincia, Nombre FROM Provincia) ORDER BY Nombre", "IDProvincia", "Nombre", "Provincias", cscpItemOrfirst, .DomicilioParticularIDProvincia) Then
            Unload Me
            Exit Sub
        End If
        datcboDomicilioParticularLocalidad.BoundText = .DomicilioParticularIDLocalidad
        txtDomicilioParticularCodigoPostal.Text = .DomicilioParticularCodigoPostal
        chkDomicilioParticularMailing.Value = IIf(.DomicilioParticularMailing, vbChecked, vbUnchecked)
        
        'DOMICILIO OTRO
        txtDomicilioOtroNombre.Text = .DomicilioOtroNombre
        txtDomicilioOtroCalle1.Text = .DomicilioOtroCalle1
        txtDomicilioOtroNumero.Text = .DomicilioOtroNumero
        txtDomicilioOtroPiso.Text = .DomicilioOtroPiso
        txtDomicilioOtroDepartamento.Text = .DomicilioOtroDepartamento
        txtDomicilioOtroCalle2.Text = .DomicilioOtroCalle2
        txtDomicilioOtroCalle3.Text = .DomicilioOtroCalle3
        If Not CSM_Control_DataCombo.FillFromSQL(datcboDomicilioOtroProvincia, "(SELECT ' ' AS IDProvincia, '----------' AS Nombre) UNION (SELECT IDProvincia, Nombre FROM Provincia) ORDER BY Nombre", "IDProvincia", "Nombre", "Provincias", cscpItemOrfirst, .DomicilioOtroIDProvincia) Then
            Unload Me
            Exit Sub
        End If
        datcboDomicilioOtroLocalidad.BoundText = .DomicilioOtroIDLocalidad
        txtDomicilioOtroCodigoPostal.Text = .DomicilioOtroCodigoPostal
        chkDomicilioOtroMailing.Value = IIf(.DomicilioOtroMailing, vbChecked, vbUnchecked)
        
        'INTERNET
        txtEmail(1).Text = .Email1
        txtEmailNombre(1).Text = .Email1Nombre
        txtEmail(2).Text = .Email2
        txtEmailNombre(2).Text = .Email2Nombre
        txtEmail(3).Text = .Email3
        txtEmailNombre(3).Text = .Email3Nombre
        txtPaginaWeb.Text = .PaginaWeb
        
        'DATOS ADICIONALES
        txtSobreNombre.Text = .SobreNombre
        If .FechaNacimiento = DATE_TIME_FIELD_NULL_VALUE Then
            dtpFechaNacimiento.Value = Date
            dtpFechaNacimiento.Value = Null
        Else
            dtpFechaNacimiento.Value = .FechaNacimiento
        End If
        txtAsistente.Text = .Asistente
        txtNotas.Text = .Notas
        chkActivo.Value = IIf(.Activo, vbChecked, vbUnchecked)
    End With
    
    ShowControls
        
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

Private Sub cmdContactoGrupo_Click()
    If pCPermiso.GotPermission(PERMISO_CONTACTO_GRUPO) Then
        Screen.MousePointer = vbHourglass
        frmContactoGrupo.Show
        On Error Resume Next
        Set frmContactoGrupo.lvwData.SelectedItem = frmContactoGrupo.lvwData.ListItems(KEY_STRINGER & datcboContactoGrupo.BoundText)
        frmContactoGrupo.lvwData.SelectedItem.EnsureVisible
        If frmContactoGrupo.WindowState = vbMinimized Then
            frmContactoGrupo.WindowState = vbNormal
        End If
        frmContactoGrupo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdOK_Click()
    Dim cmdTitulo As ADODB.command
    Dim Index As Integer
    
    If Trim(txtApellido.Text) = "" And Trim(txtNombre.Text) = "" And Trim(txtCompania.Text) = "" Then
        Set tabMain.SelectedItem = tabMain.Tabs("GENERAL")
        MsgBox "Debe ingresar el Apellido, el Nombre o la Compañía del Contacto.", vbInformation, App.Title
        txtApellido.SetFocus
        Exit Sub
    End If
    
    If Val(datcboContactoGrupo.BoundText) = 0 Then
        Set tabMain.SelectedItem = tabMain.Tabs("GENERAL")
        MsgBox "Debe seleccionar el Grupo.", vbInformation, App.Title
        datcboContactoGrupo.SetFocus
        Exit Sub
    End If
       
    For Index = 1 To 3
        If Trim(txtEmail(Index).Text) <> "" Then
            If InStr(1, txtEmail(Index).Text, "@") = 0 Then
                Set tabMain.SelectedItem = tabMain.Tabs("INTERNET")
                MsgBox "La Dirección de E-mail " & Index & " del Contacto es incorrecta porque no contiene la arroba '@'.", vbInformation, App.Title
                txtEmail(Index).SetFocus
                Exit Sub
            End If
        End If
    Next Index
    
    With mContacto
        'GENERAL
        .Titulo = Left(datcboTitulo.Text, 10)
        .Apellido = txtApellido.Text
        .Nombre = txtNombre.Text
        .Compania = txtCompania.Text
        .TituloLaboral = txtTituloLaboral.Text
        
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
        
        .IDContactoGrupo = Val(datcboContactoGrupo.BoundText)
        
        'DOMICILIO LABORAL
        .DomicilioLaboralCalle1 = txtDomicilioLaboralCalle1.Text
        .DomicilioLaboralNumero = txtDomicilioLaboralNumero.Text
        .DomicilioLaboralPiso = txtDomicilioLaboralPiso.Text
        .DomicilioLaboralDepartamento = txtDomicilioLaboralDepartamento.Text
        .DomicilioLaboralCalle2 = txtDomicilioLaboralCalle2.Text
        .DomicilioLaboralCalle3 = txtDomicilioLaboralCalle3.Text
        .DomicilioLaboralCodigoPostal = txtDomicilioLaboralCodigoPostal.Text
        .DomicilioLaboralIDLocalidad = Val(datcboDomicilioLaboralLocalidad.BoundText)
        .DomicilioLaboralIDProvincia = datcboDomicilioLaboralProvincia.BoundText
        .DomicilioLaboralMailing = (chkDomicilioLaboralMailing.Value = vbChecked)
        
        'DOMICILIO PARTICULAR
        .DomicilioParticularCalle1 = txtDomicilioParticularCalle1.Text
        .DomicilioParticularNumero = txtDomicilioParticularNumero.Text
        .DomicilioParticularPiso = txtDomicilioParticularPiso.Text
        .DomicilioParticularDepartamento = txtDomicilioParticularDepartamento.Text
        .DomicilioParticularCalle2 = txtDomicilioParticularCalle2.Text
        .DomicilioParticularCalle3 = txtDomicilioParticularCalle3.Text
        .DomicilioParticularCodigoPostal = txtDomicilioParticularCodigoPostal.Text
        .DomicilioParticularIDLocalidad = Val(datcboDomicilioParticularLocalidad.BoundText)
        .DomicilioParticularIDProvincia = datcboDomicilioParticularProvincia.BoundText
        .DomicilioParticularMailing = (chkDomicilioParticularMailing.Value = vbChecked)
        
        'DOMILIO OTRO
        .DomicilioOtroNombre = txtDomicilioOtroNombre.Text
        .DomicilioOtroCalle1 = txtDomicilioOtroCalle1.Text
        .DomicilioOtroNumero = txtDomicilioOtroNumero.Text
        .DomicilioOtroPiso = txtDomicilioOtroPiso.Text
        .DomicilioOtroDepartamento = txtDomicilioOtroDepartamento.Text
        .DomicilioOtroCalle2 = txtDomicilioOtroCalle2.Text
        .DomicilioOtroCalle3 = txtDomicilioOtroCalle3.Text
        .DomicilioOtroCodigoPostal = txtDomicilioOtroCodigoPostal.Text
        .DomicilioOtroIDLocalidad = Val(datcboDomicilioOtroLocalidad.BoundText)
        .DomicilioOtroIDProvincia = datcboDomicilioOtroProvincia.BoundText
        .DomicilioOtroMailing = (chkDomicilioOtroMailing.Value = vbChecked)
        
        'INTERNET
        .Email1 = txtEmail(1).Text
        .Email1Nombre = txtEmailNombre(1).Text
        .Email2 = txtEmail(2).Text
        .Email2Nombre = txtEmailNombre(2).Text
        .Email3 = txtEmail(3).Text
        .Email3Nombre = txtEmailNombre(3).Text
        .PaginaWeb = txtPaginaWeb.Text
        
        'DATOS ADICIONALES
        .SobreNombre = txtSobreNombre.Text
        .FechaNacimiento = IIf(IsNull(dtpFechaNacimiento.Value), DATE_TIME_FIELD_NULL_VALUE, dtpFechaNacimiento.Value)
        .Asistente = txtAsistente.Text
        .Notas = txtNotas.Text
        .Activo = (chkActivo.Value = vbChecked)
        
        If mNew Then
            If Not .AddNew() Then
                Exit Sub
            End If
        Else
            If Not .Update() Then
                Exit Sub
            End If
        End If
        
        'VERIFICO SI EXISTE EL TITULO, PARA AGREGARLO SI NO
        If Trim(datcboTitulo.Text) <> "" Then
            Set cmdTitulo = New ADODB.command
            Set cmdTitulo.ActiveConnection = pDatabase.Connection
            cmdTitulo.CommandText = "sp_Titulo_Check"
            cmdTitulo.CommandType = adCmdStoredProc
            cmdTitulo.Parameters.Append cmdTitulo.CreateParameter("Nombre_FILTER", adVarChar, adParamInput, 50, .Titulo)
            cmdTitulo.Execute
            Set cmdTitulo = Nothing
        End If
    End With

    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mContacto = Nothing
    Set frmContactoPropiedad = Nothing
End Sub

Private Sub tabMain_Click()
    ShowControls
End Sub

'///////////////////////////////////////////////////////////////////
'GENERAL
Private Sub datcboTitulo_LostFocus()
    datcboTitulo.Text = CleanInvalidSpaces(datcboTitulo.Text)
End Sub

Private Sub txtApellido_GotFocus()
    CSM_Control_TextBox.SelAllText txtApellido
End Sub

Private Sub txtApellido_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtApellido_Change()
    SetCaption
End Sub

Private Sub txtApellido_LostFocus()
    txtApellido.Text = UCase(txtApellido.Text)
    txtApellido.Text = CleanInvalidSpaces(txtApellido.Text)
End Sub

Private Sub txtNombre_GotFocus()
    CSM_Control_TextBox.SelAllText txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombre_Change()
    SetCaption
End Sub

Private Sub txtNombre_LostFocus()
    txtNombre.Text = UCase(txtNombre.Text)
    txtNombre.Text = CleanInvalidSpaces(txtNombre.Text)
End Sub

Private Sub txtCompania_GotFocus()
    CSM_Control_TextBox.SelAllText txtCompania
End Sub

Private Sub txtCompania_Change()
    SetCaption
End Sub

Private Sub txtCompania_LostFocus()
    txtCompania.Text = CleanInvalidSpaces(txtCompania.Text)
End Sub

Private Sub txtTituloLaboral_GotFocus()
    CSM_Control_TextBox.SelAllText txtTituloLaboral
End Sub

Private Sub txtTituloLaboral_LostFocus()
    txtTituloLaboral.Text = CleanInvalidSpaces(txtTituloLaboral.Text)
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

'///////////////////////////////////////////////////////////////////
'DOMICILIO LABORAL
Private Sub txtDomicilioLaboralCalle1_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioLaboralCalle1
End Sub

Private Sub txtDomicilioLaboralCalle1_LostFocus()
    txtDomicilioLaboralCalle1.Text = CleanInvalidSpaces(txtDomicilioLaboralCalle1.Text)
End Sub

Private Sub txtDomicilioLaboralNumero_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioLaboralNumero
End Sub

Private Sub txtDomicilioLaboralNumero_LostFocus()
    txtDomicilioLaboralNumero.Text = CleanInvalidSpaces(txtDomicilioLaboralNumero.Text)
End Sub

Private Sub txtDomicilioLaboralPiso_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioLaboralPiso
End Sub

Private Sub txtDomicilioLaboralPiso_LostFocus()
    txtDomicilioLaboralPiso.Text = CleanInvalidSpaces(txtDomicilioLaboralPiso.Text)
End Sub

Private Sub txtDomicilioLaboralDepartamento_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioLaboralDepartamento
End Sub

Private Sub txtDomicilioLaboralDepartamento_LostFocus()
    txtDomicilioLaboralDepartamento.Text = CleanInvalidSpaces(txtDomicilioLaboralDepartamento.Text)
End Sub

Private Sub txtDomicilioLaboralCalle2_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioLaboralCalle2
End Sub

Private Sub txtDomicilioLaboralCalle2_LostFocus()
    txtDomicilioLaboralCalle2.Text = CleanInvalidSpaces(txtDomicilioLaboralCalle2.Text)
End Sub

Private Sub txtDomicilioLaboralCalle3_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioLaboralCalle3
End Sub

Private Sub txtDomicilioLaboralCalle3_LostFocus()
    txtDomicilioLaboralCalle3.Text = CleanInvalidSpaces(txtDomicilioLaboralCalle3.Text)
End Sub

Private Sub txtDomicilioLaboralCodigoPostal_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioLaboralCodigoPostal
End Sub

Private Sub txtDomicilioLaboralCodigoPostal_LostFocus()
    txtDomicilioLaboralCodigoPostal.Text = CleanInvalidSpaces(txtDomicilioLaboralCodigoPostal.Text)
End Sub

Private Sub datcboDomicilioLaboralProvincia_Change()
    datcboDomicilioLaboralLocalidad.BoundText = ""
    Call CSM_Control_DataCombo.FillFromSQL(datcboDomicilioLaboralLocalidad, "SELECT IDLocalidad, Nombre FROM Localidad WHERE IDProvincia = '" & datcboDomicilioLaboralProvincia.BoundText & "' ORDER BY Nombre", "IDLocalidad", "Nombre", "Localidades", cscpFirstIfUnique)
End Sub

Private Sub datcboDomicilioLaboralLocalidad_Change()
    Dim Localidad As Localidad
    
    If Val(datcboDomicilioLaboralLocalidad.BoundText) = 0 Then
        txtDomicilioLaboralCodigoPostal.Text = ""
    Else
        Set Localidad = New Localidad
        
        Localidad.IDProvincia = datcboDomicilioLaboralProvincia.BoundText
        Localidad.IDLocalidad = Val(datcboDomicilioLaboralLocalidad.BoundText)
        If Localidad.Load() Then
            txtDomicilioLaboralCodigoPostal.Text = IIf(Localidad.CodigoPostal = 0, "", Localidad.CodigoPostal)
        End If
        
        Set Localidad = Nothing
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'DOMICILIO PARTICULAR
Private Sub txtDomicilioParticularCalle1_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioParticularCalle1
End Sub

Private Sub txtDomicilioParticularCalle1_LostFocus()
    txtDomicilioParticularCalle1.Text = CleanInvalidSpaces(txtDomicilioParticularCalle1.Text)
End Sub

Private Sub txtDomicilioParticularNumero_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioParticularNumero
End Sub

Private Sub txtDomicilioParticularNumero_LostFocus()
    txtDomicilioParticularNumero.Text = CleanInvalidSpaces(txtDomicilioParticularNumero.Text)
End Sub

Private Sub txtDomicilioParticularPiso_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioParticularPiso
End Sub

Private Sub txtDomicilioParticularPiso_LostFocus()
    txtDomicilioParticularPiso.Text = CleanInvalidSpaces(txtDomicilioParticularPiso.Text)
End Sub

Private Sub txtDomicilioParticularDepartamento_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioParticularDepartamento
End Sub

Private Sub txtDomicilioParticularDepartamento_LostFocus()
    txtDomicilioParticularDepartamento.Text = CleanInvalidSpaces(txtDomicilioParticularDepartamento.Text)
End Sub

Private Sub txtDomicilioParticularCalle2_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioParticularCalle2
End Sub

Private Sub txtDomicilioParticularCalle2_LostFocus()
    txtDomicilioParticularCalle2.Text = CleanInvalidSpaces(txtDomicilioParticularCalle2.Text)
End Sub

Private Sub txtDomicilioParticularCalle3_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioParticularCalle3
End Sub

Private Sub txtDomicilioParticularCalle3_LostFocus()
    txtDomicilioParticularCalle3.Text = CleanInvalidSpaces(txtDomicilioParticularCalle3.Text)
End Sub

Private Sub txtDomicilioParticularCodigoPostal_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioParticularCodigoPostal
End Sub

Private Sub txtDomicilioParticularCodigoPostal_LostFocus()
    txtDomicilioParticularCodigoPostal.Text = CleanInvalidSpaces(txtDomicilioParticularCodigoPostal.Text)
End Sub

Private Sub datcboDomicilioParticularProvincia_Change()
    datcboDomicilioParticularLocalidad.BoundText = ""
    If Not CSM_Control_DataCombo.FillFromSQL(datcboDomicilioParticularLocalidad, "SELECT IDLocalidad, Nombre FROM Localidad WHERE IDProvincia = '" & datcboDomicilioParticularProvincia.BoundText & "' ORDER BY Nombre", "IDLocalidad", "Nombre", "Localidades", cscpFirstIfUnique) Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub datcboDomicilioParticularLocalidad_Change()
    Dim Localidad As Localidad
    
    If Val(datcboDomicilioParticularLocalidad.BoundText) = 0 Then
        txtDomicilioParticularCodigoPostal.Text = ""
    Else
        Set Localidad = New Localidad
        
        Localidad.IDProvincia = datcboDomicilioParticularProvincia.BoundText
        Localidad.IDLocalidad = Val(datcboDomicilioParticularLocalidad.BoundText)
        If Localidad.Load() Then
            txtDomicilioParticularCodigoPostal.Text = IIf(Localidad.CodigoPostal = 0, "", Localidad.CodigoPostal)
        End If
        
        Set Localidad = Nothing
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'DOMICILIO OTRO
Private Sub txtDomicilioOtroNombre_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioOtroNombre
End Sub

Private Sub txtDomicilioOtroCalle1_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioOtroCalle1
End Sub

Private Sub txtDomicilioOtroCalle1_LostFocus()
    txtDomicilioOtroCalle1.Text = CleanInvalidSpaces(txtDomicilioOtroCalle1.Text)
End Sub

Private Sub txtDomicilioOtroNumero_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioOtroNumero
End Sub

Private Sub txtDomicilioOtroNumero_LostFocus()
    txtDomicilioOtroNumero.Text = CleanInvalidSpaces(txtDomicilioOtroNumero.Text)
End Sub

Private Sub txtDomicilioOtroPiso_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioOtroPiso
End Sub

Private Sub txtDomicilioOtroPiso_LostFocus()
    txtDomicilioOtroPiso.Text = CleanInvalidSpaces(txtDomicilioOtroPiso.Text)
End Sub

Private Sub txtDomicilioOtroDepartamento_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioOtroDepartamento
End Sub

Private Sub txtDomicilioOtroDepartamento_LostFocus()
    txtDomicilioOtroDepartamento.Text = CleanInvalidSpaces(txtDomicilioOtroDepartamento.Text)
End Sub

Private Sub txtDomicilioOtroCalle2_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioOtroCalle2
End Sub

Private Sub txtDomicilioOtroCalle2_LostFocus()
    txtDomicilioOtroCalle2.Text = CleanInvalidSpaces(txtDomicilioOtroCalle2.Text)
End Sub

Private Sub txtDomicilioOtroCalle3_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioOtroCalle3
End Sub

Private Sub txtDomicilioOtroCalle3_LostFocus()
    txtDomicilioOtroCalle3.Text = CleanInvalidSpaces(txtDomicilioOtroCalle3.Text)
End Sub

Private Sub txtDomicilioOtroCodigoPostal_GotFocus()
    CSM_Control_TextBox.SelAllText txtDomicilioOtroCodigoPostal
End Sub

Private Sub txtDomicilioOtroCodigoPostal_LostFocus()
    txtDomicilioOtroCodigoPostal.Text = CleanInvalidSpaces(txtDomicilioOtroCodigoPostal.Text)
End Sub

Private Sub datcboDomicilioOtroProvincia_Change()
    datcboDomicilioOtroLocalidad.BoundText = ""
    If Not CSM_Control_DataCombo.FillFromSQL(datcboDomicilioOtroLocalidad, "SELECT IDLocalidad, Nombre FROM Localidad WHERE IDProvincia = '" & datcboDomicilioOtroProvincia.BoundText & "' ORDER BY Nombre", "IDLocalidad", "Nombre", "Localidades", cscpFirstIfUnique) Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub datcboDomicilioOtroLocalidad_Change()
    Dim Localidad As Localidad
    
    If Val(datcboDomicilioOtroLocalidad.BoundText) = 0 Then
        txtDomicilioOtroCodigoPostal.Text = ""
    Else
        Set Localidad = New Localidad
        
        Localidad.IDProvincia = datcboDomicilioOtroProvincia.BoundText
        Localidad.IDLocalidad = Val(datcboDomicilioOtroLocalidad.BoundText)
        If Localidad.Load() Then
            txtDomicilioOtroCodigoPostal.Text = IIf(Localidad.CodigoPostal = 0, "", Localidad.CodigoPostal)
        End If
        
        Set Localidad = Nothing
    End If
End Sub

'///////////////////////////////////////////////////////////////////
'INTERNET
Private Sub txtEmail_GotFocus(Index As Integer)
    CSM_Control_TextBox.SelAllText txtEmail(Index)
End Sub

Private Sub txtEmail_Change(Index As Integer)
    cmdEmail(Index).Visible = (Trim(txtEmail(Index).Text <> ""))
End Sub

Private Sub txtEmail_LostFocus(Index As Integer)
    txtEmail(Index).Text = CleanInvalidSpaces(txtEmail(Index).Text)
End Sub

Private Sub cmdEmail_Click(Index As Integer)
    CSM_Instance.Execute Me.hwnd, "mailto:" & txtEmail(Index).Text
End Sub

Private Sub txtEmailNombre_GotFocus(Index As Integer)
    CSM_Control_TextBox.SelAllText txtEmailNombre(Index)
End Sub

Private Sub txtEmailNombre_LostFocus(Index As Integer)
    txtEmailNombre(Index).Text = CleanInvalidSpaces(txtEmailNombre(Index).Text)
End Sub

Private Sub txtPaginaWeb_GotFocus()
    CSM_Control_TextBox.SelAllText txtPaginaWeb
End Sub

Private Sub txtPaginaWeb_Change()
    cmdPaginaWeb.Visible = (Trim(txtPaginaWeb.Text <> ""))
End Sub

Private Sub txtPaginaWeb_LostFocus()
    txtPaginaWeb.Text = CleanInvalidSpaces(txtPaginaWeb.Text)
End Sub

Private Sub cmdPaginaWeb_Click()
    CSM_Instance.Execute Me.hwnd, "http://" & txtPaginaWeb.Text
End Sub

'///////////////////////////////////////////////////////////////////
'DATOS ADICIONALES
Private Sub txtSobreNombre_GotFocus()
    CSM_Control_TextBox.SelAllText txtSobreNombre
End Sub

Private Sub txtSobreNombre_LostFocus()
    txtSobreNombre.Text = CleanInvalidSpaces(txtSobreNombre.Text)
End Sub

Private Sub txtAsistente_GotFocus()
    CSM_Control_TextBox.SelAllText txtAsistente
End Sub

Private Sub txtAsistente_LostFocus()
    txtAsistente.Text = CleanInvalidSpaces(txtAsistente.Text)
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
    txtNotas.Text = CleanInvalidSpaces(txtNotas.Text)
End Sub

'///////////////////////////////////////////////////////////////////
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

Public Sub FillComboBoxContactoGrupo()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboContactoGrupo.BoundText)
    Set recData = datcboContactoGrupo.RowSource
    recData.Requery
    Set recData = Nothing
    datcboContactoGrupo.BoundText = KeySave
End Sub

Private Sub ShowControls()
    fraGeneral.Visible = (tabMain.SelectedItem.Key = "GENERAL")
    fraDomicilioLaboral.Visible = (tabMain.SelectedItem.Key = "DOMICILIO_LABORAL")
    fraDomicilioParticular.Visible = (tabMain.SelectedItem.Key = "DOMICILIO_PARTICULAR")
    fraDomicilioOtro.Visible = (tabMain.SelectedItem.Key = "DOMICILIO_OTRO")
    fraInternet.Visible = (tabMain.SelectedItem.Key = "INTERNET")
    fraDatosAdicionales.Visible = (tabMain.SelectedItem.Key = "DATOS_ADICIONALES")
End Sub

Private Sub SetCaption()
    If Trim(txtApellido.Text) <> "" Then
        If Trim(txtNombre.Text) <> "" Then
            Caption = "Propiedades de " & txtApellido.Text & ", " & txtNombre.Text
        Else
            Caption = "Propiedades de " & txtApellido.Text
        End If
    ElseIf Trim(txtNombre.Text) <> "" Then
        Caption = "Propiedades de " & txtNombre.Text
    ElseIf Trim(txtCompania.Text) = "" Then
        Caption = "Propiedades de " & txtCompania.Text
    Else
        Caption = "Propiedades"
    End If
End Sub
