VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmComision 
   Caption         =   "Comisiones"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13080
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Comision.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   13080
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   60
      TabIndex        =   18
      Top             =   5880
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PIN"
            ImageIndex      =   1
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   1740
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   3069
      BandCount       =   9
      FixedOrder      =   -1  'True
      _CBWidth        =   13080
      _CBHeight       =   1740
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain"
      MinWidth1       =   5700
      MinHeight1      =   570
      Width1          =   5700
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picFilterFecha"
      MinWidth2       =   6705
      MinHeight2      =   360
      Width2          =   6705
      Key2            =   "FilterFecha"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Child3          =   "picFilterRuta"
      MinWidth3       =   3015
      MinHeight3      =   360
      Width3          =   3015
      Key3            =   "FilterRuta"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Child4          =   "picFilterPersona"
      MinWidth4       =   4710
      MinHeight4      =   360
      Width4          =   4710
      Key4            =   "FilterPersona"
      NewRow4         =   0   'False
      AllowVertical4  =   0   'False
      Child5          =   "picFilterListaPrecio"
      MinWidth5       =   4665
      MinHeight5      =   330
      Width5          =   4665
      Key5            =   "FilterListaPrecio"
      NewRow5         =   0   'False
      AllowVertical5  =   0   'False
      Child6          =   "picFilterEntregada"
      MinWidth6       =   1905
      MinHeight6      =   330
      Width6          =   1905
      Key6            =   "FilterEntregada"
      NewRow6         =   0   'False
      AllowVertical6  =   0   'False
      Child7          =   "picFilterPagada"
      MinWidth7       =   1785
      MinHeight7      =   330
      Width7          =   1785
      Key7            =   "FilterPagada"
      NewRow7         =   0   'False
      AllowVertical7  =   0   'False
      Child8          =   "picFilterRendicion"
      MinWidth8       =   6930
      MinHeight8      =   330
      Width8          =   6930
      Key8            =   "FilterRendicion"
      NewRow8         =   0   'False
      AllowVertical8  =   0   'False
      Child9          =   "picFilterMostrarAnteriores"
      MinWidth9       =   3435
      MinHeight9      =   330
      Width9          =   3435
      Key9            =   "FilterDias"
      NewRow9         =   0   'False
      AllowVertical9  =   0   'False
      Begin VB.PictureBox picFilterRendicion 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   165
         ScaleHeight     =   330
         ScaleWidth      =   9165
         TabIndex        =   38
         Top             =   1380
         Width           =   9165
         Begin VB.TextBox txtRendicionDiaSemana 
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.CommandButton cmdRendicionHoyHasta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6600
            Picture         =   "Comision.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   45
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdRendicionSiguienteHasta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6300
            Picture         =   "Comision.frx":06D4
            Style           =   1  'Graphical
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdRendicionAnteriorHasta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4560
            Picture         =   "Comision.frx":0C5E
            Style           =   1  'Graphical
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdRendicionHoyDesde 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3960
            Picture         =   "Comision.frx":11E8
            Style           =   1  'Graphical
            TabIndex        =   42
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdRendicionSiguienteDesde 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3660
            Picture         =   "Comision.frx":1332
            Style           =   1  'Graphical
            TabIndex        =   41
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdRendicionAnteriorDesde 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            Picture         =   "Comision.frx":18BC
            Style           =   1  'Graphical
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.ComboBox cboRendicion 
            Height          =   330
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   0
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtpRendicionFechaDesde 
            Height          =   315
            Left            =   2220
            TabIndex        =   47
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   151257089
            CurrentDate     =   36950
         End
         Begin MSComCtl2.DTPicker dtpRendicionFechaHasta 
            Height          =   315
            Left            =   4860
            TabIndex        =   48
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   151257089
            CurrentDate     =   36950
         End
         Begin VB.Label lblRendicionFechaAnd 
            AutoSize        =   -1  'True
            Caption         =   "y"
            Height          =   210
            Left            =   4380
            TabIndex        =   50
            Top             =   60
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblRendicion 
            AutoSize        =   -1  'True
            Caption         =   "Rendición:"
            Height          =   210
            Left            =   0
            TabIndex        =   49
            Top             =   60
            Width           =   750
         End
      End
      Begin VB.PictureBox picFilterListaPrecio 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   165
         ScaleHeight     =   330
         ScaleWidth      =   8685
         TabIndex        =   35
         Top             =   1020
         Width           =   8685
         Begin VB.ComboBox cboListaPrecio 
            Height          =   330
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   0
            Width           =   3315
         End
         Begin VB.Label lblListaPrecio 
            AutoSize        =   -1  'True
            Caption         =   "Lista de Precios:"
            Height          =   210
            Left            =   0
            TabIndex        =   37
            Top             =   60
            Width           =   1200
         End
      End
      Begin VB.PictureBox picFilterPagada 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11205
         ScaleHeight     =   330
         ScaleWidth      =   1785
         TabIndex        =   32
         Top             =   1020
         Width           =   1785
         Begin VB.ComboBox cboPagada 
            Height          =   330
            Left            =   780
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   0
            Width           =   990
         End
         Begin VB.Label lblPagada 
            AutoSize        =   -1  'True
            Caption         =   "Pagada:"
            Height          =   210
            Left            =   0
            TabIndex        =   34
            Top             =   60
            Width           =   585
         End
      End
      Begin VB.PictureBox picFilterPersona 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   3405
         ScaleHeight     =   360
         ScaleWidth      =   9585
         TabIndex        =   26
         Top             =   630
         Width           =   9585
         Begin VB.CommandButton cmdPersonaUltimo 
            Caption         =   "Ultimo"
            Height          =   315
            Left            =   4140
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   0
            Width           =   555
         End
         Begin VB.TextBox txtPersona 
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   0
            Width           =   2715
         End
         Begin VB.CommandButton cmdPersona 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3420
            Picture         =   "Comision.frx":1E46
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Buscar..."
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton cmdPersonaClear 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3780
            Picture         =   "Comision.frx":23D0
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Borrar"
            Top             =   0
            Width           =   315
         End
         Begin VB.Label lblPersona 
            AutoSize        =   -1  'True
            Caption         =   "Persona:"
            Height          =   210
            Left            =   0
            TabIndex        =   31
            Top             =   60
            Width           =   645
         End
      End
      Begin VB.PictureBox picFilterMostrarAnteriores 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9555
         ScaleHeight     =   330
         ScaleWidth      =   3435
         TabIndex        =   24
         Top             =   1380
         Width           =   3435
         Begin VB.CheckBox chkMostrarTodas 
            Caption         =   "Mostrar Comisiones anteriores a 30 días."
            Height          =   195
            Left            =   60
            TabIndex        =   25
            Top             =   90
            Width           =   3315
         End
      End
      Begin VB.PictureBox picFilterEntregada 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9075
         ScaleHeight     =   330
         ScaleWidth      =   1905
         TabIndex        =   21
         Top             =   1020
         Width           =   1905
         Begin VB.ComboBox cboEntregada 
            Height          =   330
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   0
            Width           =   990
         End
         Begin VB.Label lblEntregada 
            AutoSize        =   -1  'True
            Caption         =   "Entregada:"
            Height          =   210
            Left            =   0
            TabIndex        =   23
            Top             =   60
            Width           =   780
         End
      End
      Begin VB.PictureBox picFilterRuta 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   3015
         TabIndex        =   10
         Top             =   630
         Width           =   3015
         Begin VB.ComboBox cboRuta 
            Height          =   330
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   0
            Width           =   2550
         End
         Begin VB.Label lblRuta 
            AutoSize        =   -1  'True
            Caption         =   "Ruta:"
            Height          =   210
            Left            =   0
            TabIndex        =   11
            Top             =   60
            Width           =   375
         End
      End
      Begin VB.PictureBox picFilterFecha 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5955
         ScaleHeight     =   360
         ScaleWidth      =   7035
         TabIndex        =   4
         Top             =   135
         Width           =   7035
         Begin VB.TextBox txtDiaSemana 
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.CommandButton cmdHoyHasta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6360
            Picture         =   "Comision.frx":295A
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdSiguienteHasta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6060
            Picture         =   "Comision.frx":2AA4
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdAnteriorHasta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4320
            Picture         =   "Comision.frx":302E
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdHoyDesde 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            Picture         =   "Comision.frx":35B8
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdSiguienteDesde 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3420
            Picture         =   "Comision.frx":3702
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton cmdAnteriorDesde 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            Picture         =   "Comision.frx":3C8C
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.ComboBox cboFecha 
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   0
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   315
            Left            =   1980
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   151257089
            CurrentDate     =   36950
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   315
            Left            =   4620
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   151257089
            CurrentDate     =   36950
         End
         Begin VB.Label lblFechaAnd 
            AutoSize        =   -1  'True
            Caption         =   "y"
            Height          =   210
            Left            =   4140
            TabIndex        =   9
            Top             =   60
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblFecha 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   210
            Left            =   0
            TabIndex        =   8
            Top             =   60
            Width           =   495
         End
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   570
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1005
         ButtonWidth     =   2170
         ButtonHeight    =   1005
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Propiedades"
               Key             =   "PROPERTIES"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Asistencia"
               Key             =   "ASISTENCIA"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Pago Múltiple"
               Key             =   "PAGO_MULTIPLE"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "PRINT"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "LISTADO"
                     Object.Tag             =   "Comision_Listado"
                     Text            =   "Listado"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "REMITO"
                     Text            =   "Remito"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Rendir"
               Key             =   "RENDIR"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   5625
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   661
            MinWidth        =   661
            Key             =   "PIN"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21855
            Key             =   "TEXT"
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ListView lvwData 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1860
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "DiaSemana"
         Text            =   "Día"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Fecha"
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Hora"
         Text            =   "Hora"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "IDRuta"
         Text            =   "Ruta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Envia"
         Text            =   "Envía"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Recibe"
         Text            =   "Recibe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "Importe"
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "Pagado"
         Text            =   "Pagado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "Debe"
         Text            =   "Debe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "Origen"
         Text            =   "Origen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Key             =   "Destino"
         Text            =   "Destino"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Key             =   "Descripcion"
         Text            =   "Descripción"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Key             =   "ListaPrecios"
         Text            =   "Lista de Precios"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Key             =   "Entregado"
         Text            =   "Entregada"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mLoading As Boolean
Private mrecData As ADODB.Recordset

Public Sub FillListView(ByVal FechaHora As Date, ByVal IDRuta As String, ByVal Indice As Long)
    Dim SQL_Where As String
    
    If mLoading Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    SQL_Where = ""
    
    'PERSONAL
    If pPersonal Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.Personal = 0"
    End If
    
    SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "ViajeDetalle.OcupanteTipo = '" & OCUPANTE_TIPO_COMISION & "' AND ViajeDetalle.Estado = '" & VIAJE_DETALLE_ESTADO_CONFIRMADO & "'"
    
    'FECHA
    If cboFecha.ListIndex > 0 Then
        If cboFecha.ListIndex < 4 Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "convert(char(10), ViajeDetalle.FechaHora, 111) " & cboFecha.Text & " '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & "'"
        Else
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "convert(char(10), ViajeDetalle.FechaHora, 111) BETWEEN '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & "' AND '" & Format(dtpFechaHasta.Value, "yyyy/mm/dd") & "'"
        End If
    End If
    
    'RUTA
    If cboRuta.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "ViajeDetalle.IDRuta = '" & ReplaceQuote(cboRuta.Text) & "'"
    Else
        If pCPermiso.RutaWhere <> "" Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "ViajeDetalle")
        End If
    End If
    
    'LISTA DE PRECIOS
    If cboListaPrecio.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "ViajeDetalle.IDListaPrecio = " & cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
    Else
        If pCPermiso.ListaPrecioWhere <> "" Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & Replace(pCPermiso.ListaPrecioWhere, "%TABLENAME%", "ViajeDetalle")
        End If
    End If
    
    'PERSONAS
    If Val(txtPersona.Tag) > 0 Then
        If cboPagada.ListIndex > 0 Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "((ViajeDetalle.IDPersona = " & Val(txtPersona.Tag) & " AND ViajeDetalle.PagaQuienRecibe = 0) OR (ViajeDetalle.IDPersonaRecibe = " & Val(txtPersona.Tag) & " AND ViajeDetalle.PagaQuienRecibe = 1))"
        Else
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "(ViajeDetalle.IDPersona = " & Val(txtPersona.Tag) & " OR ViajeDetalle.IDPersonaRecibe = " & Val(txtPersona.Tag) & ")"
        End If
    End If
    
    'ENTREGADA
    If cboEntregada.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "ViajeDetalle.Entregada = " & IIf(cboEntregada.ListIndex = 1, 1, 0)
    End If
    
    'PAGADA
    If cboPagada.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente " & IIf(cboPagada.ListIndex = 1, ">", "=") & " 0"
    End If
    
    'RENDICION
    If cboRendicion.ListIndex > 0 Then
        If cboRendicion.ListIndex = 1 Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "ViajeDetalle_Comision.RendicionFechaHora IS NULL"
        ElseIf cboRendicion.ListIndex < 5 Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "convert(char(10), ViajeDetalle_Comision.RendicionFechaHora, 111) " & cboRendicion.Text & " '" & Format(dtpRendicionFechaDesde.Value, "yyyy/mm/dd") & "'"
        Else
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "convert(char(10), ViajeDetalle_Comision.RendicionFechaHora, 111) BETWEEN '" & Format(dtpRendicionFechaDesde.Value, "yyyy/mm/dd") & "' AND '" & Format(dtpRendicionFechaHasta.Value, "yyyy/mm/dd") & "'"
        End If
    End If
    
    'MOSTRAR TODAS
    If chkMostrarTodas.Value = vbUnchecked Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.FechaHora BETWEEN getdate() - 30 AND getdate() + 7"
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set mrecData = New ADODB.Recordset
    mrecData.Source = "SELECT datepart(weekday, ViajeDetalle.FechaHora) AS DiaSemana, ViajeDetalle.FechaHora, convert(char(10), ViajeDetalle.FechaHora, 103) AS Fecha, convert(char(5), ViajeDetalle.FechaHora, 108) AS Hora, ViajeDetalle.IDRuta, ViajeDetalle.Indice, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, ISNULL(PersonaRecibe.Apellido + (CASE ISNULL(PersonaRecibe.Nombre, '') WHEN '' THEN '' ELSE ', ' + PersonaRecibe.Nombre END), '') + ISNULL(ViajeDetalle.Recibe, '') AS Recibe, ViajeDetalle.Importe, ViajeDetalle.ImporteContado + ViajeDetalle.ImporteCuentaCorriente AS ImportePagado, ViajeDetalle.Importe - ViajeDetalle.ImporteContado - ViajeDetalle.ImporteCuentaCorriente AS Debe, Lugar_Origen.Nombre AS Origen, ViajeDetalle.Sube, Lugar_Destino.Nombre AS Destino, ViajeDetalle.Baja, ViajeDetalle.Descripcion, ListaPrecio.Nombre AS ListaPrecioNombre, ViajeDetalle.Entregada "
    mrecData.Source = mrecData.Source & "FROM ((((((Viaje INNER JOIN ViajeDetalle ON Viaje.FechaHora = ViajeDetalle.FechaHora AND Viaje.IDRuta = ViajeDetalle.IDRuta) INNER JOIN Persona ON ViajeDetalle.IDPersona = Persona.IDPersona) INNER JOIN Lugar AS Lugar_Origen ON ViajeDetalle.IDOrigen = Lugar_Origen.IDLugar) INNER JOIN Lugar AS Lugar_Destino ON ViajeDetalle.IDDestino = Lugar_Destino.IDLugar) LEFT JOIN Persona AS PersonaRecibe ON ViajeDetalle.IDPersonaRecibe = PersonaRecibe.IDPersona) LEFT JOIN ListaPrecio ON ViajeDetalle.IDListaPrecio = ListaPrecio.IDListaPrecio) LEFT JOIN ViajeDetalle_Comision ON ViajeDetalle.FechaHora = ViajeDetalle_Comision.FechaHora AND ViajeDetalle.IDRuta = ViajeDetalle_Comision.IDRuta AND ViajeDetalle.Indice = ViajeDetalle_Comision.Indice" & SQL_Where
    mrecData.MaxRecords = pParametro.Recordset_MaxRecords
    mrecData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Call SortData(FechaHora, IDRuta, Indice)
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.Comision.FillListView", "Error al leer la lista de Comisiones de la Base de Datos."
End Sub

Private Sub SortData(ByVal FechaHora As Date, ByVal IDRuta As String, ByVal Indice As Long)
    Dim ListItem As MSComctlLib.ListItem
    Dim KeySave As Variant
    Dim CKeySave As Collection
    
    Dim SQL_OrderBy As String

    If Indice = 0 Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
            Set CKeySave = New Collection
            For Each ListItem In lvwData.ListItems
                If ListItem.Selected Then
                    CKeySave.Add ListItem.Key
                End If
            Next ListItem
        End If
    Else
        KeySave = KEY_STRINGER & FechaHora & KEY_DELIMITER & IDRuta & KEY_DELIMITER & Indice
    End If
        
    Select Case lvwData.SortKey
        Case 0  'DIA SEMANA
            SQL_OrderBy = "DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Persona" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 1  'FECHA
            SQL_OrderBy = "FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Persona" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 2  'HORA
            SQL_OrderBy = "Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Fecha" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Persona" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 3  'RUTA
            SQL_OrderBy = "IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Persona" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 4  'REMITENTE
            SQL_OrderBy = "Persona" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 5  'DESTINATARIO
            SQL_OrderBy = "Recibe" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 6  'IMPORTE
            SQL_OrderBy = "Importe" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 7  'PAGADO
            SQL_OrderBy = "ImportePagado" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 8  'DEBE
            SQL_OrderBy = "Debe" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 9 'ORIGEN
            SQL_OrderBy = "Origen" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 10 'DESTINO
            SQL_OrderBy = "Destino" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 11 'DESCRIPCION
            SQL_OrderBy = "Descripcion" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 12 'LISTA PRECIOS
            SQL_OrderBy = "ListaPrecioNombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 13 'ENTREGADA
            SQL_OrderBy = "Entregada" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    lvwData.ListItems.Clear
    With mrecData
        If Not (.BOF And .EOF) Then
            .Sort = SQL_OrderBy
            .MoveFirst
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("FechaHora").Value & KEY_DELIMITER & RTrim(.Fields("IDRuta").Value) & KEY_DELIMITER & .Fields("Indice").Value, WeekdayName(.Fields("DiaSemana").Value))
                ListItem.SubItems(1) = .Fields("Fecha").Value
                ListItem.SubItems(2) = .Fields("Hora").Value
                ListItem.SubItems(3) = RTrim(.Fields("IDRuta").Value)
                ListItem.SubItems(4) = .Fields("Persona").Value
                ListItem.SubItems(5) = .Fields("Recibe").Value & ""
                ListItem.SubItems(6) = Format(.Fields("Importe").Value, "Currency")
                ListItem.SubItems(7) = Format(.Fields("ImportePagado").Value, "Currency")
                ListItem.SubItems(8) = Format(.Fields("Debe").Value, "Currency")
                ListItem.SubItems(9) = IIf(IsNull(.Fields("Sube").Value), .Fields("Origen").Value, .Fields("Sube").Value)
                ListItem.SubItems(10) = IIf(IsNull(.Fields("Baja").Value), .Fields("Destino").Value, .Fields("Baja").Value)
                ListItem.SubItems(11) = .Fields("Descripcion").Value & ""
                ListItem.SubItems(12) = .Fields("ListaPrecioNombre").Value & ""
                ListItem.SubItems(13) = IIf(.Fields("Entregada").Value, "Sí", "No")
                .MoveNext
            Loop
            
            stbMain.Panels("TEXT").Text = .RecordCount & " items" & IIf(.RecordCount = .MaxRecords, " (Limitados)", "")
        Else
            stbMain.Panels("TEXT").Text = "No hay items."
        End If
    End With
    
    On Error Resume Next
    Dim OldSelectedKey As String
    
    OldSelectedKey = lvwData.SelectedItem.Key
    Set lvwData.SelectedItem = lvwData.ListItems(KeySave)
    lvwData.SelectedItem.EnsureVisible
    lvwData.ListItems(OldSelectedKey).Selected = False
    
    If Not CKeySave Is Nothing Then
        If CKeySave.Count > 1 Then
            For Each KeySave In CKeySave
                lvwData.ListItems(KeySave).Selected = True
            Next KeySave
        End If
    End If

    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    Exit Sub

ErrorHandler:
    ShowErrorMessage "Forms.Comision.SortData", "Error al ordenar y mostrar la lista de Comisiones."
End Sub

Public Sub FillComboBoxRuta()
    Dim KeySave As String
    
    KeySave = cboRuta.Text

    cboRuta.Clear
    cboRuta.AddItem "<Todas>"
    Call CSM_Control_ComboBox.FillFromSQL(cboRuta, "SELECT IDRuta FROM Ruta WHERE Activo = 1" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & " ORDER BY IDRuta", "", "IDRuta", "Rutas", cscpCurrentOrFirst, , False)

    cboRuta.ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboRuta, KeySave, cscpItemOrfirst)
End Sub

Public Sub FillComboBoxListaPrecio()
    Dim KeySave As Long
    
    If cboListaPrecio.ListIndex > -1 Then
        KeySave = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
    End If

    cboListaPrecio.Clear
    cboListaPrecio.AddItem "<Todas>"
    Call CSM_Control_ComboBox.FillFromSQL(cboListaPrecio, "SELECT IDListaPrecio, Nombre FROM ListaPrecio WHERE Activo = 1" & IIf(pCPermiso.ListaPrecioWhere <> "", " AND " & Replace(pCPermiso.ListaPrecioWhere, "%TABLENAME%", "ListaPrecio"), "") & " ORDER BY Nombre", "IDListaPrecio", "Nombre", "Listas de Precios", cscpCurrentOrFirst, , False)
    cboListaPrecio.ListIndex = CSM_Control_ComboBox.GetListIndexByItemData(cboListaPrecio, KeySave, cscpItemOrfirst)
End Sub

Private Sub cboFecha_Click()
    txtDiaSemana.Visible = (cboFecha.ListIndex > 0 And cboFecha.ListIndex < 4)
    cmdAnteriorDesde.Visible = (cboFecha.ListIndex > 0)
    dtpFechaDesde.Visible = (cboFecha.ListIndex > 0)
    cmdSiguienteDesde.Visible = (cboFecha.ListIndex > 0)
    cmdHoyDesde.Visible = (cboFecha.ListIndex > 0)
    
    lblFechaAnd.Visible = (cboFecha.ListIndex = 4)
    
    cmdAnteriorHasta.Visible = (cboFecha.ListIndex = 4)
    dtpFechaHasta.Visible = (cboFecha.ListIndex = 4)
    cmdSiguienteHasta.Visible = (cboFecha.ListIndex = 4)
    cmdHoyHasta.Visible = (cboFecha.ListIndex = 4)
    
    cmdAnteriorDesde.Left = 1680
    dtpFechaDesde.Left = 1980
    cmdSiguienteDesde.Left = 3420
    cmdHoyDesde.Left = 3720
    
    If cboFecha.ListIndex > 0 And cboFecha.ListIndex < 4 Then
        cmdAnteriorDesde.Left = cmdAnteriorDesde.Left + txtDiaSemana.Width
        dtpFechaDesde.Left = dtpFechaDesde.Left + txtDiaSemana.Width
        cmdSiguienteDesde.Left = cmdSiguienteDesde.Left + txtDiaSemana.Width
        cmdHoyDesde.Left = cmdHoyDesde.Left + txtDiaSemana.Width
    End If
    
    FillListView Now, "", 0
End Sub

Private Sub cmdAnteriorDesde_Click()
    dtpFechaDesde.Value = DateAdd("d", -1, dtpFechaDesde.Value)
    dtpFechaDesde.SetFocus
    dtpFechaDesde_Change
End Sub

Private Sub dtpFechaDesde_Change()
    txtDiaSemana.Text = WeekdayName(Weekday(dtpFechaDesde.Value))
    FillListView Now, "", 0
End Sub

Private Sub cmdSiguienteDesde_Click()
    dtpFechaDesde.Value = DateAdd("d", 1, dtpFechaDesde.Value)
    dtpFechaDesde.SetFocus
    dtpFechaDesde_Change
End Sub

Private Sub cmdHoyDesde_Click()
    Dim OldValue As Date
    
    OldValue = dtpFechaDesde.Value
    dtpFechaDesde.Value = Date
    dtpFechaDesde.SetFocus
    If OldValue <> dtpFechaDesde.Value Then
        dtpFechaDesde_Change
    End If
End Sub

Private Sub cmdAnteriorHasta_Click()
    dtpFechaHasta.Value = DateAdd("d", -1, dtpFechaHasta.Value)
    dtpFechaHasta.SetFocus
    dtpFechaHasta_Change
End Sub

Private Sub dtpFechaHasta_Change()
    FillListView Now, "", 0
End Sub

Private Sub cmdSiguienteHasta_Click()
    dtpFechaHasta.Value = DateAdd("d", 1, dtpFechaHasta.Value)
    dtpFechaHasta.SetFocus
    dtpFechaHasta_Change
End Sub

Private Sub cmdHoyHasta_Click()
    Dim OldValue As Date
    
    OldValue = dtpFechaHasta.Value
    dtpFechaHasta.Value = Date
    dtpFechaHasta.SetFocus
    If OldValue <> dtpFechaHasta.Value Then
        dtpFechaHasta_Change
    End If
End Sub

Private Sub cboRuta_Click()
    FillListView Now, "", 0
End Sub

Private Sub cmdPersona_Click()
    If pCPermiso.GotPermission(PERMISO_PERSONA) Then
        Call frmPersona.FindAndShowItem(Val(txtPersona.Tag), UCase(Left(txtPersona.Text, 1)), Me.Name, "", "")
    End If
End Sub

Private Sub cmdPersonaClear_Click()
    If Val(txtPersona.Tag) <> 0 Then
        txtPersona.Tag = 0
        txtPersona.Text = ""
        FillListView Now, "", 0
    End If
End Sub

Private Sub cmdPersonaUltimo_Click()
    If frmMDI.cboPersona.ListIndex > -1 Then
        PersonaSelected Val(frmMDI.cboPersona.ItemData(frmMDI.cboPersona.ListIndex)), "PP"
    End If
    cmdPersona.SetFocus
End Sub

Private Sub cboListaPrecio_Click()
    FillListView Now, "", 0
End Sub

Private Sub cboEntregada_Click()
    FillListView Now, "", 0
End Sub

Private Sub cboPagada_Click()
    FillListView Now, "", 0
End Sub

Private Sub cboRendicion_Click()
    txtRendicionDiaSemana.Visible = (cboRendicion.ListIndex > 1 And cboRendicion.ListIndex < 5)
    cmdRendicionAnteriorDesde.Visible = (cboRendicion.ListIndex > 1)
    dtpRendicionFechaDesde.Visible = (cboRendicion.ListIndex > 1)
    cmdRendicionSiguienteDesde.Visible = (cboRendicion.ListIndex > 1)
    cmdRendicionHoyDesde.Visible = (cboRendicion.ListIndex > 1)
    
    lblRendicionFechaAnd.Visible = (cboRendicion.ListIndex = 5)
    
    cmdRendicionAnteriorHasta.Visible = (cboRendicion.ListIndex = 5)
    dtpRendicionFechaHasta.Visible = (cboRendicion.ListIndex = 5)
    cmdRendicionSiguienteHasta.Visible = (cboRendicion.ListIndex = 5)
    cmdRendicionHoyHasta.Visible = (cboRendicion.ListIndex = 5)
    
    cmdRendicionAnteriorDesde.Left = 1920
    dtpRendicionFechaDesde.Left = 2220
    cmdRendicionSiguienteDesde.Left = 3660
    cmdRendicionHoyDesde.Left = 3960
    
    If cboRendicion.ListIndex > 1 And cboRendicion.ListIndex < 5 Then
        cmdRendicionAnteriorDesde.Left = cmdRendicionAnteriorDesde.Left + txtRendicionDiaSemana.Width
        dtpRendicionFechaDesde.Left = dtpRendicionFechaDesde.Left + txtRendicionDiaSemana.Width
        cmdRendicionSiguienteDesde.Left = cmdRendicionSiguienteDesde.Left + txtRendicionDiaSemana.Width
        cmdRendicionHoyDesde.Left = cmdRendicionHoyDesde.Left + txtRendicionDiaSemana.Width
    End If
    
    FillListView Now, "", 0
End Sub

Private Sub dtpRendicionFechaDesde_Change()
    txtRendicionDiaSemana.Text = WeekdayName(Weekday(dtpRendicionFechaDesde.Value))
    FillListView Now, "", 0
End Sub

Private Sub cmdRendicionAnteriorDesde_Click()
    dtpRendicionFechaDesde.Value = DateAdd("d", -1, dtpRendicionFechaDesde.Value)
    dtpRendicionFechaDesde.SetFocus
    dtpRendicionFechaDesde_Change
End Sub

Private Sub cmdRendicionSiguienteDesde_Click()
    dtpRendicionFechaDesde.Value = DateAdd("d", 1, dtpRendicionFechaDesde.Value)
    dtpRendicionFechaDesde.SetFocus
    dtpRendicionFechaDesde_Change
End Sub

Private Sub cmdRendicionHoyDesde_Click()
    Dim OldValue As Date
    
    OldValue = dtpRendicionFechaDesde.Value
    dtpRendicionFechaDesde.Value = Date
    dtpRendicionFechaDesde.SetFocus
    If OldValue <> dtpRendicionFechaDesde.Value Then
        dtpRendicionFechaDesde_Change
    End If
End Sub

Private Sub dtpRendicionFechaHasta_Change()
    FillListView Now, "", 0
End Sub

Private Sub cmdRendicionAnteriorHasta_Click()
    dtpRendicionFechaHasta.Value = DateAdd("d", -1, dtpRendicionFechaHasta.Value)
    dtpRendicionFechaHasta.SetFocus
    dtpRendicionFechaHasta_Change
End Sub

Private Sub cmdRendicionSiguienteHasta_Click()
    dtpRendicionFechaHasta.Value = DateAdd("d", 1, dtpRendicionFechaHasta.Value)
    dtpRendicionFechaHasta.SetFocus
    dtpRendicionFechaHasta_Change
End Sub

Private Sub cmdRendicionHoyHasta_Click()
    Dim OldValue As Date
    
    OldValue = dtpRendicionFechaHasta.Value
    dtpRendicionFechaHasta.Value = Date
    dtpRendicionFechaHasta.SetFocus
    If OldValue <> dtpRendicionFechaHasta.Value Then
        dtpRendicionFechaHasta_Change
    End If
End Sub

Private Sub chkMostrarTodas_Click()
    FillListView Now, "", 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyP
                tlbMain_ButtonClick tlbMain.Buttons.Item("PROPERTIES")
        End Select
    End If
End Sub

Private Sub Form_Load()
    lvwData.GridLines = pParametro.ListView_GridLines
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL TOOLBAR
    Set tlbMain.ImageList = frmMDI.ilsFormToolbar
    Set tlbMain.HotImageList = frmMDI.ilsFormToolbarHot
    tlbMain.Buttons("PROPERTIES").Image = "PROPERTIES"
    tlbMain.Buttons("ASISTENCIA").Image = "ASISTENCIA"
    tlbMain.Buttons("PAGO_MULTIPLE").Image = "PAGO_MULTIPLE"
    tlbMain.Buttons("PRINT").Image = "PRINT"
    tlbMain.Buttons("RENDIR").Image = "RENDIR"
    
    cbrMain.Bands("Toolbar").MinWidth = CSM_Control_Toolbar.GetTotalWidth(tlbMain)
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL LISTVIEW
    Set lvwData.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL PIN
    Set tlbPin.ImageList = frmMDI.ilsFormPin
    '//////////////////////////////////////////////////////////
    
    mLoading = True
    
    cboFecha.AddItem "<Todas>"
    cboFecha.AddItem "="
    cboFecha.AddItem ">="
    cboFecha.AddItem "<="
    cboFecha.AddItem "Entre"
    cboFecha.ListIndex = 0
    
    dtpFechaDesde.Value = Date
    txtDiaSemana.Text = WeekdayName(Weekday(dtpFechaDesde.Value))
    dtpFechaHasta.Value = Date
    
    FillComboBoxRuta
    cboRuta.ListIndex = 0
    
    FillComboBoxListaPrecio
    cboListaPrecio.ListIndex = 0
    
    cboEntregada.AddItem ITEM_ALL_MALE
    cboEntregada.AddItem "Sí"
    cboEntregada.AddItem "No"
    cboEntregada.ListIndex = 2
    
    cboPagada.AddItem ITEM_ALL_MALE
    cboPagada.AddItem "Sí"
    cboPagada.AddItem "No"
    cboPagada.ListIndex = 0
    
    cboRendicion.AddItem "<Todas>"
    cboRendicion.AddItem "Vacía"
    cboRendicion.AddItem "="
    cboRendicion.AddItem ">="
    cboRendicion.AddItem "<="
    cboRendicion.AddItem "Entre"
    cboRendicion.ListIndex = 0
    
    dtpRendicionFechaDesde.Value = Date
    txtRendicionDiaSemana.Text = WeekdayName(Weekday(dtpRendicionFechaDesde.Value))
    dtpRendicionFechaHasta.Value = Date
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "Comision", cbrMain
    pParametro.GetListViewSettings "Comision", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
    tlbPin.Buttons("PIN").Value = pParametro.Usuario_LeerNumero("Comision_Pin", tlbPin.Buttons("PIN").Value)
    If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If

    mLoading = False

    FillListView Now, "", 0
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    ResizeControls NewHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrecData Is Nothing Then
        If mrecData.State = adStateOpen Then
            mrecData.Close
        End If
        Set mrecData = Nothing
    End If
    
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "Comision", cbrMain
    pParametro.SaveListViewSettings "Comision", lvwData
    pParametro.Usuario_GuardarNumero "Comision_Pin", tlbPin.Buttons("PIN").Value
    Set frmComision = Nothing
End Sub

Private Sub lvwData_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    '//////////////////////////////////////////////////////////////////
    'Debido a que no se puede ordenar por Disponibles, ignoro el click en la columna
    If ColumnHeader.Key = "Disponibles" Then
        Exit Sub
    End If
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = 0
    If ColumnHeader.Index - 1 = lvwData.SortKey Then
        lvwData.SortOrder = IIf(lvwData.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwData.SortKey = ColumnHeader.Index - 1
        lvwData.SortOrder = lvwAscending
    End If
    ColumnHeader.Icon = lvwData.SortOrder + 1
    
    'FillListView Date, "", 0
    SortData Date, "", 0
End Sub

Private Sub lvwData_DblClick()
    tlbMain_ButtonClick tlbMain.Buttons.Item("PROPERTIES")
End Sub

Private Sub lvwData_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        lvwData_DblClick
    End If
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim ItemIndex As Long
    Dim SelectedItemCount As Long
    Dim CFechaHora As Collection
    Dim CIDRuta As Collection
    Dim CIndice As Collection
    Dim ImporteTotal As Currency
    
    Dim Viaje As Viaje
    Dim ViajeDetalle As ViajeDetalle
    Dim ViajeDetalle_Comision As ViajeDetalle_Comision
    
    Select Case Button.Key
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                SelectedItemCount = 0
                For ItemIndex = 1 To lvwData.ListItems.Count
                    If lvwData.ListItems(ItemIndex).Selected Then
                        SelectedItemCount = SelectedItemCount + 1
                    End If
                Next ItemIndex
                
                If SelectedItemCount = 1 Then
                    Screen.MousePointer = vbHourglass
                    Set ViajeDetalle = New ViajeDetalle
                    ViajeDetalle.FechaHora = CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                    ViajeDetalle.IDRuta = CStr(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                    ViajeDetalle.Indice = CLng(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER))
                    If ViajeDetalle.Load() Then
                        frmViajeDetallePropiedad.LoadDataAndShow Me, ViajeDetalle
                    Else
                        lvwData.SetFocus
                    End If
                    Set ViajeDetalle = Nothing
                    Screen.MousePointer = vbDefault
                Else
                    MsgBox "No se pueden abrir las propiedades de más de una Comisión a la vez.", vbExclamation, App.Title
                    lvwData.SetFocus
                End If
            End If
        Case "ASISTENCIA"
            If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                SelectedItemCount = 0
                For ItemIndex = 1 To lvwData.ListItems.Count
                    If lvwData.ListItems(ItemIndex).Selected Then
                        SelectedItemCount = SelectedItemCount + 1
                    End If
                Next ItemIndex
                
                If SelectedItemCount = 1 Then
                    Screen.MousePointer = vbHourglass
                    Set ViajeDetalle = New ViajeDetalle
                    ViajeDetalle.FechaHora = CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                    ViajeDetalle.IDRuta = CStr(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                    ViajeDetalle.Indice = CLng(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER))
                    If Not ViajeDetalle.Load() Then
                        lvwData.SetFocus
                        Set ViajeDetalle = Nothing
                        Exit Sub
                    End If
                    
                    Set Viaje = New Viaje
                    Viaje.FechaHora = ViajeDetalle.FechaHora
                    Viaje.IDRuta = ViajeDetalle.IDRuta
                    If Not Viaje.Load() Then
                        lvwData.SetFocus
                        Set Viaje = Nothing
                        Set ViajeDetalle = Nothing
                        Exit Sub
                    End If
                    If Viaje.Estado = VIAJE_ESTADO_FINALIZADO And ViajeDetalle.ImporteContado = ViajeDetalle.Importe And ViajeDetalle.Entregada Then
                        MsgBox "Esta Comisión ya ha sido Pagada y Entregada.", vbInformation, App.Title
                        lvwData.SetFocus
                        Set ViajeDetalle = Nothing
                        Exit Sub
                    End If
                    Set Viaje = Nothing
                    
                    frmViajeDetalleAsistencia.LoadDataAndShow Me, ViajeDetalle
    
                    Set ViajeDetalle = Nothing
                    Screen.MousePointer = vbDefault
                Else
                    MsgBox "No se puede dar Asistencia a más de una Comisión a la vez.", vbExclamation, App.Title
                    lvwData.SetFocus
                End If
            End If
        Case "PAGO_MULTIPLE"
            If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                SelectedItemCount = 0
                Set CFechaHora = New Collection
                Set CIDRuta = New Collection
                Set CIndice = New Collection
                ImporteTotal = 0
                For ItemIndex = 1 To lvwData.ListItems.Count
                    If lvwData.ListItems(ItemIndex).Selected Then
                        SelectedItemCount = SelectedItemCount + 1
                        CFechaHora.Add GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)
                        CIDRuta.Add GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                        CIndice.Add GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)
                        ImporteTotal = ImporteTotal + CCur(lvwData.ListItems(ItemIndex).SubItems(6))
                        If CCur(lvwData.ListItems(ItemIndex).SubItems(6)) <= 0 Then
                            MsgBox "No se pueden seleccionar Comisiones con Importe en Cero.", vbExclamation, App.Title
                            lvwData.SetFocus
                            Exit Sub
                        End If
                        If CCur(lvwData.ListItems(ItemIndex).SubItems(7)) > 0 Then
                            MsgBox "Sólo se pueden seleccionar Comisiones que no tienen Pagos parciales ni totales.", vbExclamation, App.Title
                            lvwData.SetFocus
                            Exit Sub
                        End If
                    End If
                Next ItemIndex
                
                Screen.MousePointer = vbHourglass
                frmComisionAsistenciaMultiple.LoadDataAndShow Me, CFechaHora, CIDRuta, CIndice, ImporteTotal
                Screen.MousePointer = vbDefault
            
                Set CFechaHora = Nothing
                Set CIDRuta = Nothing
                Set CIndice = Nothing
            End If
        Case "RENDIR"
            If pCPermiso.GotPermission(PERMISO_VIAJE_DETALLE_COMISION_RENDIR) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ninguna Comisión seleccionada.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                SelectedItemCount = 0
                Set CFechaHora = New Collection
                Set CIDRuta = New Collection
                Set CIndice = New Collection
                For ItemIndex = 1 To lvwData.ListItems.Count
                    If lvwData.ListItems(ItemIndex).Selected Then
                        SelectedItemCount = SelectedItemCount + 1
                        CFechaHora.Add GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)
                        CIDRuta.Add GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                        CIndice.Add GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)
                    End If
                Next ItemIndex
                
                If SelectedItemCount = 0 Then
                    MsgBox "No hay ninguna Comisión seleccionada.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                If MsgBox("¿Confirma que desea marcar como Rendidas las " & SelectedItemCount & " Comisiones seleccionadas.?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                    For ItemIndex = 1 To CFechaHora.Count
                        Set ViajeDetalle_Comision = New ViajeDetalle_Comision
                        With ViajeDetalle_Comision
                            .RefreshListSkip = True
                            .NoMatchRaiseError = False
                            .FechaHora = CFechaHora(ItemIndex)
                            .IDRuta = CIDRuta(ItemIndex)
                            .Indice = CIndice(ItemIndex)
                            Call .Load
                            .RendicionFechaHora = Now
                            .Update
                        End With
                        Set ViajeDetalle_Comision = Nothing
                    Next ItemIndex
                    RefreshList_RefreshViajeDetalle DATE_TIME_FIELD_NULL_VALUE, "", 0
                End If
            
                Set CFechaHora = Nothing
                Set CIDRuta = Nothing
                Set CIndice = Nothing
            End If
    End Select
End Sub

Private Sub tlbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim ItemIndex As Long
    Dim SelectedItemCount As Long
    Dim Reporte As Reporte
    Dim ReporteSubTitle As String
    
    Select Case ButtonMenu.Parent.Key & "_" & ButtonMenu.Key
        Case "PRINT_LISTADO"
            If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & ButtonMenu.Tag) Then
                Set Reporte = New Reporte
                Reporte.IDReporte = ButtonMenu.Tag
                If Reporte.Load() Then
                    Select Case cboFecha.ListIndex
                        Case 0
                        Case 1
                            Reporte.Parametros("FechaHoraDesde").Valor = CDate(dtpFechaDesde.Value & " 00:00:00")
                            Reporte.Parametros("FechaHoraHasta").Valor = CDate(dtpFechaDesde.Value & " 23:59:59")
                            ReporteSubTitle = "Del día " & dtpFechaDesde.Value
                        Case 2
                            Reporte.Parametros("FechaHoraDesde").Valor = CDate(dtpFechaDesde.Value & " 00:00:00")
                            ReporteSubTitle = "Desde el día " & dtpFechaDesde.Value
                        Case 3
                            Reporte.Parametros("FechaHoraHasta").Valor = CDate(dtpFechaDesde.Value & " 23:59:59")
                            ReporteSubTitle = "Hasta el día " & dtpFechaDesde.Value
                        Case 4
                            Reporte.Parametros("FechaHoraDesde").Valor = CDate(dtpFechaDesde.Value & " 00:00:00")
                            Reporte.Parametros("FechaHoraHasta").Valor = CDate(dtpFechaHasta.Value & " 23:59:59")
                            ReporteSubTitle = dtpFechaDesde.Value & " al " & dtpFechaHasta.Value
                    End Select
                    If cboRuta.ListIndex > 0 Then
                        Reporte.Parametros("IDRuta1").Valor = cboRuta.Text
                    End If
                    If Val(txtPersona.Tag) > 0 Then
                        Reporte.Parametros("IDPersona").Valor = Val(txtPersona.Tag)
                    End If
                    If cboListaPrecio.ListIndex > 0 Then
                        Reporte.Parametros("IDListaPrecio").Valor = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
                    End If
                    If cboEntregada.ListIndex > 0 Then
                        Reporte.Parametros("Entregada").Valor = (cboEntregada.ListIndex = 1)
                    End If
                    If cboPagada.ListIndex > 0 Then
                        Reporte.Parametros("Pagada").Valor = (cboPagada.ListIndex = 1)
                    End If
                    Select Case cboRendicion.ListIndex
                        Case 0
                        Case 1
                            Reporte.Parametros("RendicionVacia").Valor = True
                        Case 2
                            Reporte.Parametros("RendicionFechaHoraDesde").Valor = CDate(dtpRendicionFechaDesde.Value & " 00:00:00")
                            Reporte.Parametros("RendicionFechaHoraHasta").Valor = CDate(dtpRendicionFechaDesde.Value & " 23:59:59")
                            ReporteSubTitle = "Del día " & dtpRendicionFechaDesde.Value
                        Case 3
                            Reporte.Parametros("RendicionFechaHoraDesde").Valor = CDate(dtpRendicionFechaDesde.Value & " 00:00:00")
                            ReporteSubTitle = "Desde el día " & dtpRendicionFechaDesde.Value
                        Case 4
                            Reporte.Parametros("RendicionFechaHoraHasta").Valor = CDate(dtpRendicionFechaDesde.Value & " 23:59:59")
                            ReporteSubTitle = "Hasta el día " & dtpRendicionFechaDesde.Value
                        Case 5
                            Reporte.Parametros("RendicionFechaHoraDesde").Valor = CDate(dtpRendicionFechaDesde.Value & " 00:00:00")
                            Reporte.Parametros("RendicionFechaHoraHasta").Valor = CDate(dtpRendicionFechaHasta.Value & " 23:59:59")
                            ReporteSubTitle = dtpFechaDesde.Value & " al " & dtpFechaHasta.Value
                    End Select
                    Reporte.Parametros("MostrarTodas").Valor = (chkMostrarTodas.Value = vbChecked)
                    Reporte.Titulo = Reporte.Titulo & IIf(ReporteSubTitle = "", "", vbCr) & ReporteSubTitle
                    
                    If Reporte.OpenReport() Then
                        Reporte.PrintReport pParametro.Report_Preview
                    End If
                    
                    Set Reporte = Nothing
                End If
            End If
        Case "PRINT_REMITO"
            If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Comision_Remito") Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If Val(Mid(lvwData.SelectedItem.Key, 2)) < 0 Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                SelectedItemCount = 0
                For ItemIndex = 1 To lvwData.ListItems.Count
                    If lvwData.ListItems(ItemIndex).Selected Then
                        SelectedItemCount = SelectedItemCount + 1
                    End If
                Next ItemIndex
                
                If SelectedItemCount = 1 Then
                    Set Reporte = New Reporte
                    Reporte.IDReporte = "Comision_Remito"
                    If Reporte.Load() Then
                        Reporte.Parametros("FechaHora_FILTER").Valor = CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                        Reporte.Parametros("IDRuta_FILTER").Valor = CStr(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
                        Reporte.Parametros("Indice_FILTER").Valor = CLng(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER))
                        
                        If Reporte.OpenReport() Then
                            Reporte.PrintReport pParametro.Report_Preview
                        End If
                        
                        Set Reporte = Nothing
                    End If
                Else
                    MsgBox "No se pueden imprimir los Remitos de más de una Comisión a la vez.", vbExclamation, App.Title
                    lvwData.SetFocus
                End If
            End If
    End Select
End Sub

Private Sub tlbPin_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Value = tbrUnpressed Then
        Button.Image = 1
    Else
        Button.Image = 2
    End If
End Sub

Private Sub ResizeControls(ByVal CoolBarHeight As Single)
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
    
    lvwData.Top = CoolBarHeight + CONTROL_SPACE
    lvwData.Left = CONTROL_SPACE
    lvwData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lvwData.Height = ScaleHeight - lvwData.Top - CONTROL_SPACE - stbMain.Height
    
    tlbPin.Top = ScaleHeight - 330
    tlbPin.Left = 15
End Sub

Public Sub PersonaSelected(ByVal IDPersona As Long, ByVal Tag As String)
    Dim Persona As Persona
    
    Set Persona = New Persona
    Persona.IDPersona = IDPersona
    If Not Persona.Load() Then
        Set Persona = Nothing
        Exit Sub
    End If
    txtPersona.Tag = IDPersona
    txtPersona.Text = Persona.ApellidoNombre
    Set Persona = Nothing
    
    On Error Resume Next
    lvwData.SetFocus
    On Error GoTo 0
    
    FillListView Now, "", 0
End Sub
