VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCuentaCorrienteTransferencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia de Cuenta Corriente"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CuentaCorrienteTransferencia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   10425
   Begin VB.Frame fraDestino 
      Caption         =   "Origen:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4035
      Left            =   5580
      TabIndex        =   24
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtSaldoActualDestino_Tarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1275
      End
      Begin VB.TextBox txtSaldoFinalDestino_Tarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1275
      End
      Begin VB.CommandButton cmdGrupoDestino 
         Caption         =   "..."
         Height          =   315
         Left            =   4200
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Grupos"
         Top             =   300
         Width           =   255
      End
      Begin VB.TextBox txtSaldoFinalDestino_Efectivo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1275
      End
      Begin VB.TextBox txtSaldoActualDestino_Efectivo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1275
      End
      Begin VB.CommandButton cmdCajaDestino 
         Caption         =   "..."
         Height          =   315
         Left            =   4200
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Cajas"
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSDataListLib.DataCombo datcboCajaDestino 
         Height          =   330
         Left            =   840
         TabIndex        =   29
         Top             =   720
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datcboGrupoDestino 
         Height          =   330
         Left            =   840
         TabIndex        =   26
         Top             =   300
         Width           =   3315
         _ExtentX        =   5847
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
      Begin VB.Line Line9 
         X1              =   4440
         X2              =   4440
         Y1              =   1320
         Y2              =   3840
      End
      Begin VB.Label lblDestino_Tarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjetas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3360
         TabIndex        =   32
         Top             =   1380
         Width           =   720
      End
      Begin VB.Line Line8 
         X1              =   2940
         X2              =   2940
         Y1              =   1320
         Y2              =   3840
      End
      Begin VB.Label lblDestino_Efectivo 
         AutoSize        =   -1  'True
         Caption         =   "Efectivo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1860
         TabIndex        =   31
         Top             =   1380
         Width           =   690
      End
      Begin VB.Line Line7 
         X1              =   1440
         X2              =   4440
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line6 
         Index           =   0
         X1              =   1440
         X2              =   1440
         Y1              =   1320
         Y2              =   3840
      End
      Begin VB.Label lblCajaDestino 
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
         Height          =   210
         Left            =   240
         TabIndex        =   28
         Top             =   780
         Width           =   360
      End
      Begin VB.Label lblGrupoDestino 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   210
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSaldoFinalDestino 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Final:"
         Height          =   210
         Left            =   240
         TabIndex        =   36
         Top             =   3540
         Width           =   825
      End
      Begin VB.Label lblSaldoActualDestino 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Actual:"
         Height          =   210
         Left            =   240
         TabIndex        =   33
         Top             =   1860
         Width           =   960
      End
   End
   Begin VB.Frame fraOrigen 
      Caption         =   "Origen:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin TrueOleDBGrid80.TDBGrid tdbgrdDetalleTarjetas 
         Height          =   1935
         Left            =   120
         TabIndex        =   23
         Top             =   3900
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3413
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "IDMedioPago"
         Columns(0).DataField=   "IDMedioPago"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nombre"
         Columns(1).DataField=   "MedioPagoNombre"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Saldo Actual"
         Columns(2).DataField=   "SaldoActual"
         Columns(2).NumberFormat=   "Currency"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   68
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Transferir"
         Columns(3).DataField=   "Transferir"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   953
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   15790320
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=3519"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3440"
         Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8704"
         Splits(0)._ColumnProps(13)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=2117"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2037"
         Splits(0)._ColumnProps(18)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=8706"
         Splits(0)._ColumnProps(20)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(22)=   "Column(3).Width=1535"
         Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=1455"
         Splits(0)._ColumnProps(25)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   0
         DeadAreaBackColor=   15790320
         RowDividerColor =   15790320
         RowSubDividerColor=   15790320
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Arial"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Arial"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Arial"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton cmdTransferirSaldoActual_Efectivo_Borrar 
         Caption         =   "û"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   21.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2220
         Picture         =   "CuentaCorrienteTransferencia.frx":000C
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Borrar importe a transferir"
         Top             =   2040
         Width           =   480
      End
      Begin VB.CommandButton cmdTransferirSaldoActual_Tarjeta_Borrar 
         Caption         =   "û"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   21.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3720
         Picture         =   "CuentaCorrienteTransferencia.frx":0216
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Borrar importe a transferir"
         Top             =   2040
         Width           =   480
      End
      Begin VB.TextBox txtSaldoFinalOrigen_Tarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1275
      End
      Begin VB.CommandButton cmdTransferirSaldoActual_Tarjeta 
         Caption         =   "ê"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   21.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3240
         Picture         =   "CuentaCorrienteTransferencia.frx":0420
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Transferir el total"
         Top             =   2040
         Width           =   480
      End
      Begin VB.TextBox txtImporteTransferir_Tarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2580
         Width           =   1275
      End
      Begin VB.TextBox txtSaldoActualOrigen_Tarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1275
      End
      Begin VB.CommandButton cmdCajaOrigen 
         Caption         =   "..."
         Height          =   315
         Left            =   4200
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Cajas"
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtSaldoActualOrigen_Efectivo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1275
      End
      Begin VB.TextBox txtImporteTransferir_Efectivo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Top             =   2580
         Width           =   1275
      End
      Begin VB.CommandButton cmdTransferirSaldoActual_Efectivo 
         Caption         =   "ê"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   21.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1740
         Picture         =   "CuentaCorrienteTransferencia.frx":062A
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Transferir el total"
         Top             =   2040
         Width           =   480
      End
      Begin VB.TextBox txtSaldoFinalOrigen_Efectivo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1275
      End
      Begin VB.CommandButton cmdGrupoOrigen 
         Caption         =   "..."
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Grupos"
         Top             =   300
         Width           =   255
      End
      Begin MSDataListLib.DataCombo datcboCajaOrigen 
         Height          =   330
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datcboGrupoOrigen 
         Height          =   330
         Left            =   840
         TabIndex        =   2
         Top             =   300
         Width           =   3315
         _ExtentX        =   5847
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
      Begin VB.Line Line6 
         Index           =   2
         X1              =   1440
         X2              =   4440
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Line Line6 
         Index           =   1
         X1              =   1440
         X2              =   4440
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblDetalleTarjetas 
         AutoSize        =   -1  'True
         Caption         =   "Detalle de Tarjetas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   22
         Top             =   3660
         Width           =   1575
      End
      Begin VB.Line Line5 
         X1              =   4440
         X2              =   4440
         Y1              =   1200
         Y2              =   3420
      End
      Begin VB.Label lblOrigen_Tarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjetas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3360
         TabIndex        =   8
         Top             =   1260
         Width           =   720
      End
      Begin VB.Line Line4 
         X1              =   2940
         X2              =   2940
         Y1              =   1200
         Y2              =   3420
      End
      Begin VB.Label lblOrigen_Efectivo 
         AutoSize        =   -1  'True
         Caption         =   "Efectivo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1860
         TabIndex        =   7
         Top             =   1260
         Width           =   690
      End
      Begin VB.Line Line2 
         X1              =   1440
         X2              =   4440
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         X1              =   1440
         X2              =   1440
         Y1              =   1200
         Y2              =   3420
      End
      Begin VB.Label lblSaldoActualOrigen 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Actual:"
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   1740
         Width           =   960
      End
      Begin VB.Label lblImporteTransferir 
         AutoSize        =   -1  'True
         Caption         =   "&Transferir:"
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   765
      End
      Begin VB.Label lblSaldoFinalOrigen 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Final:"
         Height          =   210
         Left            =   240
         TabIndex        =   19
         Top             =   3060
         Width           =   825
      End
      Begin VB.Label lblGrupoOrigen 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   210
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblCajaOrigen 
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   780
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   7740
      TabIndex        =   39
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   9060
      TabIndex        =   40
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   4980
      Picture         =   "CuentaCorrienteTransferencia.frx":0834
      Top             =   1920
      Width           =   480
   End
End
Attribute VB_Name = "frmCuentaCorrienteTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mKeyDecimal As Boolean

Private mImporteTransferir_Tarjeta As Currency

Private Const GRID_COLUMN_IDMEDIOPAGO As Integer = 0
Private Const GRID_COLUMN_SALDOACTUAL As Integer = 2
Private Const GRID_COLUMN_TRANSFERIR As Integer = 3

Private Sub Form_Load()
    Call CSM_Control_DataCombo.FillFromSQL(datcboGrupoOrigen, "SELECT IDCuentaCorrienteGrupo, Nombre FROM CuentaCorrienteGrupo WHERE Activo = 1 AND IDCuentaCorrienteGrupo <> " & pParametro.CuentaCorrienteGrupo_ID_ViajeDebito & " AND IDCuentaCorrienteGrupo <> " & pParametro.CuentaCorrienteGrupo_ID_ViajeCredito & " ORDER BY Nombre", "IDCuentaCorrienteGrupo", "Nombre", "Grupos", cscpItemOrfirst, pParametro.CuentaCorrienteGrupo_ID_Transferencia)
    datcboGrupoOrigen.Enabled = pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_TRANSFER_GRUPO_CAMBIAR, False)
    Call CSM_Control_DataCombo.FillFromSQL(datcboCajaOrigen, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas")
    
    Call CSM_Control_DataCombo.FillFromSQL(datcboGrupoDestino, "SELECT IDCuentaCorrienteGrupo, Nombre FROM CuentaCorrienteGrupo WHERE Activo = 1 AND IDCuentaCorrienteGrupo <> " & pParametro.CuentaCorrienteGrupo_ID_ViajeDebito & " AND IDCuentaCorrienteGrupo <> " & pParametro.CuentaCorrienteGrupo_ID_ViajeCredito & " ORDER BY Nombre", "IDCuentaCorrienteGrupo", "Nombre", "Grupos", cscpItemOrfirst, pParametro.CuentaCorrienteGrupo_ID_Transferencia)
    datcboGrupoDestino.Enabled = pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_TRANSFER_GRUPO_CAMBIAR, False)
    Call CSM_Control_DataCombo.FillFromSQL(datcboCajaDestino, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas")
    
    txtImporteTransferir_Efectivo.Text = Format(0, "Currency")
    txtImporteTransferir_Tarjeta.Text = Format(0, "Currency")
    
    CSM_Forms.CenterToParent frmMDI, Me
    
    tdbgrdDetalleTarjetas.EvenRowStyle.BackColor = pParametro.GridEvenRowBackColor
    tdbgrdDetalleTarjetas.EvenRowStyle.ForeColor = pParametro.GridEvenRowForeColor
    tdbgrdDetalleTarjetas.OddRowStyle.BackColor = pParametro.GridOddRowBackColor
    tdbgrdDetalleTarjetas.OddRowStyle.ForeColor = pParametro.GridOddRowForeColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCuentaCorrienteTransferencia = Nothing
End Sub

Private Sub cmdGrupoOrigen_Click()
    If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_GRUPO) Then
        Screen.MousePointer = vbHourglass
        frmCuentaCorrienteGrupo.Show
        On Error Resume Next
        Set frmCuentaCorrienteGrupo.lvwData.SelectedItem = frmCuentaCorrienteGrupo.lvwData.ListItems(KEY_STRINGER & Val(datcboGrupoOrigen.BoundText))
        frmCuentaCorrienteGrupo.lvwData.SelectedItem.EnsureVisible
        If frmCuentaCorrienteGrupo.WindowState = vbMinimized Then
            frmCuentaCorrienteGrupo.WindowState = vbNormal
        End If
        frmCuentaCorrienteGrupo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub datcboCajaOrigen_Change()
    Dim CuentaCorrienteCaja As CuentaCorrienteCaja
    
    If Val(datcboCajaOrigen.BoundText) > 0 Then
        Set CuentaCorrienteCaja = New CuentaCorrienteCaja
        CuentaCorrienteCaja.IDCuentaCorrienteCaja = Val(datcboCajaOrigen.BoundText)
        If CuentaCorrienteCaja.Load() Then
            If CuentaCorrienteCaja.OcultarSaldo And pUsuario.IDCuentaCorrienteCaja <> Val(datcboCajaOrigen.BoundText) And (Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_SALDOOCULTO_VIEW, False)) Then
                txtSaldoActualOrigen_Efectivo.Text = Format(0, "Currency")
                txtSaldoActualOrigen_Tarjeta.Text = Format(0, "Currency")
            Else
                If CuentaCorrienteCaja.LoadSaldoActual() Then
                    txtSaldoActualOrigen_Efectivo.Text = CuentaCorrienteCaja.SaldoActual_Efectivo_Formatted
                    txtSaldoActualOrigen_Tarjeta.Text = CuentaCorrienteCaja.SaldoActual_Tarjeta_Formatted
                    Call FillGrid
                End If
            End If
            cmdTransferirSaldoActual_Efectivo_Borrar_Click
            'cmdTransferirSaldoActual_Tarjeta_Borrar_Click
        End If
        Set CuentaCorrienteCaja = Nothing
        
        CalcularSaldoFinal_Efectivo
        CalcularSaldoFinal_Tarjeta
    End If
End Sub

Private Sub cmdGrupoDestino_Click()
    If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_GRUPO) Then
        Screen.MousePointer = vbHourglass
        frmCuentaCorrienteGrupo.Show
        On Error Resume Next
        Set frmCuentaCorrienteGrupo.lvwData.SelectedItem = frmCuentaCorrienteGrupo.lvwData.ListItems(KEY_STRINGER & Val(datcboGrupoDestino.BoundText))
        frmCuentaCorrienteGrupo.lvwData.SelectedItem.EnsureVisible
        If frmCuentaCorrienteGrupo.WindowState = vbMinimized Then
            frmCuentaCorrienteGrupo.WindowState = vbNormal
        End If
        frmCuentaCorrienteGrupo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub datcboCajaDestino_Change()
    Dim CuentaCorrienteCaja As CuentaCorrienteCaja
    
    If Val(datcboCajaDestino.BoundText) > 0 Then
        Set CuentaCorrienteCaja = New CuentaCorrienteCaja
        CuentaCorrienteCaja.IDCuentaCorrienteCaja = Val(datcboCajaDestino.BoundText)
        If CuentaCorrienteCaja.Load() Then
            If CuentaCorrienteCaja.OcultarSaldo And pUsuario.IDCuentaCorrienteCaja <> Val(datcboCajaOrigen.BoundText) And (Not pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_SALDOOCULTO_VIEW, False)) Then
                txtSaldoActualDestino_Efectivo.Text = Format(0, "Currency")
                txtSaldoActualDestino_Tarjeta.Text = Format(0, "Currency")
            Else
                If CuentaCorrienteCaja.LoadSaldoActual() Then
                    txtSaldoActualDestino_Efectivo.Text = CuentaCorrienteCaja.SaldoActual_Efectivo_Formatted
                    txtSaldoActualDestino_Tarjeta.Text = CuentaCorrienteCaja.SaldoActual_Tarjeta_Formatted
                End If
            End If
        End If
        Set CuentaCorrienteCaja = Nothing
        
        CalcularSaldoFinal_Efectivo
        CalcularSaldoFinal_Tarjeta
    End If
End Sub

Private Sub cmdTransferirSaldoActual_Efectivo_Click()
    txtImporteTransferir_Efectivo.Text = txtSaldoActualOrigen_Efectivo.Text
End Sub

Private Sub cmdTransferirSaldoActual_Efectivo_Borrar_Click()
    txtImporteTransferir_Efectivo.Text = Format(0, "Currency")
End Sub

Private Sub cmdTransferirSaldoActual_Tarjeta_Click()
    'txtImporteTransferir_Tarjeta.Text = txtSaldoActualOrigen_Tarjeta.Text
    mImporteTransferir_Tarjeta = 0
    Call MarkGrid(True)
End Sub

Private Sub cmdTransferirSaldoActual_Tarjeta_Borrar_Click()
    'txtImporteTransferir_Tarjeta.Text = Format(0, "Currency")
    Call MarkGrid(False)
    mImporteTransferir_Tarjeta = 0
End Sub

Private Sub txtImporteTransferir_Efectivo_GotFocus()
    CSM_Control_TextBox.SelAllText txtImporteTransferir_Efectivo
End Sub

Private Sub txtImporteTransferir_Efectivo_Change()
    CalcularSaldoFinal_Efectivo
End Sub

Private Sub txtImporteTransferir_Efectivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtImporteTransferir_Efectivo_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, txtImporteTransferir_Efectivo.Text, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, txtImporteTransferir_Efectivo.Text, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtImporteTransferir_Efectivo_LostFocus()
    If Not IsNumeric(txtImporteTransferir_Efectivo.Text) Then
        txtImporteTransferir_Efectivo.Text = Val(txtImporteTransferir_Efectivo.Text)
    End If
    txtImporteTransferir_Efectivo.Text = Format(CCur(txtImporteTransferir_Efectivo.Text), "Currency")
End Sub

Private Sub txtImporteTransferir_Tarjeta_GotFocus()
    CSM_Control_TextBox.SelAllText txtImporteTransferir_Tarjeta
End Sub

Private Sub txtImporteTransferir_Tarjeta_Change()
    CalcularSaldoFinal_Tarjeta
End Sub

Private Sub txtImporteTransferir_Tarjeta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtImporteTransferir_Tarjeta_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, txtImporteTransferir_Tarjeta.Text, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, txtImporteTransferir_Tarjeta.Text, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtImporteTransferir_Tarjeta_LostFocus()
    If Not IsNumeric(txtImporteTransferir_Tarjeta.Text) Then
        txtImporteTransferir_Tarjeta.Text = Val(txtImporteTransferir_Tarjeta.Text)
    End If
    txtImporteTransferir_Tarjeta.Text = Format(CCur(txtImporteTransferir_Tarjeta.Text), "Currency")
End Sub

Private Sub tdbgrdDetalleTarjetas_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = GRID_COLUMN_TRANSFERIR Then
        If CBool(tdbgrdDetalleTarjetas.Columns(GRID_COLUMN_TRANSFERIR).Value) Then
            mImporteTransferir_Tarjeta = mImporteTransferir_Tarjeta + CCur(tdbgrdDetalleTarjetas.Columns(GRID_COLUMN_SALDOACTUAL).Value)
        Else
            mImporteTransferir_Tarjeta = mImporteTransferir_Tarjeta - CCur(tdbgrdDetalleTarjetas.Columns(GRID_COLUMN_SALDOACTUAL).Value)
        End If
        txtImporteTransferir_Tarjeta.Text = Format(mImporteTransferir_Tarjeta, "Currency")
    End If
End Sub

Private Sub cmdOK_Click()
    Dim CuentaCorriente As CuentaCorriente
    Dim Index As Integer
    'GRUPO
    If Val(datcboGrupoOrigen.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo de Origen.", vbInformation, App.Title
        datcboGrupoOrigen.SetFocus
        Exit Sub
    End If
    If Val(datcboGrupoDestino.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo de Destino.", vbInformation, App.Title
        datcboGrupoDestino.SetFocus
        Exit Sub
    End If
    
    'CAJA
    If Val(datcboCajaOrigen.BoundText) = 0 Then
        MsgBox "Debe seleccionar la Caja de Origen.", vbInformation, App.Title
        datcboCajaOrigen.SetFocus
        Exit Sub
    End If
    If Val(datcboCajaDestino.BoundText) = 0 Then
        MsgBox "Debe seleccionar la Caja de Destino.", vbInformation, App.Title
        datcboCajaDestino.SetFocus
        Exit Sub
    End If
    If Val(datcboCajaOrigen.BoundText) = Val(datcboCajaDestino.BoundText) Then
        MsgBox "La Caja de Destino debe ser distinta a la Caja de Origen.", vbInformation, App.Title
        datcboCajaDestino.SetFocus
        Exit Sub
    End If
    
    'EFECTIVO
    If Not IsNumeric(txtImporteTransferir_Efectivo.Text) Then
        MsgBox "El Importe en Efectivo debe ser un valor numérico.", vbInformation, App.Title
        txtImporteTransferir_Efectivo.SetFocus
        Exit Sub
    End If
    If CCur(txtSaldoActualOrigen_Efectivo.Text) < 0 And CCur(txtImporteTransferir_Efectivo.Text) <> 0 Then
        MsgBox "No se puede transferir el Efectivo porque tiene saldo negativo.", vbInformation, App.Title
        txtImporteTransferir_Efectivo.SetFocus
        Exit Sub
    End If
    If (CCur(txtImporteTransferir_Efectivo.Text) > CCur(txtSaldoActualOrigen_Efectivo.Text)) And CCur(txtImporteTransferir_Efectivo.Text) <> 0 Then
        MsgBox "El Importe en Efectivo a transferir debe ser menor o igual al Saldo de la Caja de Origen.", vbInformation, App.Title
        txtImporteTransferir_Efectivo.SetFocus
        Exit Sub
    End If
    
    'TARJETA
    If Not IsNumeric(txtImporteTransferir_Tarjeta.Text) Then
        MsgBox "El Importe de las Tarjetas debe ser un valor numérico.", vbInformation, App.Title
        txtImporteTransferir_Tarjeta.SetFocus
        Exit Sub
    End If
    If CCur(txtSaldoActualOrigen_Tarjeta.Text) < 0 And CCur(txtImporteTransferir_Tarjeta.Text) <> 0 Then
        MsgBox "No se pueden transferir las Tarejetas porque tienen saldo negativo.", vbInformation, App.Title
        txtImporteTransferir_Tarjeta.SetFocus
        Exit Sub
    End If
    If CCur(txtImporteTransferir_Tarjeta.Text) > CCur(txtSaldoActualOrigen_Tarjeta.Text) And CCur(txtImporteTransferir_Tarjeta.Text) <> 0 Then
        MsgBox "El Importe de Tarjetas a transferir debe ser menor o igual al Saldo de la Caja de Origen.", vbInformation, App.Title
        txtImporteTransferir_Tarjeta.SetFocus
        Exit Sub
    End If
    
    If MsgBox("¿Confirma la Transferencia de Cajas?" & vbCr & vbCr & "Caja Origen: " & datcboCajaOrigen.Text & vbCr & vbCr & "Caja Destino: " & datcboCajaDestino.Text & vbCr & vbCr & "Importe en Efectivo: " & txtImporteTransferir_Efectivo.Text & vbCr & vbCr & "Importe de Tarjetas: " & txtImporteTransferir_Tarjeta.Text, vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    
    cmdOK.Enabled = False
    
    'EFECTIVO
    If CCur(txtImporteTransferir_Efectivo.Text) <> 0 Then
        Set CuentaCorriente = New CuentaCorriente
        With CuentaCorriente
            'ORIGEN
            .IDCuentaCorrienteGrupo = Val(datcboGrupoOrigen.BoundText)
            .IDCuentaCorrienteCaja = Val(datcboCajaOrigen.BoundText)
            .Descripcion = "Transferencia a: " & datcboCajaDestino.Text
            
            'DESTINO
            .IDCuentaCorrienteGrupo_Destino = Val(datcboGrupoDestino.BoundText)
            .IDCuentaCorrienteCaja_Destino = Val(datcboCajaDestino.BoundText)
            .Descripcion_Destino = "Transferencia desde: " & datcboCajaOrigen.Text
    
            'COMUN
            .FechaHora = Now
            .Importe = CCur(txtImporteTransferir_Efectivo.Text)
            .IDMedioPago = pParametro.MedioPago_Predeterminado_ID
            
            If Not .Transferir() Then
                Set CuentaCorriente = New CuentaCorriente
                cmdOK.Enabled = True
                Exit Sub
            End If
        End With
        Set CuentaCorriente = Nothing
    End If
    
    'TARJETAS
    tdbgrdDetalleTarjetas.Visible = False
    tdbgrdDetalleTarjetas.MoveFirst
    
    For Index = 0 To tdbgrdDetalleTarjetas.ApproxCount - 1
        If CBool(tdbgrdDetalleTarjetas.Columns(GRID_COLUMN_TRANSFERIR).Value) Then
            If CCur(tdbgrdDetalleTarjetas.Columns(GRID_COLUMN_SALDOACTUAL).Value) <> 0 Then
                Set CuentaCorriente = New CuentaCorriente
                With CuentaCorriente
                    'ORIGEN
                    .IDCuentaCorrienteGrupo = Val(datcboGrupoOrigen.BoundText)
                    .IDCuentaCorrienteCaja = Val(datcboCajaOrigen.BoundText)
                    .Descripcion = "Transferencia a: " & datcboCajaDestino.Text
                    
                    'DESTINO
                    .IDCuentaCorrienteGrupo_Destino = Val(datcboGrupoDestino.BoundText)
                    .IDCuentaCorrienteCaja_Destino = Val(datcboCajaDestino.BoundText)
                    .Descripcion_Destino = "Transferencia desde: " & datcboCajaOrigen.Text
            
                    'COMUN
                    .FechaHora = Now
                    .Importe = CCur(tdbgrdDetalleTarjetas.Columns(GRID_COLUMN_SALDOACTUAL).Value)
                    
                    .IDMedioPago = Val(tdbgrdDetalleTarjetas.Columns(GRID_COLUMN_IDMEDIOPAGO).Value)
                    
                    If Not .Transferir() Then
                        Set CuentaCorriente = New CuentaCorriente
                        cmdOK.Enabled = True
                        Exit Sub
                    End If
                End With
                Set CuentaCorriente = Nothing
            End If
        End If
        
        tdbgrdDetalleTarjetas.MoveNext
    Next Index
    
    tdbgrdDetalleTarjetas.MoveFirst
    tdbgrdDetalleTarjetas.Visible = True
    
        
    cmdOK.Enabled = True
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CalcularSaldoFinal_Efectivo()
    If IsNumeric(txtSaldoActualOrigen_Efectivo.Text) Then
        If IsNumeric(txtImporteTransferir_Efectivo.Text) Then
            txtSaldoFinalOrigen_Efectivo.Text = Format(CCur(txtSaldoActualOrigen_Efectivo.Text) - CCur(txtImporteTransferir_Efectivo.Text), "Currency")
        Else
            txtSaldoFinalOrigen_Efectivo.Text = Format(CCur(txtSaldoActualOrigen_Efectivo.Text), "Currency")
        End If
    End If
    
    If IsNumeric(txtSaldoActualDestino_Efectivo.Text) Then
        If IsNumeric(txtImporteTransferir_Efectivo.Text) Then
            txtSaldoFinalDestino_Efectivo.Text = Format(CCur(txtSaldoActualDestino_Efectivo.Text) + CCur(txtImporteTransferir_Efectivo.Text), "Currency")
        Else
            txtSaldoFinalDestino_Efectivo.Text = Format(CCur(txtSaldoActualDestino_Efectivo.Text), "Currency")
        End If
    End If
End Sub

Private Sub CalcularSaldoFinal_Tarjeta()
    If IsNumeric(txtSaldoActualOrigen_Tarjeta.Text) Then
        If IsNumeric(txtImporteTransferir_Tarjeta.Text) Then
            txtSaldoFinalOrigen_Tarjeta.Text = Format(CCur(txtSaldoActualOrigen_Tarjeta.Text) - CCur(txtImporteTransferir_Tarjeta.Text), "Currency")
        Else
            txtSaldoFinalOrigen_Tarjeta.Text = Format(CCur(txtSaldoActualOrigen_Tarjeta.Text), "Currency")
        End If
    End If
    
    If IsNumeric(txtSaldoActualDestino_Tarjeta.Text) Then
        If IsNumeric(txtImporteTransferir_Tarjeta.Text) Then
            txtSaldoFinalDestino_Tarjeta.Text = Format(CCur(txtSaldoActualDestino_Tarjeta.Text) + CCur(txtImporteTransferir_Tarjeta.Text), "Currency")
        Else
            txtSaldoFinalDestino_Tarjeta.Text = Format(CCur(txtSaldoActualDestino_Tarjeta.Text), "Currency")
        End If
    End If
End Sub

Public Sub FillComboBoxCuentaCorrienteGrupo()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboGrupoOrigen.BoundText)
    Set recData = datcboGrupoOrigen.RowSource
    recData.Requery
    Set recData = Nothing
    datcboGrupoOrigen.BoundText = KeySave

    KeySave = Val(datcboGrupoDestino.BoundText)
    Set recData = datcboGrupoDestino.RowSource
    recData.Requery
    Set recData = Nothing
    datcboGrupoDestino.BoundText = KeySave
End Sub

Public Sub FillComboBoxCuentaCorrienteCaja()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboCajaOrigen.BoundText)
    Set recData = datcboCajaOrigen.RowSource
    recData.Requery
    Set recData = Nothing
    datcboCajaOrigen.BoundText = KeySave

    KeySave = Val(datcboCajaDestino.BoundText)
    Set recData = datcboCajaDestino.RowSource
    recData.Requery
    Set recData = Nothing
    datcboCajaDestino.BoundText = KeySave
End Sub

Private Sub FillGrid()
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    Dim XArrayDB As XArrayDBObject.XArrayDB
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass
    
    Set XArrayDB = New XArrayDBObject.XArrayDB
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandType = adCmdStoredProc
    cmdData.CommandText = "sp_CuentaCorrienteCaja_SaldoActual_Tarjeta"
    cmdData.Parameters.Append cmdData.CreateParameter("IDCuentaCorrienteCaja", adInteger, adParamInput, , Val(datcboCajaOrigen.BoundText))
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    Set cmdData = Nothing
    
    If Not recData.EOF Then
        Call XArrayDB.LoadRows(recData.GetRows())
    End If
    Set tdbgrdDetalleTarjetas.Array = XArrayDB
    tdbgrdDetalleTarjetas.ReBind
    
    mImporteTransferir_Tarjeta = 0
    txtImporteTransferir_Tarjeta.Text = Format(mImporteTransferir_Tarjeta, "Currency")
    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorHandler:
    ShowErrorMessage "Forms.CuentaCorrienteTransferencia.FillGrid", "Error al obtener el Saldo de las Tarjetas de la Caja de Cuenta Corriente." & vbCr & vbCr & "IDCuentaCorrienteCaja: " & Val(datcboCajaOrigen.BoundText)
End Sub

Private Sub MarkGrid(ByVal Value As Boolean)
    Dim Index As Integer
    
    tdbgrdDetalleTarjetas.Visible = False
    tdbgrdDetalleTarjetas.MoveFirst
    For Index = 0 To tdbgrdDetalleTarjetas.ApproxCount - 1
        tdbgrdDetalleTarjetas.Columns(GRID_COLUMN_TRANSFERIR).Value = Value
        Call tdbgrdDetalleTarjetas_AfterColUpdate(GRID_COLUMN_TRANSFERIR)
        tdbgrdDetalleTarjetas.MoveNext
    Next Index
    tdbgrdDetalleTarjetas.MoveFirst
    tdbgrdDetalleTarjetas.Visible = True
End Sub
