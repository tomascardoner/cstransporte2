VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmVehiculoUtilizacion 
   Caption         =   "Utilización de Vehículos"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11490
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VehiculoUtilizacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   11490
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   14
      Top             =   5025
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   767
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   159
            MinWidth        =   18
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   741
      FixedOrder      =   -1  'True
      _CBWidth        =   11490
      _CBHeight       =   420
      _Version        =   "6.7.9782"
      Caption1        =   "Fecha:"
      Child1          =   "picFecha"
      MinWidth1       =   3465
      MinHeight1      =   360
      Width1          =   3465
      Key1            =   "Fecha"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "Intervalo:"
      Child2          =   "picIntervalo"
      MinWidth2       =   2025
      MinHeight2      =   315
      Width2          =   2025
      Key2            =   "Intervalo"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Caption3        =   "Rango:"
      Child3          =   "picRango"
      MinWidth3       =   2055
      MinHeight3      =   315
      Width3          =   2865
      Key3            =   "Rango"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Begin VB.PictureBox picRango 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   9345
         ScaleHeight     =   315
         ScaleWidth      =   2055
         TabIndex        =   10
         Top             =   45
         Width           =   2055
         Begin VB.ComboBox cboDiaHoraInicio 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox cboDiaHoraFin 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblDiaHoraA 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   210
            Left            =   960
            TabIndex        =   13
            Top             =   90
            Width           =   90
         End
      End
      Begin VB.PictureBox picIntervalo 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   6480
         ScaleHeight     =   315
         ScaleWidth      =   2025
         TabIndex        =   8
         Top             =   45
         Width           =   2025
         Begin VB.ComboBox cboIntervalo 
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
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "Intervalo de Tiempo por Columna"
            Top             =   0
            Width           =   2040
         End
      End
      Begin VB.PictureBox picFecha 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   660
         ScaleHeight     =   360
         ScaleWidth      =   4800
         TabIndex        =   2
         Top             =   30
         Width           =   4800
         Begin VB.CommandButton cmdFechaAnterior 
            Height          =   315
            Left            =   1080
            Picture         =   "VehiculoUtilizacion.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton cmdFechaSiguiente 
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
            Left            =   2820
            Picture         =   "VehiculoUtilizacion.frx":0E54
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton cmdFechaHoy 
            Height          =   315
            Left            =   3120
            Picture         =   "VehiculoUtilizacion.frx":13DE
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Width           =   315
         End
         Begin VB.TextBox txtFechaDiaSemana 
            Height          =   315
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   1050
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   315
            Left            =   1380
            TabIndex        =   7
            Top             =   0
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
            Format          =   102891521
            CurrentDate     =   36950
         End
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbgrdData 
      Height          =   2535
      Left            =   360
      TabIndex        =   0
      Top             =   660
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4471
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "IDVehiculo"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Vehículo"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).PartialRightColumn=   0   'False
      Splits(0).Locked=   -1  'True
      Splits(0).AllowSizing=   -1  'True
      Splits(0).SizeMode=   1
      Splits(0).Size  =   3000,189
      Splits(0).Size.vt=   4
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   953
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8212"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=344"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=265"
      Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=8720"
      Splits(0)._ColumnProps(14)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
      Splits(1)._UserFlags=   0
      Splits(1).Locked=   -1  'True
      Splits(1).AllowSizing=   -1  'True
      Splits(1).AllowRowSizing=   0   'False
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   953
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).AllowColSelect=   0   'False
      Splits(1).AllowRowSelect=   0   'False
      Splits(1).DividerColor=   15790320
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=2"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=8212"
      Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(9)=   "Column(1).Width=2725"
      Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=2646"
      Splits(1)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=8212"
      Splits(1)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(16)=   "Column(1).Order=2"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   0
      DeadAreaBackColor=   15790320
      RowDividerColor =   6579300
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
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.valignment=2,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0,.locked=-1,.bold=-1"
      _StyleDefs(41)  =   ":id=32,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(42)  =   ":id=32,.fontname=Arial"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(1).Style:id=43,.parent=1,.bgcolor=&H808080&"
      _StyleDefs(47)  =   "Splits(1).CaptionStyle:id=52,.parent=4"
      _StyleDefs(48)  =   "Splits(1).HeadingStyle:id=44,.parent=2"
      _StyleDefs(49)  =   "Splits(1).FooterStyle:id=45,.parent=3"
      _StyleDefs(50)  =   "Splits(1).InactiveStyle:id=46,.parent=5"
      _StyleDefs(51)  =   "Splits(1).SelectedStyle:id=48,.parent=6"
      _StyleDefs(52)  =   "Splits(1).EditorStyle:id=47,.parent=7"
      _StyleDefs(53)  =   "Splits(1).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(54)  =   "Splits(1).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(55)  =   "Splits(1).OddRowStyle:id=51,.parent=10"
      _StyleDefs(56)  =   "Splits(1).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(57)  =   "Splits(1).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(58)  =   "Splits(1).Columns(0).Style:id=58,.parent=43,.locked=-1"
      _StyleDefs(59)  =   "Splits(1).Columns(0).HeadingStyle:id=55,.parent=44"
      _StyleDefs(60)  =   "Splits(1).Columns(0).FooterStyle:id=56,.parent=45"
      _StyleDefs(61)  =   "Splits(1).Columns(0).EditorStyle:id=57,.parent=47"
      _StyleDefs(62)  =   "Splits(1).Columns(1).Style:id=62,.parent=43,.locked=-1"
      _StyleDefs(63)  =   "Splits(1).Columns(1).HeadingStyle:id=59,.parent=44"
      _StyleDefs(64)  =   "Splits(1).Columns(1).FooterStyle:id=60,.parent=45"
      _StyleDefs(65)  =   "Splits(1).Columns(1).EditorStyle:id=61,.parent=47"
      _StyleDefs(66)  =   "Named:id=33:Normal"
      _StyleDefs(67)  =   ":id=33,.parent=0"
      _StyleDefs(68)  =   "Named:id=34:Heading"
      _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   ":id=34,.wraptext=-1"
      _StyleDefs(71)  =   "Named:id=35:Footing"
      _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   "Named:id=36:Selected"
      _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=37:Caption"
      _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(77)  =   "Named:id=38:HighlightRow"
      _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(79)  =   "Named:id=39:EvenRow"
      _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(81)  =   "Named:id=40:OddRow"
      _StyleDefs(82)  =   ":id=40,.parent=33"
      _StyleDefs(83)  =   "Named:id=41:RecordSelector"
      _StyleDefs(84)  =   ":id=41,.parent=34"
      _StyleDefs(85)  =   "Named:id=42:FilterBar"
      _StyleDefs(86)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmVehiculoUtilizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Estas variables son para ver la posicion del mouse durante un drag and drop
Private m_iOverCol As Integer
Private m_iOverRow As Integer
Private m_LastFirstRow As Integer

Private m_recTableroComando As New ADODB.Recordset

Private m_RefreshGrid As Boolean
Private m_IntervalMinutes As Long
Private m_DayTimeStartMinutes As Integer
Private m_DayTimeEndMinutes As Integer
Private m_ColumnHourLast As Long
Private m_ComodidadNombreBeingDrag As String
Private m_InicioBeingDrag As Date
Private m_FinBeingDrag As Date

Private mxaEstado As XArrayDBObject.XArrayDB
Private mxaFechaHora As XArrayDBObject.XArrayDB
Private mxaIDRuta As XArrayDBObject.XArrayDB

Private mCViajesEnCeldaActual As Collection

Private mVehiculoUtilizacion_ViajeEnProgreso_FinalizaEnHoraActual As Boolean

Const COLUMN_INDEX_IDVEHICULO = 0
Const COLUMN_INDEX_VEHICULONOMBRE = 1
Const COLUMN_HOUR_FIRST = 2

Const COLUMN_DIVIDER_STYLE = dbgBlackLine

Const RESERVA_INICIO As String = "<"
Const RESERVA_FIN As String = ">"
Const RESERVA_CONFLICTO As String = "!"

Public Sub LoadDataAndShow()
    m_RefreshGrid = False
    
    Load frmVehiculoUtilizacion
    
    m_RefreshGrid = True
    
    FormatGrid
    
    dtpFecha_Change
    
    frmVehiculoUtilizacion.Show
End Sub

'=================================================================
'Pinta la grilla poniendo las letras correspondientes en cada celda.
'Además, agrega el IDComodidadUtilizada del row actual en una colección en una segunda (tercer dimensión) del XArray
Private Sub PaintGrid(ByVal FechaHora As Date, ByVal IDRuta As String, ByVal Estado As String, ByVal Inicio As Date, ByVal Fin As Date, ByVal RowIndex As Long)
    Dim ColumnaIndex As Integer
    Dim ColumnaInicio As Long
    Dim ColumnaFin As Long
    
    Dim CFechaHora As Collection
    Dim CIDRuta As Collection
    
    'Calculo las Columnas de Inicio y Fin
    GetStartAndEndColumns Estado, Inicio, Fin, ColumnaInicio, ColumnaFin
    
    If ColumnaInicio = -1 Or ColumnaFin = -1 Then
        Exit Sub
    End If
    
    'Recorro tantas columnas como tengo que pintar según la diferencia entre iInicio y iFin
    For ColumnaIndex = ColumnaInicio To ColumnaFin
        'Asigno los IDs de las reservas
        If IsEmpty(mxaFechaHora(RowIndex, ColumnaIndex)) Then
            Set CFechaHora = New Collection
            Set CIDRuta = New Collection
            mxaFechaHora(RowIndex, ColumnaIndex) = CFechaHora
            mxaIDRuta(RowIndex, ColumnaIndex) = CIDRuta
        Else
            Set CFechaHora = mxaFechaHora(RowIndex, ColumnaIndex)
            Set CIDRuta = mxaIDRuta(RowIndex, ColumnaIndex)
        End If
        CFechaHora.Add FechaHora
        CIDRuta.Add IDRuta
        
        If IsEmpty(mxaEstado(RowIndex, ColumnaIndex)) Then
            If (ColumnaInicio = ColumnaIndex) And (ColumnaIndex > COLUMN_HOUR_FIRST) Then
                mxaEstado(RowIndex, ColumnaIndex) = Estado & RESERVA_INICIO
            Else
                If (ColumnaFin = ColumnaIndex) And (Estado <> VIAJE_ESTADO_EN_PROGRESO) And (ColumnaFin < m_ColumnHourLast) Then
                    mxaEstado(RowIndex, ColumnaIndex) = Estado & RESERVA_FIN
                Else
                    mxaEstado(RowIndex, ColumnaIndex) = Estado
                End If
            End If
        ElseIf InStr(1, mxaEstado(RowIndex, ColumnaIndex), KEY_DELIMITER) = 0 And mxaEstado(RowIndex, ColumnaIndex) <> RESERVA_CONFLICTO Then
            'Está asignado un sólo viaje, le sumo el estado del viaje actual
            mxaEstado(RowIndex, ColumnaIndex) = Left(mxaEstado(RowIndex, ColumnaIndex), 2) & KEY_DELIMITER & Estado
        Else
            'Hay conflicto
            mxaEstado(RowIndex, ColumnaIndex) = RESERVA_CONFLICTO
        End If
    Next ColumnaIndex
End Sub

'Esta Función me devuelve los Indices de las Columnas de Inicio y de Fin Para una Fecha de Inicio y de Fin
Private Sub GetStartAndEndColumns(ByVal Estado As String, ByVal FechaInicio As Date, ByVal FechaFin As Date, ByRef ColumnaInicio As Long, ByRef ColumnaFin As Long)
    If DateDiff("n", dtpFecha.Value, FechaInicio) < 0 Then
        ColumnaInicio = COLUMN_HOUR_FIRST
    Else
        ColumnaInicio = TimeToColIndex(FechaInicio)
    End If
    
    If Estado = VIAJE_ESTADO_EN_PROGRESO And mVehiculoUtilizacion_ViajeEnProgreso_FinalizaEnHoraActual Then
        If IIf(Format(FechaFin, "yyyy/mm/dd") < Format(Now, "yyyy/mm/dd"), Format(Now, "yyyy/mm/dd"), Format(FechaFin, "yyyy/mm/dd")) > Format(dtpFecha.Value, "yyyy/mm/dd") Then
            ColumnaFin = m_ColumnHourLast
        Else
            ColumnaFin = TimeToColIndex(IIf(Format(FechaFin, "yyyy/mm/dd hh:mm") < Format(Now, "yyyy/mm/dd hh:mm"), Format(Now, "yyyy/mm/dd hh:mm"), Format(FechaFin, "yyyy/mm/dd hh:mm")))
        End If
    Else
        If DateDiff("d", dtpFecha.Value, FechaFin) > 0 Then
            ColumnaFin = m_ColumnHourLast
        Else
            ColumnaFin = TimeToColIndex(FechaFin)
        End If
    End If
    If ColumnaInicio < COLUMN_HOUR_FIRST Then
        If ColumnaFin < COLUMN_HOUR_FIRST Then
            ColumnaInicio = -1
            ColumnaFin = -1
        Else
            ColumnaInicio = COLUMN_HOUR_FIRST
        End If
    End If
    If ColumnaFin > m_ColumnHourLast Then
        If ColumnaInicio > m_ColumnHourLast Then
            ColumnaInicio = -1
            ColumnaFin = -1
        Else
            ColumnaFin = m_ColumnHourLast
        End If
    End If
    If ColumnaInicio > ColumnaFin Then
        ColumnaInicio = -1
        ColumnaFin = -1
    End If
End Sub

'Esta Funcion me devuelve el índice de la columna correspondiente a la hora que le paso
'teniendo en cuenta el intervalo de tiempo de cada columna.
Private Function TimeToColIndex(DateTime As Variant) As Integer
    TimeToColIndex = MinuteToColIndex((Hour(DateTime) * 60) + Minute(DateTime) - m_DayTimeStartMinutes)
End Function

'Esta Funcion me devuelve el índice de la columna correspondiente a los minutos que le paso
'teniendo en cuenta el intervalo de tiempo de cada columna
Private Function MinuteToColIndex(iMinutes As Integer) As Integer
    MinuteToColIndex = COLUMN_HOUR_FIRST + Int(iMinutes / m_IntervalMinutes)
End Function

Private Sub cboDiaHoraFin_Click()
    m_DayTimeEndMinutes = DateDiff("n", "00:00", cboDiaHoraFin)
    FormatGrid
End Sub

Private Sub cboDiaHoraInicio_Click()
    m_DayTimeStartMinutes = DateDiff("n", "00:00", cboDiaHoraInicio)
    FormatGrid
End Sub

Private Sub cboIntervalo_Click()
    FormatGrid
End Sub

Private Sub cmdFechaAnterior_Click()
    dtpFecha.Value = DateAdd("d", -1, dtpFecha.Value)
    dtpFecha.SetFocus
    dtpFecha_Change
End Sub

Private Sub dtpFecha_Change()
    txtFechaDiaSemana.Text = WeekdayName(Weekday(dtpFecha.Value))
    FillGrid
End Sub

Private Sub cmdFechaSiguiente_Click()
    dtpFecha.Value = DateAdd("d", 1, dtpFecha.Value)
    dtpFecha.SetFocus
    dtpFecha_Change
End Sub

Private Sub cmdFechaHoy_Click()
    Dim OldValue As Date
    
    OldValue = dtpFecha.Value
    dtpFecha.Value = Date
    dtpFecha.SetFocus
    If OldValue <> dtpFecha.Value Then
        dtpFecha_Change
    End If
End Sub

Private Sub Form_Load()
    Dim Intervalo As Integer
    Dim CIntervalos As Collection
    Dim CIntervalosNombres As Collection
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings Mid(Me.Name, 4), cbrMain
    
    'HORAS DEL DIA
    Call CSM_Control_ComboBox.FillWithHoursAndMinutes(cboDiaHoraInicio, True, CDate("00:00"), CDate("23:59"), 30)
    Call CSM_Control_ComboBox.FillWithHoursAndMinutes(cboDiaHoraFin, False, CDate("00:00"), CDate("23:59"), 30)
    cboDiaHoraInicio.ListIndex = pCSC_Parameter.GetParameterNumberInteger("VehiculoUtilizacion_HoraInicio", 11)
    cboDiaHoraFin.ListIndex = cboDiaHoraFin.ListCount - 1
    
    'CARGO LOS INTERVALOS
    Set CIntervalos = pCSC_Parameter.GetParameterCollection("VehiculoUtilizacion_Intervalos", "5;10;20;30;45;60;120;240;360;480;720")
    Set CIntervalosNombres = pCSC_Parameter.GetParameterCollection("VehiculoUtilizacion_IntervalosNombres", "5 minutos.;10 minutos.;20 minutos.;30 minutos.;45 minutos.;1 hora.;2 horas.;4 horas.;6 horas.;8 horas.;12 horas.")
    
    For Intervalo = 1 To CIntervalos.Count
        cboIntervalo.AddItem CIntervalosNombres(Intervalo)
        cboIntervalo.ItemData(cboIntervalo.NewIndex) = CIntervalos(Intervalo)
    Next Intervalo
    cboIntervalo.ListIndex = CSM_Control_ComboBox.GetListIndexByItemData(cboIntervalo, pCSC_Parameter.GetParameterNumberInteger("VehiculoUtilizacion_IntervaloPredeterminado", 30), cscpCurrentOrFirst)
    
    tdbgrdData.Splits(0).Columns(COLUMN_INDEX_VEHICULONOMBRE).Width = pCSC_Parameter.SetParameterNumberInteger("VehiculoUtilizacion_GridSplitWidth", 500)
    
    mVehiculoUtilizacion_ViajeEnProgreso_FinalizaEnHoraActual = pCSC_Parameter.GetParameterBoolean("VehiculoUtilizacion_ViajeEnProgreso_FinalizaEnHoraActual", True)
    
    dtpFecha.Value = Date
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub ResizeControls(ByVal CoolBarHeight As Single)
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
    
    tdbgrdData.Top = CoolBarHeight + CONTROL_SPACE
    tdbgrdData.Left = CONTROL_SPACE
    tdbgrdData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    tdbgrdData.Height = ScaleHeight - tdbgrdData.Top - CONTROL_SPACE - stbMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    Set mxaEstado = Nothing
    Set mxaFechaHora = Nothing
    Set mxaIDRuta = Nothing
    Set mCViajesEnCeldaActual = Nothing
    pParametro.SaveCoolBarSettings Mid(Me.Name, 4), cbrMain
    Set frmVehiculoUtilizacion = Nothing
End Sub

'Esta Sub formatea la grilla de Estado según el intervalo de cada columna
Private Sub FormatGrid()
    Dim intMinuto As Integer
    Dim intColIndex As Integer
    Dim intIndex As Integer
    Dim tdbgrdcolCurrent As TrueOleDBGrid80.Column
    Dim ColumnHourWidth As Integer
    
    If Not m_RefreshGrid Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'DIMENSIONO LOS VALUE ITEMS DE CADA ESTADO BASANDOME EN LA COLUMNA EN LOS QUE TIENE LA COLUMNA 00:00
    'PARA DESPUES ASIGNARSELOS A CADA COLUMNA:
    '===================================================================================================
    
    '### SIMPLES ###
    'ACTIVO
    Dim viActivo As New ValueItem
    viActivo.Value = VIAJE_ESTADO_ACTIVO
    viActivo.DisplayValue = LoadResPicture("GRID_ACTIVO", vbResBitmap)
    'EN PROGRESO
    Dim viEnProgreso As New ValueItem
    viEnProgreso.Value = VIAJE_ESTADO_EN_PROGRESO
    viEnProgreso.DisplayValue = LoadResPicture("GRID_ENPROGRESO", vbResBitmap)
    'FINALIZADO
    Dim viFinalizado As New ValueItem
    viFinalizado.Value = VIAJE_ESTADO_FINALIZADO
    viFinalizado.DisplayValue = LoadResPicture("GRID_FINALIZADO", vbResBitmap)
    'CANCELADO
    Dim viCancelado As New ValueItem
    viCancelado.Value = VIAJE_ESTADO_CANCELADO
    viCancelado.DisplayValue = LoadResPicture("GRID_CANCELADO", vbResBitmap)
    
    '### DOBLES - PRIMERO ACTIVO ###
    'ACTIVO
    Dim viActivo_Activo As New ValueItem
    viActivo_Activo.Value = VIAJE_ESTADO_ACTIVO & KEY_DELIMITER & VIAJE_ESTADO_ACTIVO
    viActivo_Activo.DisplayValue = LoadResPicture("GRID_ACTIVO_ACTIVO", vbResBitmap)
    'EN PROGRESO
    Dim viActivo_EnProgreso As New ValueItem
    viActivo_EnProgreso.Value = VIAJE_ESTADO_ACTIVO & KEY_DELIMITER & VIAJE_ESTADO_EN_PROGRESO
    viActivo_EnProgreso.DisplayValue = LoadResPicture("GRID_ACTIVO_ENPROGRESO", vbResBitmap)
    'FINALIZADO
    Dim viActivo_Finalizado As New ValueItem
    viActivo_Finalizado.Value = VIAJE_ESTADO_ACTIVO & KEY_DELIMITER & VIAJE_ESTADO_FINALIZADO
    viActivo_Finalizado.DisplayValue = LoadResPicture("GRID_ACTIVO_FINALIZADO", vbResBitmap)
    'CANCELADO
    Dim viActivo_Cancelado As New ValueItem
    viActivo_Cancelado.Value = VIAJE_ESTADO_ACTIVO & KEY_DELIMITER & VIAJE_ESTADO_CANCELADO
    viActivo_Cancelado.DisplayValue = LoadResPicture("GRID_ACTIVO_CANCELADO", vbResBitmap)
    
    '### DOBLES - PRIMERO EN PROGRESO ###
    'ACTIVO
    Dim viEnProgreso_Activo As New ValueItem
    viEnProgreso_Activo.Value = VIAJE_ESTADO_EN_PROGRESO & KEY_DELIMITER & VIAJE_ESTADO_ACTIVO
    viEnProgreso_Activo.DisplayValue = LoadResPicture("GRID_ENPROGRESO_ACTIVO", vbResBitmap)
    'EN PROGRESO
    Dim viEnProgreso_EnProgreso As New ValueItem
    viEnProgreso_EnProgreso.Value = VIAJE_ESTADO_EN_PROGRESO & KEY_DELIMITER & VIAJE_ESTADO_EN_PROGRESO
    viEnProgreso_EnProgreso.DisplayValue = LoadResPicture("GRID_ENPROGRESO_ENPROGRESO", vbResBitmap)
    'FINALIZADO
    Dim viEnProgreso_Finalizado As New ValueItem
    viEnProgreso_Finalizado.Value = VIAJE_ESTADO_EN_PROGRESO & KEY_DELIMITER & VIAJE_ESTADO_FINALIZADO
    viEnProgreso_Finalizado.DisplayValue = LoadResPicture("GRID_ENPROGRESO_FINALIZADO", vbResBitmap)
    'CANCELADO
    Dim viEnProgreso_Cancelado As New ValueItem
    viEnProgreso_Cancelado.Value = VIAJE_ESTADO_EN_PROGRESO & KEY_DELIMITER & VIAJE_ESTADO_CANCELADO
    viEnProgreso_Cancelado.DisplayValue = LoadResPicture("GRID_ENPROGRESO_CANCELADO", vbResBitmap)
    
    '### DOBLES - PRIMERO FINALIZADO ###
    'ACTIVO
    Dim viFinalizado_Activo As New ValueItem
    viFinalizado_Activo.Value = VIAJE_ESTADO_FINALIZADO & KEY_DELIMITER & VIAJE_ESTADO_ACTIVO
    viFinalizado_Activo.DisplayValue = LoadResPicture("GRID_FINALIZADO_ACTIVO", vbResBitmap)
    'EN PROGRESO
    Dim viFinalizado_EnProgreso As New ValueItem
    viFinalizado_EnProgreso.Value = VIAJE_ESTADO_FINALIZADO & KEY_DELIMITER & VIAJE_ESTADO_EN_PROGRESO
    viFinalizado_EnProgreso.DisplayValue = LoadResPicture("GRID_FINALIZADO_ENPROGRESO", vbResBitmap)
    'FINALIZADO
    Dim viFinalizado_Finalizado As New ValueItem
    viFinalizado_Finalizado.Value = VIAJE_ESTADO_FINALIZADO & KEY_DELIMITER & VIAJE_ESTADO_FINALIZADO
    viFinalizado_Finalizado.DisplayValue = LoadResPicture("GRID_FINALIZADO_FINALIZADO", vbResBitmap)
    'CANCELADO
    Dim viFinalizado_Cancelado As New ValueItem
    viFinalizado_Cancelado.Value = VIAJE_ESTADO_FINALIZADO & KEY_DELIMITER & VIAJE_ESTADO_CANCELADO
    viFinalizado_Cancelado.DisplayValue = LoadResPicture("GRID_FINALIZADO_CANCELADO", vbResBitmap)
    
    '### DOBLES - PRIMERO CANCELADO###
    'ACTIVO
    Dim viCancelado_Activo As New ValueItem
    viCancelado_Activo.Value = VIAJE_ESTADO_CANCELADO & KEY_DELIMITER & VIAJE_ESTADO_ACTIVO
    viCancelado_Activo.DisplayValue = LoadResPicture("GRID_CANCELADO_ACTIVO", vbResBitmap)
    'EN PROGRESO
    Dim viCancelado_EnProgreso As New ValueItem
    viCancelado_EnProgreso.Value = VIAJE_ESTADO_CANCELADO & KEY_DELIMITER & VIAJE_ESTADO_EN_PROGRESO
    viCancelado_EnProgreso.DisplayValue = LoadResPicture("GRID_CANCELADO_ENPROGRESO", vbResBitmap)
    'FINALIZADO
    Dim viCancelado_Finalizado As New ValueItem
    viCancelado_Finalizado.Value = VIAJE_ESTADO_CANCELADO & KEY_DELIMITER & VIAJE_ESTADO_FINALIZADO
    viCancelado_Finalizado.DisplayValue = LoadResPicture("GRID_CANCELADO_FINALIZADO", vbResBitmap)
    'CANCELADO
    Dim viCancelado_Cancelado As New ValueItem
    viCancelado_Cancelado.Value = VIAJE_ESTADO_CANCELADO & KEY_DELIMITER & VIAJE_ESTADO_CANCELADO
    viCancelado_Cancelado.DisplayValue = LoadResPicture("GRID_CANCELADO_CANCELADO", vbResBitmap)
    
    '### TRIPLES O MÁS --> EN CONFLICTO
    Dim viConflicto As New ValueItem
    viConflicto.Value = RESERVA_CONFLICTO
    viConflicto.DisplayValue = LoadResPicture("GRID_CONFLICTO", vbResBitmap)
    
    '===================================================================================================
    '===ESTOS ESTILOS SON PARA LOS INICIOS  Y  LOS FINALES==============================================
    '===================================================================================================
    'INICIOS
    'Activo
    Dim viActivo_Inicio As New ValueItem
    viActivo_Inicio.Value = VIAJE_ESTADO_ACTIVO & RESERVA_INICIO
    viActivo_Inicio.DisplayValue = LoadResPicture("GRID_ACTIVO_INICIO", vbResBitmap)
    'En Progreso
    Dim viEnProgreso_Inicio As New ValueItem
    viEnProgreso_Inicio.Value = VIAJE_ESTADO_EN_PROGRESO & RESERVA_INICIO
    viEnProgreso_Inicio.DisplayValue = LoadResPicture("GRID_EnProgreso_INICIO", vbResBitmap)
    'Finalizado
    Dim viFinalizado_Inicio As New ValueItem
    viFinalizado_Inicio.Value = VIAJE_ESTADO_FINALIZADO & RESERVA_INICIO
    viFinalizado_Inicio.DisplayValue = LoadResPicture("GRID_Finalizado_INICIO", vbResBitmap)
    'Cancelado
    Dim viCancelado_Inicio As New ValueItem
    viCancelado_Inicio.Value = VIAJE_ESTADO_CANCELADO & RESERVA_INICIO
    viCancelado_Inicio.DisplayValue = LoadResPicture("GRID_Cancelado_INICIO", vbResBitmap)

    'FINALES
    'Activo
    Dim viActivo_Fin As New ValueItem
    viActivo_Fin.Value = VIAJE_ESTADO_ACTIVO & RESERVA_FIN
    viActivo_Fin.DisplayValue = LoadResPicture("GRID_ACTIVO_Fin", vbResBitmap)
    'En Progreso
    Dim viEnProgreso_Fin As New ValueItem
    viEnProgreso_Fin.Value = VIAJE_ESTADO_EN_PROGRESO & RESERVA_FIN
    viEnProgreso_Fin.DisplayValue = LoadResPicture("GRID_EnProgreso_Fin", vbResBitmap)
    'Finalizado
    Dim viFinalizado_Fin As New ValueItem
    viFinalizado_Fin.Value = VIAJE_ESTADO_FINALIZADO & RESERVA_FIN
    viFinalizado_Fin.DisplayValue = LoadResPicture("GRID_Finalizado_Fin", vbResBitmap)
    'Cancelado
    Dim viCancelado_Fin As New ValueItem
    viCancelado_Fin.Value = VIAJE_ESTADO_CANCELADO & RESERVA_FIN
    viCancelado_Fin.DisplayValue = LoadResPicture("GRID_Cancelado_Fin", vbResBitmap)
    
    'frmWorking.Show
    'fraFormatGridProgress.Visible = True
    tdbgrdData.Visible = False
    
    'Formateo la Grilla
    If cboIntervalo.ListIndex = -1 Then
        Exit Sub
    End If
    
    m_IntervalMinutes = cboIntervalo.ItemData(cboIntervalo.ListIndex)
    m_ColumnHourLast = COLUMN_HOUR_FIRST + Int((m_DayTimeEndMinutes - m_DayTimeStartMinutes) / m_IntervalMinutes)
    
    If m_ColumnHourLast < COLUMN_HOUR_FIRST - 1 Then m_ColumnHourLast = COLUMN_HOUR_FIRST - 1
    
    tdbgrdData.RowHeight = pCSC_Parameter.GetParameterNumberInteger("VehiculoUtilizacion_GridRowHeight", 500)
    ColumnHourWidth = pCSC_Parameter.GetParameterNumberInteger("VehiculoUtilizacion_GridColumnWidth", 550)
    
    For intMinuto = m_DayTimeStartMinutes To m_DayTimeEndMinutes Step m_IntervalMinutes
        'Indice de la Columna
        intColIndex = MinuteToColIndex(intMinuto - m_DayTimeStartMinutes)
        
        If tdbgrdData.Columns.Count < intColIndex + 1 Then
            Set tdbgrdcolCurrent = tdbgrdData.Columns.Add(intColIndex)
        End If
        
        'SPLIT 0
        tdbgrdData.Splits(0).Columns(intColIndex).Visible = False
        tdbgrdData.Splits(0).Columns(intColIndex).AllowSizing = False
        tdbgrdData.Splits(0).Columns(intColIndex).Locked = True
        
        'SPLIT 1
        tdbgrdData.Splits(1).Columns(intColIndex).Visible = True
        tdbgrdData.Columns(intColIndex).DividerStyle = COLUMN_DIVIDER_STYLE
        tdbgrdData.Splits(1).Columns(intColIndex).AllowSizing = False
        tdbgrdData.Splits(1).Columns(intColIndex).Width = ColumnHourWidth
        tdbgrdData.Splits(1).Columns(intColIndex).Locked = True
        tdbgrdData.Splits(1).Columns(intColIndex).WrapText = True
        tdbgrdData.Splits(1).Columns(intColIndex).FetchStyle = True
        
        'ALINEACIONES
        tdbgrdData.Splits(1).Columns(intColIndex).Alignment = dbgCenter
        tdbgrdData.Splits(1).Columns(intColIndex).Style.VerticalAlignment = dbgVertCenter
        tdbgrdData.Splits(1).Columns(intColIndex).HeadAlignment = dbgCenter
        
        'Título de la Columna
        tdbgrdData.Columns(intColIndex).Caption = Format(DateAdd("n", intMinuto, "00:00"), "hh:mm") + " - " + Format(DateAdd("n", IIf(intMinuto + m_IntervalMinutes - 1 > m_DayTimeEndMinutes, m_DayTimeEndMinutes, intMinuto + m_IntervalMinutes - 1), "00:00"), "hh:mm")
        
        'frmWorking.ProgressBar.Value = ((intColIndex - COLUMN_HOUR_FIRST + 1) / (m_ColumnHourLast - COLUMN_HOUR_FIRST + 1)) * 100
        'prgFormatGrid.Value = ((intColIndex - COLUMN_HOUR_FIRST + 1) / (m_ColumnHourLast - COLUMN_HOUR_FIRST + 1)) * 100
        
        'ASIGNO LOS VALUE ITEMS A LA COLUMNA
        tdbgrdData.Columns(intColIndex).ValueItems.Translate = True
        tdbgrdData.Columns(intColIndex).ValueItems.Add viActivo
        tdbgrdData.Columns(intColIndex).ValueItems.Add viEnProgreso
        tdbgrdData.Columns(intColIndex).ValueItems.Add viFinalizado
        tdbgrdData.Columns(intColIndex).ValueItems.Add viCancelado
        
        tdbgrdData.Columns(intColIndex).ValueItems.Add viActivo_Activo
        tdbgrdData.Columns(intColIndex).ValueItems.Add viActivo_EnProgreso
        tdbgrdData.Columns(intColIndex).ValueItems.Add viActivo_Finalizado
        tdbgrdData.Columns(intColIndex).ValueItems.Add viActivo_Cancelado
        
        tdbgrdData.Columns(intColIndex).ValueItems.Add viEnProgreso_Activo
        tdbgrdData.Columns(intColIndex).ValueItems.Add viEnProgreso_EnProgreso
        tdbgrdData.Columns(intColIndex).ValueItems.Add viEnProgreso_Finalizado
        tdbgrdData.Columns(intColIndex).ValueItems.Add viEnProgreso_Cancelado
        
        tdbgrdData.Columns(intColIndex).ValueItems.Add viFinalizado_Activo
        tdbgrdData.Columns(intColIndex).ValueItems.Add viFinalizado_EnProgreso
        tdbgrdData.Columns(intColIndex).ValueItems.Add viFinalizado_Finalizado
        tdbgrdData.Columns(intColIndex).ValueItems.Add viFinalizado_Cancelado
        
        tdbgrdData.Columns(intColIndex).ValueItems.Add viCancelado_Activo
        tdbgrdData.Columns(intColIndex).ValueItems.Add viCancelado_EnProgreso
        tdbgrdData.Columns(intColIndex).ValueItems.Add viCancelado_Finalizado
        tdbgrdData.Columns(intColIndex).ValueItems.Add viCancelado_Cancelado
        
        tdbgrdData.Columns(intColIndex).ValueItems.Add viConflicto
        

        'Para los inicios de cada grupo
        tdbgrdData.Columns(intColIndex).ValueItems.Add viActivo_Inicio
        tdbgrdData.Columns(intColIndex).ValueItems.Add viEnProgreso_Inicio
        tdbgrdData.Columns(intColIndex).ValueItems.Add viFinalizado_Inicio
        tdbgrdData.Columns(intColIndex).ValueItems.Add viCancelado_Inicio

        'Para los Finales de cada grupo
        tdbgrdData.Columns(intColIndex).ValueItems.Add viActivo_Fin
        tdbgrdData.Columns(intColIndex).ValueItems.Add viEnProgreso_Fin
        tdbgrdData.Columns(intColIndex).ValueItems.Add viFinalizado_Fin
        tdbgrdData.Columns(intColIndex).ValueItems.Add viCancelado_Fin
                
        DoEvents
        
    Next intMinuto
        
    For intIndex = tdbgrdData.Columns.Count - 1 To m_ColumnHourLast + 1 Step -1
        tdbgrdData.Columns.Remove intIndex
    Next intIndex
    
    tdbgrdData.Visible = True
    'Unload frmWorking
    'fraFormatGridProgress.Visible = False
    
    'Hago un Refresh
    RefreshGrid
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub RefreshGrid()
    Dim RowTopSave As Long
    Dim RowSave As Long
    Dim LeftColSave As Long
    Dim ColSave As Long
    
    If Not m_RefreshGrid Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    If tdbgrdData.Row > -1 And Not IsNull(tdbgrdData.FirstRow) Then
        RowTopSave = tdbgrdData.FirstRow
        RowSave = tdbgrdData.Row
        LeftColSave = tdbgrdData.LeftCol
        ColSave = tdbgrdData.Col
    End If
    tdbgrdData.Visible = False
    Call FillGrid
    
    On Error Resume Next
    tdbgrdData.FirstRow = RowTopSave
    tdbgrdData.Row = RowSave
    tdbgrdData.LeftCol = LeftColSave
    tdbgrdData.Col = ColSave
    tdbgrdData.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Public Sub FillGrid()
    Dim recVehiculo As New ADODB.Recordset
    Dim lngRowCountEstimated As Long
    Dim lngRowCountReal As Long
    
    Dim recEstado As New ADODB.Recordset
    Dim RowIndex As Long
    
    'Limpio todos los arrays
    Set mxaEstado = New XArrayDBObject.XArrayDB
    Set mxaFechaHora = New XArrayDBObject.XArrayDB
    Set mxaIDRuta = New XArrayDBObject.XArrayDB

    Set mCViajesEnCeldaActual = New Collection
        
    Call pDatabase.OpenRecordset(recVehiculo, "SELECT IDVehiculo, Nombre FROM Vehiculo WHERE Activo = 1 ORDER BY Nombre", adOpenForwardOnly, adLockReadOnly, adCmdText, "Error al leer la lista de vehículos.", "Forms.VehiculoUtilizacion.FillGrid")
    
    lngRowCountEstimated = 1
    lngRowCountReal = 0
    mxaEstado.ReDim 1, lngRowCountEstimated, 0, m_ColumnHourLast
    mxaFechaHora.ReDim 1, lngRowCountEstimated, 0, m_ColumnHourLast
    mxaIDRuta.ReDim 1, lngRowCountEstimated, 0, m_ColumnHourLast
    Do While Not recVehiculo.EOF
        lngRowCountReal = lngRowCountReal + 1
        If lngRowCountReal > lngRowCountEstimated Then
            lngRowCountEstimated = lngRowCountEstimated * 2
            mxaEstado.ReDim 1, lngRowCountEstimated, 0, m_ColumnHourLast
            mxaFechaHora.ReDim 1, lngRowCountEstimated, 0, m_ColumnHourLast
            mxaIDRuta.ReDim 1, lngRowCountEstimated, 0, m_ColumnHourLast
        End If
        mxaEstado(lngRowCountReal, 0) = recVehiculo("IDVehiculo").Value
        mxaEstado(lngRowCountReal, 1) = recVehiculo("Nombre").Value
        
        recVehiculo.MoveNext
    Loop
    mxaEstado.ReDim 1, lngRowCountReal, 0, m_ColumnHourLast
    mxaFechaHora.ReDim 1, lngRowCountReal, 0, m_ColumnHourLast
    mxaIDRuta.ReDim 1, lngRowCountReal, 0, m_ColumnHourLast
    
    Call pDatabase.OpenRecordset(recEstado, "sp_Vehiculo_Utilizacion " & CSM_String.FormatDateTimeToSQL(dtpFecha.Value), adOpenStatic, adLockReadOnly, adCmdText, "Error al leer los viajes del día.", "Forms.VehiculoUtilizacion.FillGrid")
    
    If (Not recEstado.BOF) And (Not recEstado.EOF) Then
        For RowIndex = 1 To lngRowCountReal
            recEstado.MoveFirst
            Do While Not recEstado.EOF
                If recEstado("IDVehiculo").Value = mxaEstado(RowIndex, 0) Then
                    PaintGrid recEstado("FechaHora").Value, recEstado("IDRuta").Value, recEstado("Estado").Value, recEstado("FechaHora").Value, DateAdd("n", IIf(IsNull(recEstado("Duracion").Value), 1, recEstado("Duracion").Value), recEstado("FechaHora").Value), RowIndex
                End If
                recEstado.MoveNext
            Loop
        Next RowIndex
    End If
    
    Set tdbgrdData.Array = mxaEstado
    tdbgrdData.ReBind
    tdbgrdData.MoveFirst
    
    recVehiculo.Close
    Set recVehiculo = Nothing
    recEstado.Close
    Set recEstado = Nothing
End Sub

Private Sub tdbgrdData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim CFechaHora As Collection
    Dim CIDRuta As Collection
    Dim Indice As Integer
    Dim StatusText As String

    Dim Viaje As Viaje
    
    If tdbgrdData.Row = -1 Then
        'Nothing
        Set mCViajesEnCeldaActual = Nothing
    Else
        StatusText = "Vehículo: " & tdbgrdData.Columns(COLUMN_INDEX_VEHICULONOMBRE).Value
        
        Set mCViajesEnCeldaActual = New Collection
        If Not IsEmpty(mxaFechaHora(tdbgrdData.Row + tdbgrdData.FirstRow, tdbgrdData.Col)) Then
            Set CFechaHora = mxaFechaHora(tdbgrdData.Row + tdbgrdData.FirstRow, tdbgrdData.Col)
            Set CIDRuta = mxaIDRuta(tdbgrdData.Row + tdbgrdData.FirstRow, tdbgrdData.Col)
            
            For Indice = 1 To CFechaHora.Count
                Set Viaje = New Viaje
                Viaje.FechaHora = CDate(CFechaHora(Indice))
                Viaje.IDRuta = CStr(CIDRuta(Indice))
                If Viaje.Load() Then
                    StatusText = StatusText & " || Fecha/Hora: " & Viaje.FechaHora_Formatted & " - Ruta: " & Viaje.Ruta_DisplayName
                    mCViajesEnCeldaActual.Add Viaje
                End If
                Set Viaje = Nothing
            Next Indice
        End If
        stbMain.SimpleText = StatusText
    End If
End Sub

Private Sub tdbgrdData_DblClick()
    Dim Viaje As Viaje
    
    If tdbgrdData.Columns(tdbgrdData.Col) = "" Then
        MsgBox "No hay ninguna Comodidad para Modificar.", vbInformation, App.Title
    Else
        If pCPermiso.GotPermission(PERMISO_VIAJE) Then
            Set Viaje = frmViajeSelect.LoadDataAndShow(tdbgrdData.Columns(COLUMN_INDEX_VEHICULONOMBRE).Value, dtpFecha.Value & " " & Left(tdbgrdData.Columns(tdbgrdData.Col).Caption, 5), dtpFecha.Value & " " & Right(tdbgrdData.Columns(tdbgrdData.Col).Caption, 5), mCViajesEnCeldaActual)
            If Not Viaje Is Nothing Then
                If pCPermiso.GotPermission(PERMISO_VIAJE_MODIFY) Then
                    Screen.MousePointer = vbHourglass
                    frmViajePropiedad.LoadDataAndShow Me, Viaje
                    Set Viaje = Nothing
                    Screen.MousePointer = vbDefault
                End If
            End If
        End If
    End If
End Sub
