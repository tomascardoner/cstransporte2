VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmPersona 
   Caption         =   "Personas"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13365
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Persona.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   13365
   Begin TrueOleDBGrid80.TDBGrid tdbgrdData 
      Height          =   3675
      Left            =   120
      TabIndex        =   0
      Top             =   1620
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   6482
      _LayoutType     =   4
      _RowHeight      =   13
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "IDPersona"
      Columns(0).DataField=   "IDPersona"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Apellido"
      Columns(1).DataField=   "Apellido"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nombre"
      Columns(2).DataField=   "Nombre"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Tipo"
      Columns(3).DataField=   "EntidadTipo"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Documento"
      Columns(4).DataField=   "Documento"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Activo"
      Columns(5).DataField=   "Activo"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Lista de Pasajeros"
      Columns(6).DataField=   "ListaPasajero"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._VertColor=12632256"
      Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=8811"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=8731"
      Splits(0)._ColumnProps(12)=   "Column(1)._VertColor=12632256"
      Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(16)=   "Column(2).Width=6165"
      Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=6085"
      Splits(0)._ColumnProps(19)=   "Column(2)._VertColor=12632256"
      Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(23)=   "Column(3).Width=2461"
      Splits(0)._ColumnProps(24)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._WidthInPix=2381"
      Splits(0)._ColumnProps(26)=   "Column(3)._VertColor=12632256"
      Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=8196"
      Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(30)=   "Column(4).Width=3175"
      Splits(0)._ColumnProps(31)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(4)._WidthInPix=3096"
      Splits(0)._ColumnProps(33)=   "Column(4)._VertColor=12632256"
      Splits(0)._ColumnProps(34)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(4)._ColStyle=8196"
      Splits(0)._ColumnProps(36)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(37)=   "Column(5).Width=1402"
      Splits(0)._ColumnProps(38)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(5)._WidthInPix=1323"
      Splits(0)._ColumnProps(40)=   "Column(5)._VertColor=12632256"
      Splits(0)._ColumnProps(41)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(5)._ColStyle=8705"
      Splits(0)._ColumnProps(43)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(44)=   "Column(6).Width=2646"
      Splits(0)._ColumnProps(45)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(6)._WidthInPix=2566"
      Splits(0)._ColumnProps(47)=   "Column(6)._VertColor=12632256"
      Splits(0)._ColumnProps(48)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(49)=   "Column(6)._ColStyle=8705"
      Splits(0)._ColumnProps(50)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   0
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   -2147483643
      RowDividerColor =   16777215
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   0
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.locked=-1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.locked=-1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(64)  =   "Named:id=33:Normal"
      _StyleDefs(65)  =   ":id=33,.parent=0"
      _StyleDefs(66)  =   "Named:id=34:Heading"
      _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   ":id=34,.wraptext=-1"
      _StyleDefs(69)  =   "Named:id=35:Footing"
      _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   "Named:id=36:Selected"
      _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=37:Caption"
      _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(75)  =   "Named:id=38:HighlightRow"
      _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=39:EvenRow"
      _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(79)  =   "Named:id=40:OddRow"
      _StyleDefs(80)  =   ":id=40,.parent=33"
      _StyleDefs(81)  =   "Named:id=41:RecordSelector"
      _StyleDefs(82)  =   ":id=41,.parent=34"
      _StyleDefs(83)  =   "Named:id=42:FilterBar"
      _StyleDefs(84)  =   ":id=42,.parent=33"
   End
   Begin VB.Timer tmrKeyDelay 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5880
      Top             =   2640
   End
   Begin MSComctlLib.TabStrip tabIndex 
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   1140
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   556
      TabMinWidth     =   18
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   27
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "A"
            Key             =   "A"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "B"
            Key             =   "B"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "C"
            Key             =   "C"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "D"
            Key             =   "D"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "E"
            Key             =   "E"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "F"
            Key             =   "F"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "G"
            Key             =   "G"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "H"
            Key             =   "H"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "I"
            Key             =   "I"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "J"
            Key             =   "J"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "K"
            Key             =   "K"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "L"
            Key             =   "L"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "M"
            Key             =   "M"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab14 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "N"
            Key             =   "N"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab15 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "O"
            Key             =   "O"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab16 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "P"
            Key             =   "P"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab17 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Q"
            Key             =   "Q"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab18 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "R"
            Key             =   "R"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab19 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "S"
            Key             =   "S"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab20 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "T"
            Key             =   "T"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab21 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "U"
            Key             =   "U"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab22 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "V"
            Key             =   "V"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab23 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "W"
            Key             =   "W"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab24 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "X"
            Key             =   "X"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab25 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Y"
            Key             =   "Y"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab26 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Z"
            Key             =   "Z"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab27 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Otro"
            Key             =   "OTRO"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   60
      TabIndex        =   10
      Top             =   5460
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
      Height          =   990
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   13365
      _ExtentX        =   23574
      _ExtentY        =   1746
      FixedOrder      =   -1  'True
      _CBWidth        =   13365
      _CBHeight       =   990
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain"
      MinWidth1       =   9915
      MinHeight1      =   570
      Width1          =   9915
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picFilterTipo"
      MinWidth2       =   1830
      MinHeight2      =   330
      Width2          =   1830
      Key2            =   "FilterTipo"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Child3          =   "picFilterActivo"
      MinWidth3       =   1605
      MinHeight3      =   330
      Width3          =   1605
      Key3            =   "FilterActivo"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Begin VB.PictureBox picFilterTipo 
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
         Left            =   11445
         ScaleHeight     =   330
         ScaleWidth      =   1830
         TabIndex        =   8
         Top             =   150
         Width           =   1830
         Begin VB.ComboBox cboFilterTipo 
            Height          =   330
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   0
            Width           =   1410
         End
         Begin VB.Label lblFilterTipo 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   210
            Left            =   0
            TabIndex        =   9
            Top             =   60
            Width           =   345
         End
      End
      Begin VB.PictureBox picFilterActivo 
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
         ScaleWidth      =   13110
         TabIndex        =   6
         Top             =   630
         Width           =   13110
         Begin VB.ComboBox cboFilterActivo 
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   0
            Width           =   990
         End
         Begin VB.Label lblFilterActivo 
            AutoSize        =   -1  'True
            Caption         =   "Activo:"
            Height          =   210
            Left            =   0
            TabIndex        =   7
            Top             =   60
            Width           =   510
         End
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   570
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   11190
         _ExtentX        =   19738
         _ExtentY        =   1005
         ButtonWidth     =   2170
         ButtonHeight    =   1005
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Nuevo"
               Key             =   "NEW"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Propiedades"
               Key             =   "PROPERTIES"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Eliminar"
               Key             =   "DELETE"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Seleccionar"
               Key             =   "SELECT"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Horarios"
               Key             =   "HORARIO"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Rutas"
               Key             =   "RUTA"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Información"
               Key             =   "INFO"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Respuestas"
               Key             =   "RESPUESTA"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Prepagos"
               Key             =   "PREPAGO"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5445
      Width           =   13365
      _ExtentX        =   23574
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
            Object.Width           =   22357
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
End
Attribute VB_Name = "frmPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean

Public FormWaitingForSelect As String
Public FormKeepOpenOnSelect As Boolean
Public SelectTypeFilter As String
Public SelectTag As String

Private mLastPressedKeyDelay As Long
Private mSearchString As String

Private mSortColumn As Integer
Private mSortOrderAscending As Boolean

Public Sub FillData(ByVal IDPersona As Long)
    Dim KeySave As Long
    Dim Persona As Persona
    
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    Dim OrderBy As String
    
    If mLoading Then
        Exit Sub
    End If
        
    Screen.MousePointer = vbHourglass
    
    If IDPersona = 0 Then
        If Val(tdbgrdData.FirstRow) <> 0 And Not IsNull(tdbgrdData.Columns("IDPersona").Value) Then
            KeySave = tdbgrdData.Columns("IDPersona").Value
        End If
    Else
        Set Persona = New Persona
        Persona.IDPersona = IDPersona
        If Persona.Load Then
            mLoading = True
            If UCase(Left(Persona.Apellido, 1)) >= "A" And UCase(Left(Persona.Apellido, 1)) <= "Z" Then
                Set tabIndex.SelectedItem = tabIndex.Tabs(UCase(Left(Persona.Apellido, 1)))
            Else
                Set tabIndex.SelectedItem = tabIndex.Tabs("OTRO")
            End If
            mLoading = False
        End If
        Set Persona = Nothing
        KeySave = IDPersona
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_Persona_ListGrid"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("FirstLetter", adChar, adParamInput, 1, IIf(tabIndex.SelectedItem.Key = "OTRO", Null, tabIndex.SelectedItem.Key))
    cmdData.Parameters.Append cmdData.CreateParameter("EntidadTipo", adChar, adParamInput, 2, Choose(cboFilterTipo.ListIndex, ENTIDAD_TIPO_PERSONA_CLIENTE, ENTIDAD_TIPO_PERSONA_CONDUCTOR, ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO))
    cmdData.Parameters.Append cmdData.CreateParameter("Activo", adBoolean, adParamInput, , Switch(cboFilterActivo.ListIndex = 0, Null, cboFilterActivo.ListIndex = 1, 1, cboFilterActivo.ListIndex = 2, 0))
    Set recData = New ADODB.Recordset
    Set recData.ActiveConnection = pDatabase.Connection
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    Set cmdData = Nothing
        
    Select Case mSortColumn
        Case 1  'APELLIDO
            OrderBy = "Apellido" & IIf(mSortOrderAscending, "", " DESC") & ", Nombre" & IIf(mSortOrderAscending, "", " DESC") & ", EntidadTipo" & IIf(mSortOrderAscending, "", " DESC")
        Case 2  'NOMBRE
            OrderBy = "Nombre" & IIf(mSortOrderAscending, "", " DESC") & ", Apellido" & IIf(mSortOrderAscending, "", " DESC") & ", EntidadTipo" & IIf(mSortOrderAscending, "", " DESC")
        Case 3  'ENTIDAD TIPO
            OrderBy = "EntidadTipo" & IIf(mSortOrderAscending, "", " DESC") & ", Apellido" & IIf(mSortOrderAscending, "", " DESC") & ", Nombre" & IIf(mSortOrderAscending, "", " DESC")
        Case 4  'DOCUMENTO NUMERO
            OrderBy = "IDDocumentoTipo" & IIf(mSortOrderAscending, "", " DESC") & ", DocumentoNumero" & IIf(mSortOrderAscending, "", " DESC") & ", Apellido" & IIf(mSortOrderAscending, "", " DESC") & ", Nombre" & IIf(mSortOrderAscending, "", " DESC") & ", EntidadTipo" & IIf(mSortOrderAscending, "", " DESC")
        Case 5  'ACTIVO
            OrderBy = "Activo" & IIf(mSortOrderAscending, "", " DESC") & ", Apellido" & IIf(mSortOrderAscending, "", " DESC") & ", Nombre" & IIf(mSortOrderAscending, "", " DESC") & ", EntidadTipo" & IIf(mSortOrderAscending, "", " DESC")
        Case 6  'LISTA DE PASAJEROS
            OrderBy = "ListaPasajero" & IIf(mSortOrderAscending, "", " DESC") & ", Apellido" & IIf(mSortOrderAscending, "", " DESC") & ", Nombre" & IIf(mSortOrderAscending, "", " DESC") & ", EntidadTipo" & IIf(mSortOrderAscending, "", " DESC")
    End Select
    
    recData.Sort = OrderBy
    If KeySave > 0 And Not (recData.BOF And recData.EOF) Then
        recData.Find "IDPersona = " & KeySave
        If recData.EOF Then
            recData.MoveFirst
        End If
    End If
    
    Set tdbgrdData.DataSource = recData
    tdbgrdData.ReBind
    
    Set recData = Nothing
    
    On Error Resume Next
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        tdbgrdData.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.Persona.FillData", "Error al leer la Lista de Personas."
End Sub

Public Sub FindAndShowItem(ByVal IDPersona As Long, ByVal CapitalLetter As String, ByVal FormNameWaitingForSelect As String, ByVal SelectTypeFilterValue As String, ByVal SelectTagValue As String)
    Dim recData As ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    Load frmPersona
    If IDPersona > 0 Then
        If CapitalLetter >= "A" And CapitalLetter <= "Z" Then
            Set frmPersona.tabIndex.SelectedItem = frmPersona.tabIndex.Tabs(CapitalLetter)
        Else
            Set frmPersona.tabIndex.SelectedItem = frmPersona.tabIndex.Tabs("OTRO")
        End If
    End If
    frmPersona.Show
    
    On Error Resume Next
    
    Set recData = tdbgrdData.DataSource
    If Not (recData.BOF And recData.EOF) Then
        recData.MoveFirst
        recData.Find "IDPersona = " & IDPersona
        If recData.EOF Then
            recData.MoveFirst
        End If
    End If
    
    If frmPersona.WindowState = vbMinimized Then
        frmPersona.WindowState = vbNormal
    End If
    SelectTypeFilter = SelectTypeFilterValue
    SelectTag = SelectTagValue
    FormWaitingForSelect = FormNameWaitingForSelect
    frmPersona.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call tdbgrdData_DblClick
    End If
End Sub

Private Sub Form_Load()
    mLoading = True
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL TOOLBAR
    Set tlbMain.ImageList = frmMDI.ilsFormToolbar
    Set tlbMain.HotImageList = frmMDI.ilsFormToolbarHot
    tlbMain.Buttons("NEW").Image = "NEW"
    tlbMain.Buttons("PROPERTIES").Image = "PROPERTIES"
    tlbMain.Buttons("DELETE").Image = "DELETE"
    tlbMain.Buttons("SELECT").Image = "SELECT"
    tlbMain.Buttons("HORARIO").Image = "HORARIO"
    tlbMain.Buttons("RUTA").Image = "RUTA"
    tlbMain.Buttons("INFO").Image = "INFO"
    tlbMain.Buttons("RESPUESTA").Image = "RESPUESTA"
    tlbMain.Buttons("PREPAGO").Image = "PREPAGO"
    
    cbrMain.Bands("Toolbar").MinWidth = CSM_Control_Toolbar.GetTotalWidth(tlbMain)
    '//////////////////////////////////////////////////////////
        
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL PIN
    Set tlbPin.ImageList = frmMDI.ilsFormPin
    '//////////////////////////////////////////////////////////
    
    cboFilterTipo.AddItem ITEM_ALL_MALE
    cboFilterTipo.AddItem ENTIDAD_TIPO_PERSONA_CLIENTE_NOMBRE
    cboFilterTipo.AddItem ENTIDAD_TIPO_PERSONA_CONDUCTOR_NOMBRE
    cboFilterTipo.AddItem ENTIDAD_TIPO_PERSONA_ADMINISTRATIVO_NOMBRE
    cboFilterTipo.ListIndex = 0
    
    cboFilterActivo.AddItem ITEM_ALL_MALE
    cboFilterActivo.AddItem "Sí"
    cboFilterActivo.AddItem "No"
    cboFilterActivo.ListIndex = FILTER_ACTIVO_LIST_INDEX
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "Persona", cbrMain
    
    mSortColumn = 1
    mSortOrderAscending = True
    Call pParametro.GetTrueDBGridSettings("Persona", tdbgrdData, mSortColumn, mSortOrderAscending)
    
    tdbgrdData.Splits(0).Columns(mSortColumn).HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
    'tdbgrdData.Splits(0).Columns(mSortColumn).HeadingStyle.ForegroundPicture = frmMDI.ilsFormSortColumn.ListImages(Abs(mSortOrderAscending)).Picture
    
    tlbPin.Buttons("PIN").Value = pParametro.Usuario_LeerNumero("Persona_Pin", tlbPin.Buttons("PIN").Value)
    If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If

    mLoading = False
    
    FillData 0
    
    FormKeepOpenOnSelect = False
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim PreviousItem As Long
    Dim recData As ADODB.Recordset
    
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyN
                tlbMain_ButtonClick tlbMain.Buttons.Item("NEW")
            Case vbKeyP
                tlbMain_ButtonClick tlbMain.Buttons.Item("PROPERTIES")
            Case vbKeyE
                tlbMain_ButtonClick tlbMain.Buttons.Item("DELETE")
            Case vbKeyS
                tlbMain_ButtonClick tlbMain.Buttons.Item("SELECT")
        End Select
    Else
        If (KeyCode >= vbKeyA And KeyCode <= vbKeyZ) Or KeyCode = vbKeySpace Or KeyCode = 192 Then
            If tmrKeyDelay.Enabled Then
                If KeyCode = 192 Then
                    mSearchString = mSearchString & "Ñ"
                Else
                    mSearchString = mSearchString & Chr(KeyCode)
                End If
                
                Set recData = tdbgrdData.DataSource
                If Not (recData.BOF And recData.EOF) Then
                    PreviousItem = tdbgrdData.Columns("IDPersona").Value
                    
                    recData.MoveFirst
                    recData.Find "Apellido LIKE '" & mSearchString & "*'"
                    
                    If recData.EOF Then
                        recData.MoveFirst
                        recData.Find "IDPersona = " & PreviousItem
                        If recData.EOF Then
                            recData.MoveFirst
                        End If
                    End If
                End If
                Set recData = Nothing

                tdbgrdData.SetFocus
            Else
                If KeyCode <> vbKeySpace And KeyCode <> 192 Then
                    mSearchString = Chr(KeyCode)
                    If tabIndex.SelectedItem.Key <> Chr(KeyCode) Then
                        Set tabIndex.SelectedItem = tabIndex.Tabs(Chr(KeyCode))
                    End If
                End If
                KeyCode = 0
            End If
            mLastPressedKeyDelay = 0
            tmrKeyDelay.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    Call pParametro.SaveCoolBarSettings("Persona", cbrMain)
    Call pParametro.SaveTrueDBGridSettings("Persona", tdbgrdData, mSortColumn, mSortOrderAscending)
    Call pParametro.Usuario_GuardarNumero("Persona_Pin", tlbPin.Buttons("PIN").Value)
End Sub

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    ResizeControls NewHeight
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim FormIndex As Long
    Dim Persona As Persona
    Dim Feriado As Feriado
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_PERSONA_ADD) Then
                Screen.MousePointer = vbHourglass
                
                Set Persona = New Persona
                frmPersonaPropiedad.LoadDataAndShow Me, Persona
                Set Persona = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_PERSONA_MODIFY) Then
                If Val(tdbgrdData.FirstRow) = 0 Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                
                Set Persona = New Persona
                Persona.IDPersona = Val(tdbgrdData.Columns("IDPersona").Value)
                If Persona.Load() Then
                    frmPersonaPropiedad.LoadDataAndShow Me, Persona
                End If
                Set Persona = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_PERSONA_DELETE) Then
                If Val(tdbgrdData.FirstRow) = 0 Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
                If MsgBox("¿Desea eliminar la Persona seleccionada?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    Set Persona = New Persona
                    Persona.IDPersona = Val(tdbgrdData.Columns("IDPersona").Value)
                    If Persona.Load() Then
                        If Persona.Delete() Then
                            SetLastPersona 0, ""
                        End If
                    End If
                    Set Persona = Nothing
                End If
            End If
            tdbgrdData.SetFocus
        Case "SELECT"
            FormIndex = GetFormIndex(FormWaitingForSelect)
            FormWaitingForSelect = ""
            If FormIndex >= 0 Then
                If Val(tdbgrdData.FirstRow) = 0 Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
                
                Set Persona = New Persona
                Persona.IDPersona = Val(tdbgrdData.Columns("IDPersona").Value)
                If Not Persona.Load() Then
                    Set Persona = Nothing
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
                
'                If Not Persona.Activo Then
'                    Set Persona = Nothing
'                    MsgBox "No puede seleccionar este Item ya que está inactivo.", vbInformation, App.Title
'                    tdbgrdData.SetFocus
'                    Exit Sub
'                End If
                If SelectTypeFilter <> "" And SelectTypeFilter <> Persona.EntidadTipo Then
                    Set Persona = Nothing
                    MsgBox "Sólo se permite seleccionar Personas de tipo " & EntidadTipo_GetNombre(SelectTypeFilter) & ".", vbInformation, App.Title
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
    
                SelectTypeFilter = ""
                
                Screen.MousePointer = vbHourglass
                
                SetLastPersona Persona.IDPersona, Persona.ApellidoNombre
                
                Forms(FormIndex).PersonaSelected Persona.IDPersona, SelectTag
                Forms(FormIndex).SetFocus
                If tlbPin.Buttons("PIN").Value = tbrUnpressed And Not FormKeepOpenOnSelect Then
                    Unload frmPersona
                End If
                FormKeepOpenOnSelect = False
                Set Persona = Nothing
                
                If CSM_Forms.IsLoaded("frmPersonaSaldo") Then
                    frmPersonaSaldo.SetFocus
                End If
                
                Screen.MousePointer = vbDefault
            End If
        Case "HORARIO"
            If pCPermiso.GotPermission(PERMISO_PERSONA_HORARIO) Then
                If Val(tdbgrdData.FirstRow) = 0 Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
                                
                SetLastPersona Val(tdbgrdData.Columns("IDPersona").Value)
                
                Set Feriado = New Feriado
                Feriado.VerificarReservasDelPasajero Val(tdbgrdData.Columns("IDPersona").Value)
                Set Feriado = Nothing
                
                Screen.MousePointer = vbHourglass
                frmPersonaHorario.LoadDataAndShow Val(tdbgrdData.Columns("IDPersona").Value)
                Screen.MousePointer = vbDefault
            End If
        Case "RUTA"
            If pCPermiso.GotPermission(PERMISO_PERSONA_RUTA) Then
                If Val(tdbgrdData.FirstRow) = 0 Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
                
                Set Persona = New Persona
                Persona.IDPersona = Val(tdbgrdData.Columns("IDPersona").Value)
                If Not Persona.Load() Then
                    Set Persona = Nothing
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
                
                SetLastPersona Persona.IDPersona, Persona.ApellidoNombre
                
                Set Feriado = New Feriado
                Feriado.VerificarReservasDelPasajero Persona.IDPersona
                Set Feriado = Nothing
                
                Screen.MousePointer = vbHourglass
                frmPersonaRuta.LoadDataAndShow Persona
                Screen.MousePointer = vbDefault
                
                Set Persona = Nothing
            End If
        Case "INFO"
            If pCPermiso.GotPermission(PERMISO_PERSONA_INFO) Then
                If Val(tdbgrdData.FirstRow) = 0 Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
                
                Set Persona = New Persona
                Persona.IDPersona = Val(tdbgrdData.Columns("IDPersona").Value)
                If Not Persona.Load() Then
                    Set Persona = Nothing
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
                                
                SetLastPersona Persona.IDPersona, Persona.ApellidoNombre
                
                Set Feriado = New Feriado
                Feriado.VerificarReservasDelPasajero Persona.IDPersona
                Set Feriado = Nothing
                
                Screen.MousePointer = vbHourglass
                frmPersonaInfo.LoadDataAndShow Persona
                Screen.MousePointer = vbDefault
                
                Set Persona = Nothing
            End If
        Case "RESPUESTA"
            If pCPermiso.GotPermission(PERMISO_PERSONA_RESPUESTA) Then
                If Val(tdbgrdData.FirstRow) = 0 Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
                
                Set Persona = New Persona
                Persona.IDPersona = Val(tdbgrdData.Columns("IDPersona").Value)
                If Not Persona.Load() Then
                    Set Persona = Nothing
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
                                
                SetLastPersona Persona.IDPersona, Persona.ApellidoNombre
                
                Set Feriado = New Feriado
                Feriado.VerificarReservasDelPasajero Persona.IDPersona
                Set Feriado = Nothing
                
                Screen.MousePointer = vbHourglass
                frmPersonaRespuesta.LoadDataAndShow Persona
                Screen.MousePointer = vbDefault
                
                Set Persona = Nothing
            End If
        Case "PREPAGO"
            If pCPermiso.GotPermission(PERMISO_PERSONA_PREPAGO) Then
                If Val(tdbgrdData.FirstRow) = 0 Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
                
                Set Persona = New Persona
                Persona.IDPersona = Val(tdbgrdData.Columns("IDPersona").Value)
                If Not Persona.Load() Then
                    Set Persona = Nothing
                    tdbgrdData.SetFocus
                    Exit Sub
                End If
                
                SetLastPersona Persona.IDPersona, Persona.ApellidoNombre
                
                Set Feriado = New Feriado
                Feriado.VerificarReservasDelPasajero Persona.IDPersona
                Set Feriado = Nothing
                
                Screen.MousePointer = vbHourglass
                frmPersonaPrepago.LoadDataAndShow Persona
                Screen.MousePointer = vbDefault
                
                Set Persona = Nothing
            End If
    End Select
End Sub

Private Sub tabIndex_Click()
    FillData 0
    mSearchString = tabIndex.SelectedItem.Key
    mLastPressedKeyDelay = 0
    tmrKeyDelay.Enabled = False
End Sub

Private Sub cboFilterTipo_Click()
    FillData 0
End Sub

Private Sub cboFilterActivo_Click()
    FillData 0
End Sub

Private Sub tdbgrdData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        mSearchString = tabIndex.SelectedItem.Key
        mLastPressedKeyDelay = 0
        tmrKeyDelay.Enabled = False
    End If
End Sub

Private Sub tdbgrdData_DblClick()
    If GetFormIndex(FormWaitingForSelect) > 0 Then
        tlbMain_ButtonClick tlbMain.Buttons.Item("SELECT")
    Else
        tlbMain_ButtonClick tlbMain.Buttons.Item("PROPERTIES")
    End If
    DoEvents
End Sub

Private Sub tdbgrdData_SelChange(Cancel As Integer)
    If tdbgrdData.SelStartCol > 0 And tdbgrdData.SelEndCol > 0 Then
        'tdbgrdData.ColumnHeaders(tdbgrdData.SortKey + 1).Icon = 0
        If tdbgrdData.SelStartCol = mSortColumn Then
            mSortOrderAscending = Not mSortOrderAscending
        Else
            mSortColumn = tdbgrdData.SelStartCol
            mSortOrderAscending = True
        End If
        tdbgrdData.SelStartCol = 0
        tdbgrdData.SelEndCol = 0
        'ColumnHeader.Icon = tdbgrdData.SortOrder + 1
        
        FillData 0
    End If
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
    
    tabIndex.Top = CoolBarHeight + CONTROL_SPACE
    tabIndex.Width = ScaleWidth - (CONTROL_SPACE * 2)
    tabIndex.TabMinWidth = tabIndex.Width / 28
    
    tdbgrdData.Top = tabIndex.Top + tabIndex.Height
    tdbgrdData.Left = CONTROL_SPACE
    tdbgrdData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    tdbgrdData.Height = ScaleHeight - tdbgrdData.Top - CONTROL_SPACE - stbMain.Height
    
    tlbPin.Top = ScaleHeight - 330
    tlbPin.Left = 15
End Sub

Private Sub tmrKeyDelay_Timer()
    If mLastPressedKeyDelay > pParametro.Persona_Apellido_Busqueda_Delay_Milliseconds Then
        tmrKeyDelay.Enabled = False
        mLastPressedKeyDelay = 0
    Else
        mLastPressedKeyDelay = mLastPressedKeyDelay + tmrKeyDelay.Interval
    End If
End Sub
