VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmListaPrecioDetalle 
   Caption         =   "Detalle de Lista de Precios"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ListaPrecioDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   8820
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   1111
      FixedOrder      =   -1  'True
      _CBWidth        =   8820
      _CBHeight       =   630
      _Version        =   "6.7.9782"
      Child1          =   "picRuta"
      MinWidth1       =   3855
      MinHeight1      =   435
      Width1          =   3855
      Key1            =   "RUTA"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picTipo"
      MinWidth2       =   2775
      MinHeight2      =   435
      Width2          =   2775
      Key2            =   "TIPO"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Child3          =   "tlbMain"
      MinWidth3       =   795
      MinHeight3      =   570
      Width3          =   795
      Key3            =   "MAIN"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   570
         Left            =   7935
         TabIndex        =   9
         Top             =   30
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1005
         ButtonWidth     =   1402
         ButtonHeight    =   1005
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "PRINT"
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox picTipo 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   4935
         ScaleHeight     =   435
         ScaleWidth      =   2775
         TabIndex        =   4
         Top             =   90
         Width           =   2775
         Begin VB.OptionButton optTipoComision 
            Caption         =   "Comisión"
            Height          =   210
            Left            =   540
            TabIndex        =   6
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton optTipoPasajero 
            Caption         =   "Pasajero"
            Height          =   210
            Left            =   1740
            TabIndex        =   5
            Top             =   120
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.Label lblTipo 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   210
            Left            =   60
            TabIndex        =   7
            Top             =   120
            Width           =   345
         End
      End
      Begin VB.PictureBox picRuta 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   30
         ScaleHeight     =   435
         ScaleWidth      =   4680
         TabIndex        =   2
         Top             =   90
         Width           =   4680
         Begin MSDataListLib.DataCombo datcboRuta 
            Height          =   330
            Left            =   540
            TabIndex        =   8
            Top             =   60
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
         Begin VB.Label lblRuta 
            AutoSize        =   -1  'True
            Caption         =   "Ruta:"
            Height          =   210
            Left            =   60
            TabIndex        =   3
            Top             =   120
            Width           =   375
         End
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbgrdData 
      Height          =   4635
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8176
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "IDLugarGrupoOrigen"
      Columns(0).DataField=   "IDLugarGrupoOrigen"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Origen"
      Columns(1).DataField=   "LugarGrupoOrigen"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "IDLugarGrupoDestino"
      Columns(2).DataField=   "IDLugarGrupoDestino"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Destino"
      Columns(3).DataField=   "LugarGrupoDestino"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Importe"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Currency"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8208"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=5292"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=5212"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8720"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(17)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=8208"
      Splits(0)._ColumnProps(19)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(21)=   "Column(3).Width=5292"
      Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=5212"
      Splits(0)._ColumnProps(24)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=8720"
      Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(27)=   "Column(4).Width=1773"
      Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=1693"
      Splits(0)._ColumnProps(30)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._ColStyle=514"
      Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      TabAction       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
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
      _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(37)  =   ":id=28,.locked=-1"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(42)  =   ":id=32,.locked=-1"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(47)  =   ":id=58,.locked=-1"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(52)  =   ":id=62,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
      _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
      _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(60)  =   "Named:id=33:Normal"
      _StyleDefs(61)  =   ":id=33,.parent=0"
      _StyleDefs(62)  =   "Named:id=34:Heading"
      _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   ":id=34,.wraptext=-1"
      _StyleDefs(65)  =   "Named:id=35:Footing"
      _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   "Named:id=36:Selected"
      _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=37:Caption"
      _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(71)  =   "Named:id=38:HighlightRow"
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmListaPrecioDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean
Private mIDListaPrecio As Long

Private mData() As Variant
Private mUpdating As Boolean
Private mCloned As Boolean
Private mClonedRecordset As ADODB.Recordset

Private mKeyDecimal As Boolean

Public Sub LoadDataAndShow(ByVal IDListaPrecio As Long)
    Dim ListaPrecio As ListaPrecio
    
    mLoading = True
    
    mIDListaPrecio = IDListaPrecio
    
    Load Me
    
    If Not CSM_Control_DataCombo.FillFromSQL(datcboRuta, "SELECT RTRIM(IDRuta) AS IDRuta FROM Ruta WHERE IDRuta <> '" & ReplaceQuote(pParametro.Ruta_ID_Otra) & "' AND Activo = 1" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & " ORDER BY IDRuta", "IDRuta", "IDRuta", "Rutas", cscpFirst) Then
        Unload Me
        Exit Sub
    End If
    
    mLoading = False
    
    If Not FillListView(mIDListaPrecio) Then
        Unload Me
        Exit Sub
    End If

    Set ListaPrecio = New ListaPrecio
    ListaPrecio.IDListaPrecio = mIDListaPrecio
    If ListaPrecio.Load() Then
        Caption = "Detalle de la Lista de Precios: " & ListaPrecio.Nombre
    End If
    Set ListaPrecio = Nothing

    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Sub ForceRefresh()
    FillListView mIDListaPrecio
End Sub

Public Function FillListView(ByVal IDListaPrecio As Long) As Boolean
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    If mIDListaPrecio <> IDListaPrecio Then
        Exit Function
    End If

    If mLoading Then
        Exit Function
    End If

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Screen.MousePointer = vbHourglass

    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "sp_ListaPrecioDetalle_DataGrid_Complete"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDListaPrecio_FILTER", adInteger, adParamInput, , mIDListaPrecio)
    cmdData.Parameters.Append cmdData.CreateParameter("OcupanteTipo_FILTER", adChar, adParamInput, 2, IIf(optTipoComision.Value, OCUPANTE_TIPO_COMISION, OCUPANTE_TIPO_PASAJERO))
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, datcboRuta.BoundText)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenStatic, adLockReadOnly
    Set cmdData = Nothing

    mCloned = False
    mUpdating = False
    If Not recData.EOF Then
        mData = recData.GetRows(recData.RecordCount)
    End If
    Set tdbgrdData.DataSource = recData
    tdbgrdData.Columns("Importe").Locked = False
    
    Set recData = Nothing
    
    Screen.MousePointer = vbDefault
    FillListView = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Forms.ListaPrecioDetalle.FillListView", "Error al obtener el Detalle de la Lista de Precios." & vbCr & vbCr & "IDListaPrecio: " & mIDListaPrecio
End Function

Private Sub datcboRuta_Change()
    FillListView mIDListaPrecio
End Sub

Private Sub datcboRuta_Click(Area As Integer)
    tdbgrdData.Update
End Sub

Private Sub tdbgrdData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub tdbgrdData_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyEscape And ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        KeyAscii = 0
    End If
    If mKeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        mKeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, tdbgrdData.Columns("Importe").Value, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, tdbgrdData.Columns("Importe").Value, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub tdbgrdData_UnboundColumnFetch(Bookmark As Variant, ByVal Col As Integer, Value As Variant)
    Dim recData As ADODB.Recordset
    
    If Col = 4 Then
        If Not mCloned Then
            Set recData = tdbgrdData.DataSource
            Set mClonedRecordset = recData.Clone
            Set recData = Nothing
            mCloned = True
        End If
        On Error Resume Next
        mClonedRecordset.Bookmark = Bookmark
        Value = mData(4, mClonedRecordset.AbsolutePosition - 1)
    End If
End Sub

Private Sub tdbgrdData_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If Not mUpdating Then
        If Trim(tdbgrdData.Columns("Importe").Value) <> "" Then
            If Not IsNumeric(tdbgrdData.Columns("Importe").Value) Then
                MsgBox "El Importe debe ser un valor numérico.", vbExclamation, App.Title
                mUpdating = True
                tdbgrdData.Columns("Importe").Value = OldValue
                mUpdating = True
                tdbgrdData.SetFocus
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub tdbgrdData_AfterColUpdate(ByVal ColIndex As Integer)
    Dim ListaPrecioDetalle As ListaPrecioDetalle
    Dim ImporteSave As Currency
    Dim recData As ADODB.Recordset
    
    If Not mUpdating Then
        Set ListaPrecioDetalle = New ListaPrecioDetalle
        ListaPrecioDetalle.NoMatchRaiseError = False
        ListaPrecioDetalle.IDListaPrecio = mIDListaPrecio
        ListaPrecioDetalle.OcupanteTipo = IIf(optTipoComision.Value, OCUPANTE_TIPO_COMISION, OCUPANTE_TIPO_PASAJERO)
        ListaPrecioDetalle.IDRuta = datcboRuta.BoundText
        ListaPrecioDetalle.IDLugarGrupoOrigen = tdbgrdData.Columns("IDLugarGrupoOrigen").Value
        ListaPrecioDetalle.IDLugarGrupoDestino = tdbgrdData.Columns("IDLugarGrupoDestino").Value
        If ListaPrecioDetalle.Load() Then
            If Trim(tdbgrdData.Columns("Importe").Value) = "" Then
                'DELETE
                If ListaPrecioDetalle.NoMatch Then
                    Set recData = tdbgrdData.DataSource
                    mData(4, recData.AbsolutePosition - 1) = ""
                    Set recData = Nothing
                Else
                    If ListaPrecioDetalle.Delete() Then
                        Set recData = tdbgrdData.DataSource
                        mData(4, recData.AbsolutePosition - 1) = ""
                        Set recData = Nothing
                    Else
                        mUpdating = True
                        tdbgrdData.Columns("Importe").Value = ListaPrecioDetalle.Importe
                        mUpdating = False
                    End If
                End If
            Else
                ImporteSave = ListaPrecioDetalle.Importe
                ListaPrecioDetalle.Importe = CCur(tdbgrdData.Columns("Importe").Value)
                If ListaPrecioDetalle.NoMatch Then
                    'ADD
                    If ListaPrecioDetalle.AddNew() Then
                        Set recData = tdbgrdData.DataSource
                        mData(4, recData.AbsolutePosition - 1) = ListaPrecioDetalle.Importe
                        Set recData = Nothing
                    Else
                        mUpdating = True
                        tdbgrdData.Columns("Importe").Value = ""
                        mUpdating = False
                    End If
                Else
                    'UPDATE
                    If ListaPrecioDetalle.Update() Then
                        Set recData = tdbgrdData.DataSource
                        mData(4, recData.AbsolutePosition - 1) = ListaPrecioDetalle.Importe
                        Set recData = Nothing
                    Else
                        mUpdating = True
                        tdbgrdData.Columns("Importe").Value = ImporteSave
                        mUpdating = False
                    End If
                End If
            End If
            tdbgrdData.Columns("Importe").Locked = False
        End If
        Set ListaPrecioDetalle = Nothing
    End If
End Sub

Public Sub FillComboBoxRuta()
    Dim KeySave As String
    Dim recData As ADODB.Recordset
    
    KeySave = datcboRuta.BoundText
    Set recData = datcboRuta.RowSource
    recData.Requery
    Set recData = Nothing
    datcboRuta.BoundText = KeySave
End Sub

Private Sub Form_Load()
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL TOOLBAR
    Set tlbMain.ImageList = frmMDI.ilsFormToolbar
    Set tlbMain.HotImageList = frmMDI.ilsFormToolbarHot
    tlbMain.Buttons("PRINT").Image = "PRINT"
    '//////////////////////////////////////////////////////////
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "ListaPrecioDetalle", cbrMain
    
    tdbgrdData.EvenRowStyle.BackColor = pParametro.GridEvenRowBackColor
    tdbgrdData.EvenRowStyle.ForeColor = pParametro.GridEvenRowForeColor
    tdbgrdData.OddRowStyle.BackColor = pParametro.GridOddRowBackColor
    tdbgrdData.OddRowStyle.ForeColor = pParametro.GridOddRowForeColor
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pParametro.SaveCoolBarSettings "ListaPrecioDetalle", cbrMain
End Sub

Private Sub optTipoComision_Click()
    FillListView mIDListaPrecio
End Sub

Private Sub optTipoPasajero_Click()
    FillListView mIDListaPrecio
End Sub

Private Sub ResizeControls(ByVal CoolBarHeight As Single)
    Const CONTROL_SPACE = 60
    Const SCROLLBAR_WIDTH = 450
    
    On Error Resume Next
    
    tdbgrdData.Top = CoolBarHeight + CONTROL_SPACE
    tdbgrdData.Left = CONTROL_SPACE
    tdbgrdData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    tdbgrdData.Height = ScaleHeight - tdbgrdData.Top - CONTROL_SPACE
    
    tdbgrdData.Columns("LugarGrupoOrigen").Width = (tdbgrdData.Width - SCROLLBAR_WIDTH - tdbgrdData.Columns("Importe").Width) / 2
    tdbgrdData.Columns("LugarGrupoDestino").Width = tdbgrdData.Columns("LugarGrupoOrigen").Width
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PRINT"
            tdbgrdData.Columns("LugarGrupoOrigen").Width = 4000
            tdbgrdData.Columns("LugarGrupoDestino").Width = 4000
        
            With tdbgrdData.PrintInfo
                .PageHeader = Caption & " - " & datcboRuta.Text & " - " & IIf(optTipoPasajero.Value, "Pasajeros", "Comisiones")
                .PageHeaderFont.Name = "Arial"
                .PageHeaderFont.Size = 12
                .PageHeaderFont.Bold = True
                .PageFooter = Format(Date, "Short Date") & "   -   " & pParametro.CompanyName
                .SettingsPaperSize = vbPRPSLegal
                .SettingsMarginTop = 1000
                .SettingsMarginLeft = 1500
                .PreviewCaption = Caption
                .PreviewMaximize = True
                .PreviewInitZoom = 100
                .PrintPreview
            End With
            
            ResizeControls cbrMain.Height
    End Select
End Sub
