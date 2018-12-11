VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmViajeDetalleAsistenciaMultiple 
   Caption         =   "Asistencia Múltiple"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "ViajeDetalleAsistenciaMultiple.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   8205
   Begin VB.ComboBox cboCaja 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3510
   End
   Begin TrueOleDBGrid80.TDBGrid tdbgrdData 
      Height          =   4635
      Left            =   120
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
      Columns(0).Caption=   "Indice"
      Columns(0).DataField=   "Indice"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Orden"
      Columns(1).DataField=   "Orden"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Pasajero"
      Columns(2).DataField=   "Pasajero"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Importe"
      Columns(3).DataField=   "Importe"
      Columns(3).NumberFormat=   "Currency"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Pagado"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Currency"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Realizado"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8208"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1058"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=979"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8705"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=6562"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=6482"
      Splits(0)._ColumnProps(17)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=8720"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=1773"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1693"
      Splits(0)._ColumnProps(23)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=8706"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=1773"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1693"
      Splits(0)._ColumnProps(29)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=514"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=1561"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1482"
      Splits(0)._ColumnProps(35)=   "Column(5).AllowSizing=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
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
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
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
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=43,.alignment=2,.locked=-1"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=58,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(46)  =   ":id=58,.locked=-1"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=62,.parent=43,.alignment=1,.locked=-1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=44,.alignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=45"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=47"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=1"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44,.alignment=2"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=16,.parent=43,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=44,.alignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=45"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=47"
      _StyleDefs(62)  =   "Named:id=33:Normal"
      _StyleDefs(63)  =   ":id=33,.parent=0"
      _StyleDefs(64)  =   "Named:id=34:Heading"
      _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(66)  =   ":id=34,.wraptext=-1"
      _StyleDefs(67)  =   "Named:id=35:Footing"
      _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(69)  =   "Named:id=36:Selected"
      _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(71)  =   "Named:id=37:Caption"
      _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(73)  =   "Named:id=38:HighlightRow"
      _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=39:EvenRow"
      _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(77)  =   "Named:id=40:OddRow"
      _StyleDefs(78)  =   ":id=40,.parent=33"
      _StyleDefs(79)  =   "Named:id=41:RecordSelector"
      _StyleDefs(80)  =   ":id=41,.parent=34"
      _StyleDefs(81)  =   "Named:id=42:FilterBar"
      _StyleDefs(82)  =   ":id=42,.parent=33"
   End
   Begin VB.Label lblCaja 
      AutoSize        =   -1  'True
      Caption         =   "Caja:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   360
   End
End
Attribute VB_Name = "frmViajeDetalleAsistenciaMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean
Private mViaje As Viaje

Private mData() As Variant
Private mUpdating As Boolean
Private mCloned As Boolean
Private mClonedRecordset As ADODB.Recordset

Private mKeyDecimal As Boolean

Public Sub LoadDataAndShow(ByRef Viaje As Viaje)
    Set mViaje = Viaje

    mLoading = True
    
    Load Me
        
    mLoading = False
    
    If Not FillListView(mViaje.FechaHora, mViaje.IDRuta) Then
        Unload Me
        Exit Sub
    End If

    Caption = "Asistencia al Viaje: " & Viaje.FechaHora_WeekdayName & " " & Viaje.FechaHora_Formatted & " - " & Viaje.Ruta_DisplayName

    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
    On Error Resume Next
    tdbgrdData.Col = 4
End Sub

Public Function FillListView(ByVal FechaHora As Date, ByVal IDRuta As String) As Boolean
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    
    If cboCaja.ListIndex = -1 Then
        MsgBox "No hay ninguna Caja para marcar Asistencias.", vbInformation, App.Title
        Exit Function
    End If
    
    If FechaHora <> mViaje.FechaHora Or IDRuta <> mViaje.IDRuta Then
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
    cmdData.CommandText = "sp_ViajeDetalle_Asistencia"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("FechaHora_FILTER", adDate, adParamInput, , mViaje.FechaHora)
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta_FILTER", adChar, adParamInput, 20, mViaje.IDRuta)
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenStatic, adLockReadOnly
    Set cmdData = Nothing

    mCloned = False
    mUpdating = False
    If Not recData.EOF Then
        mData = recData.GetRows(recData.RecordCount)
    End If
    Set tdbgrdData.DataSource = recData
    tdbgrdData.Columns("Pagado").Locked = False
    tdbgrdData.Columns("Realizado").Locked = False
    
    Set recData = Nothing
    
    Screen.MousePointer = vbDefault
    FillListView = True
    Exit Function

ErrorHandler:
    ShowErrorMessage "Forms.ViajeDetalleAsistencia.FillListView", "Error al obtener el Detalle del Viaje." & vbCr & vbCr & "Fecha/Hora: " & mViaje.FechaHora_Formatted & vbCr & "IDRuta: " & mViaje.IDRuta
End Function

Private Sub cboCaja_Click()
    FillListView mViaje.FechaHora, mViaje.IDRuta
End Sub

Private Sub Form_Load()
    Dim CuentaCorrienteCaja As CuentaCorrienteCaja
    
    mLoading = True
        
    CSM_Forms.ResizeAndPosition frmMDI, Me
    
    tdbgrdData.EvenRowStyle.BackColor = pParametro.GridEvenRowBackColor
    tdbgrdData.EvenRowStyle.ForeColor = pParametro.GridEvenRowForeColor
    tdbgrdData.OddRowStyle.BackColor = pParametro.GridOddRowBackColor
    tdbgrdData.OddRowStyle.ForeColor = pParametro.GridOddRowForeColor
    
    FillComboBoxCuentaCorrienteCaja
    
    If mViaje.IDConductor > 0 Then
        Set CuentaCorrienteCaja = New CuentaCorrienteCaja
        CuentaCorrienteCaja.IDPersona = mViaje.IDConductor
        If CuentaCorrienteCaja.LoadByPersona() Then
            cboCaja.ListIndex = CSM_Control_ComboBox.GetListIndexByItemData(cboCaja, CuentaCorrienteCaja.IDCuentaCorrienteCaja, cscpItemOrfirst)
        End If
    End If

    mLoading = False
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tdbgrdData.Update
    Set mViaje = Nothing
    Set frmViajeDetalleAsistenciaMultiple = Nothing
End Sub

Private Sub ResizeControls()
    Const CONTROL_SPACE = 60
    Const SCROLLBAR_WIDTH = 450
    
    On Error Resume Next
    
    tdbgrdData.Top = cboCaja.Top + cboCaja.Height + CONTROL_SPACE
    tdbgrdData.Left = CONTROL_SPACE
    tdbgrdData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    tdbgrdData.Height = ScaleHeight - tdbgrdData.Top - CONTROL_SPACE
    
    tdbgrdData.Columns("Pasajero").Width = tdbgrdData.Width - SCROLLBAR_WIDTH - tdbgrdData.Columns("Orden").Width - tdbgrdData.Columns("Importe").Width - tdbgrdData.Columns("Pagado").Width - tdbgrdData.Columns("Realizado").Width
End Sub

Public Sub FillComboBoxCuentaCorrienteCaja()
    Dim KeySave As Long
    
    If cboCaja.ListCount > 0 Then
        KeySave = cboCaja.ItemData(cboCaja.ListIndex)
    End If

    If pCPermiso.GotPermission(PERMISO_CUENTA_CORRIENTE_CAJA_VIEW_ALL, False) Then
        If Not CSM_Control_ComboBox.FillFromSQL(cboCaja, "SELECT IDCuentaCorrienteCaja, Nombre FROM CuentaCorrienteCaja WHERE Activo = 1 ORDER BY Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpItemOrNone, KeySave) Then
            Unload Me
            Exit Sub
        End If
    Else
        If Not CSM_Control_ComboBox.FillFromSQL(cboCaja, "SELECT CuentaCorrienteCaja.IDCuentaCorrienteCaja, CuentaCorrienteCaja.Nombre FROM CuentaCorrienteCaja LEFT JOIN Persona ON CuentaCorrienteCaja.IDPersona = Persona.IDPersona WHERE CuentaCorrienteCaja.Activo = 1 AND (CuentaCorrienteCaja.MostrarSiempre = 1 OR CuentaCorrienteCaja.IDCuentaCorrienteCaja = " & pUsuario.IDCuentaCorrienteCaja & " OR Persona.IDPersona = " & mViaje.IDConductor & " OR Persona.IDPersona = " & mViaje.IDConductor2 & ") ORDER BY CuentaCorrienteCaja.Nombre", "IDCuentaCorrienteCaja", "Nombre", "Cajas", cscpFirst, KeySave) Then
            Unload Me
            Exit Sub
        End If
    End If
End Sub

Private Sub tdbgrdData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub tdbgrdData_KeyPress(KeyAscii As Integer)
    If tdbgrdData.Col = tdbgrdData.Columns("Pagado").ColIndex Then
        If KeyAscii <> vbKeyEscape And ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
            KeyAscii = 0
        End If
        If mKeyDecimal Then
            KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
            mKeyDecimal = False
        End If
        If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
            If InStr(1, tdbgrdData.Columns("Pagado").Value, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
                KeyAscii = 0
            End If
        End If
        If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
            If InStr(1, tdbgrdData.Columns("Pagado").Value, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
                KeyAscii = 0
            End If
        End If
    Else
        If KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeySpace And KeyAscii <> vbKeyEscape And Asc(UCase(Chr(KeyAscii))) <> vbKeyS And Asc(UCase(Chr(KeyAscii))) <> vbKeyN Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub tdbgrdData_UnboundColumnFetch(Bookmark As Variant, ByVal Col As Integer, Value As Variant)
    Dim recData As ADODB.Recordset
    
    Select Case Col
        Case 4
            If Not mCloned Then
                Set recData = tdbgrdData.DataSource
                Set mClonedRecordset = recData.Clone
                Set recData = Nothing
                mCloned = True
            End If
            On Error Resume Next
            mClonedRecordset.Bookmark = Bookmark
            Value = mData(4, mClonedRecordset.AbsolutePosition - 1)
        Case 5
            If Not mCloned Then
                Set recData = tdbgrdData.DataSource
                Set mClonedRecordset = recData.Clone
                Set recData = Nothing
                mCloned = True
            End If
            On Error Resume Next
            mClonedRecordset.Bookmark = Bookmark
            Value = IIf(IsNull(mData(5, mClonedRecordset.AbsolutePosition - 1)), "", IIf(mData(5, mClonedRecordset.AbsolutePosition - 1) = 1, "S", "N"))
    End Select
End Sub

Private Sub tdbgrdData_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If Not mUpdating Then
        Select Case ColIndex
            Case 4
                If Trim(tdbgrdData.Columns("Pagado").Value) <> "" Then
                    If Not IsNumeric(tdbgrdData.Columns("Pagado").Value) Then
                        MsgBox "El Importe Pagado debe ser un valor numérico.", vbExclamation, App.Title
                        Cancel = True
                        tdbgrdData.SetFocus
                    End If
                End If
                If OldValue > 0 And OldValue <> tdbgrdData.Columns("Pagado").Value Then
                    MsgBox "Esta Reserva ya tenía un Pago, por lo tanto, no puede modificarse.", vbExclamation, App.Title
                    Cancel = True
                    tdbgrdData.SetFocus
                End If
            Case 5
                If tdbgrdData.Columns("Realizado").Value <> "" And UCase(tdbgrdData.Columns("Realizado").Value) <> "S" And UCase(tdbgrdData.Columns("Realizado").Value) <> "N" Then
                    MsgBox "Realizado solo pueder ser Vacío, S=Sí o N=No.", vbExclamation, App.Title
                    Cancel = True
                    tdbgrdData.SetFocus
                End If
        End Select
    End If
End Sub

Private Sub tdbgrdData_AfterColUpdate(ByVal ColIndex As Integer)
    Dim ViajeDetalle As ViajeDetalle
    Dim ImporteSave As Currency
    Dim RealizadoSave As Long
    Dim recData As ADODB.Recordset
    
    If Not mUpdating Then
        Select Case ColIndex
            Case 4
                Set ViajeDetalle = New ViajeDetalle
                ViajeDetalle.FechaHora = mViaje.FechaHora
                ViajeDetalle.IDRuta = mViaje.IDRuta
                ViajeDetalle.Indice = tdbgrdData.Columns("Indice").Value
                If ViajeDetalle.Load() Then
                    ImporteSave = ViajeDetalle.ImporteContado
                    RealizadoSave = ViajeDetalle.Realizado
                    ViajeDetalle.Realizado = 1
                    If Trim(tdbgrdData.Columns("Pagado").Value) = "" Then
                        ViajeDetalle.ImporteContado = 0
                        ViajeDetalle.IDMedioPago = 0
                        ViajeDetalle.Cuotas = 0
                        ViajeDetalle.Operacion = ""
                        ViajeDetalle.IDCuentaCorrienteCaja = 0
                    Else
                        'PAGO CONTADO
                        ViajeDetalle.ImporteContado = CCur(tdbgrdData.Columns("Pagado").Value)
                        If ViajeDetalle.IDMedioPago = 0 Then
                            ViajeDetalle.IDMedioPago = pParametro.MedioPago_Predeterminado_ID
                            ViajeDetalle.Cuotas = 1
                            ViajeDetalle.Operacion = ""
                        End If
                        If ViajeDetalle.ImporteCuentaCorriente > 0 And ViajeDetalle.ImporteCuentaCorriente + ViajeDetalle.ImporteContado > ViajeDetalle.Importe Then
                            If ViajeDetalle.Importe - ViajeDetalle.ImporteContado > 0 Then
                                ViajeDetalle.ImporteCuentaCorriente = ViajeDetalle.Importe - ViajeDetalle.ImporteContado
                            Else
                                ViajeDetalle.ImporteCuentaCorriente = 0
                            End If
                        End If
                        ViajeDetalle.IDCuentaCorrienteCaja = cboCaja.ItemData(cboCaja.ListIndex)
                    End If
                    If ViajeDetalle.Realizar() Then
                        Set recData = tdbgrdData.DataSource
                        mData(4, recData.AbsolutePosition - 1) = ViajeDetalle.ImporteContado
                        mData(5, recData.AbsolutePosition - 1) = ViajeDetalle.Realizado
                        Set recData = Nothing
                    Else
                        mUpdating = True
                        tdbgrdData.Columns("Pagado").Value = ImporteSave
                        tdbgrdData.Columns("Realizado").Value = IIf(RealizadoSave = 0, "", IIf(RealizadoSave, "S", "N"))
                        mUpdating = False
                    End If
                    tdbgrdData.Columns("Pagado").Locked = False
                    tdbgrdData.Columns("Realizado").Locked = False
                End If
                Set ViajeDetalle = Nothing
            Case 5
                Set ViajeDetalle = New ViajeDetalle
                ViajeDetalle.FechaHora = mViaje.FechaHora
                ViajeDetalle.IDRuta = mViaje.IDRuta
                ViajeDetalle.Indice = tdbgrdData.Columns("Indice").Value
                If ViajeDetalle.Load() Then
                    ViajeDetalle.Realizado = IIf(tdbgrdData.Columns("Realizado").Value = "", 0, IIf(UCase(tdbgrdData.Columns("Realizado").Value) = "S", 1, 2))
                    If ViajeDetalle.Realizado = 2 Then
                        Call ViajeDetalle_ShowViajeVuelta(ViajeDetalle, "Hay %1 Reserva(s) sin Asistencia para este Pasajero en el mismo Día.")
                        tdbgrdData.SetFocus
                    End If
                    ViajeDetalle.ForzarDebito = IIf(ViajeDetalle.Realizado = 2, True, False)
                    If ViajeDetalle.Realizar() Then
                        Set recData = tdbgrdData.DataSource
                        mData(5, recData.AbsolutePosition - 1) = IIf(ViajeDetalle.Realizado = 0, Null, ViajeDetalle.Realizado)
                        Set recData = Nothing
                    Else
                        mUpdating = True
                        tdbgrdData.Columns("Realizado").Value = RealizadoSave
                        mUpdating = False
                    End If
                    tdbgrdData.Columns("Pagado").Locked = False
                    tdbgrdData.Columns("Realizado").Locked = False
                End If
                Set ViajeDetalle = Nothing
        End Select
    End If
End Sub

