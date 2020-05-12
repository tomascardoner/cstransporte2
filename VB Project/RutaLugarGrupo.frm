VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmRutaLugarGrupo 
   Caption         =   "Rutas - Grupos de Lugares"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RutaLugarGrupo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   9195
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   1111
      BandCount       =   2
      FixedOrder      =   -1  'True
      _CBWidth        =   9195
      _CBHeight       =   630
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain"
      MinHeight1      =   570
      Width1          =   6450
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picFilterRuta"
      MinWidth2       =   3015
      MinHeight2      =   360
      Width2          =   555
      Key2            =   "Ruta"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin VB.PictureBox picFilterRuta 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   6090
         ScaleHeight     =   360
         ScaleWidth      =   3015
         TabIndex        =   4
         Top             =   135
         Width           =   3015
         Begin VB.ComboBox cboRuta 
            Height          =   330
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   0
            Width           =   2550
         End
         Begin VB.Label lblRuta 
            AutoSize        =   -1  'True
            Caption         =   "Ruta:"
            Height          =   210
            Left            =   0
            TabIndex        =   6
            Top             =   60
            Width           =   375
         End
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   570
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   1005
         ButtonWidth     =   2170
         ButtonHeight    =   1005
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
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
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   5475
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   635
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
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
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Ruta"
         Text            =   "Ruta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Nombre"
         Text            =   "Grupo de Lugares"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Lugar Predeterminado"
         Object.Width           =   4410
      EndProperty
   End
End
Attribute VB_Name = "frmRutaLugarGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean

Public Function FillListView(ByVal IDRuta As String, ByVal IDLugarGrupo As Long) As Boolean
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    Dim SQL_Where As String
    
    If mLoading Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If IDRuta = "" Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & IDRuta & KEY_DELIMITER & IDLugarGrupo
    End If
    
    SQL_Where = ""
    
    If cboRuta.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "RutaLugarGrupo.IDRuta = '" & ReplaceQuote(cboRuta.Text) & "'"
    End If
    
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT RutaLugarGrupo.IDRuta, LugarGrupo.IDLugarGrupo, LugarGrupo.Nombre AS LugarGrupoNombre, Lugar.Nombre AS LugarPredeterminadoNombre FROM (RutaLugarGrupo INNER JOIN LugarGrupo ON RutaLugarGrupo.IDLugarGrupo = LugarGrupo.IDLugarGrupo) INNER JOIN Lugar ON RutaLugarGrupo.IDLugarPredeterminado = Lugar.IDLugar" & SQL_Where & " ORDER BY RutaLugarGrupo.IDRuta, LugarGrupo.Nombre"
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("IDRuta").Value & KEY_DELIMITER & .Fields("IDLugarGrupo").Value, RTrim(.Fields("IDRuta").Value))
                ListItem.SubItems(1) = .Fields("LugarGrupoNombre").Value
                ListItem.SubItems(2) = .Fields("LugarPredeterminadoNombre").Value
                .MoveNext
            Loop
            
            stbMain.SimpleText = .RecordCount & " items"
        Else
            stbMain.SimpleText = "No hay items."
        End If
        .Close
    End With
    Set recData = Nothing
    
    On Error Resume Next
    Set lvwData.SelectedItem = lvwData.ListItems(KeySave)
    lvwData.SelectedItem.EnsureVisible
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    
    FillListView = True
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.RutaLugarGrupo.FillListView", "Error al obtener la lista de Rutas-Grupo de Lugares."
End Function

Public Sub FillComboBoxRuta()
    Dim recRuta As ADODB.Recordset
    Dim IDRutaSave As String
    
    IDRutaSave = cboRuta.Text

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recRuta = New ADODB.Recordset
    recRuta.Source = "SELECT IDRuta FROM Ruta WHERE IDRuta <> '" & ReplaceQuote(pParametro.Ruta_ID_Otra) & "' AND Activo = 1" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & " ORDER BY IDRuta"
    recRuta.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    cboRuta.Clear
    cboRuta.AddItem ITEM_ALL_FEMALE
    Do While Not recRuta.EOF
        cboRuta.AddItem RTrim(recRuta("IDRuta").Value)
        recRuta.MoveNext
    Loop
    recRuta.Close
    Set recRuta = Nothing

    cboRuta.ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboRuta, IDRutaSave, cscpItemOrfirst)
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.RutaDetalle.FillComboBoxRuta", "Error al leer la lista de Rutas."
End Sub

Private Sub cboRuta_Click()
    FillListView "", 0
End Sub

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    ResizeControls NewHeight
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
    End If
End Sub

Private Sub Form_Load()
    mLoading = True
    
    lvwData.GridLines = pParametro.ListView_GridLines
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL TOOLBAR
    Set tlbMain.ImageList = frmMDI.ilsFormToolbar
    Set tlbMain.HotImageList = frmMDI.ilsFormToolbarHot
    tlbMain.Buttons("NEW").Image = "NEW"
    tlbMain.Buttons("PROPERTIES").Image = "PROPERTIES"
    tlbMain.Buttons("DELETE").Image = "DELETE"
    '//////////////////////////////////////////////////////////
    
    cbrMain.Bands("Toolbar").MinWidth = CSM_Control_Toolbar.GetTotalWidth(tlbMain)
    
    FillComboBoxRuta
    cboRuta.ListIndex = 0
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "RutaLugarGrupo", cbrMain
    pParametro.GetListViewSettings "RutaLugarGrupo", lvwData
    
    mLoading = False
    
    FillListView "", 0
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "RutaLugarGrupo", cbrMain
    pParametro.SaveListViewSettings "RutaLugarGrupo", lvwData
    Set frmRutaLugarGrupo = Nothing
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
    Dim RutaLugarGrupo As RutaLugarGrupo
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_RUTALUGARGRUPO_ADD) Then
                Screen.MousePointer = vbHourglass
                Set RutaLugarGrupo = New RutaLugarGrupo
                frmRutaLugarGrupoPropiedad.LoadDataAndShow Me, RutaLugarGrupo
                If cboRuta.ListIndex > 0 Then
                    frmRutaLugarGrupoPropiedad.datcboRuta.BoundText = cboRuta.Text
                    frmRutaLugarGrupoPropiedad.datcboLugarGrupo.SetFocus
                End If
                Set RutaLugarGrupo = Nothing
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_RUTALUGARGRUPO_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Set RutaLugarGrupo = New RutaLugarGrupo
                RutaLugarGrupo.IDRuta = GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)
                RutaLugarGrupo.IDLugarGrupo = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                If Not RutaLugarGrupo.Load() Then
                    Set RutaLugarGrupo = Nothing
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                frmRutaLugarGrupoPropiedad.LoadDataAndShow Me, RutaLugarGrupo
                Screen.MousePointer = vbDefault
                
                Set RutaLugarGrupo = Nothing
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_RUTALUGARGRUPO_DELETE) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If MsgBox("¿Desea eliminar la Ruta-Grupo de Lugar seleccionada?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                    Set RutaLugarGrupo = New RutaLugarGrupo
                    RutaLugarGrupo.IDRuta = GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)
                    RutaLugarGrupo.IDLugarGrupo = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                    If Not RutaLugarGrupo.Load() Then
                        Set RutaLugarGrupo = Nothing
                        lvwData.SetFocus
                        Exit Sub
                    End If
                    If Not RutaLugarGrupo.Delete() Then
                        Set RutaLugarGrupo = Nothing
                        lvwData.SetFocus
                        Exit Sub
                    End If
                End If
            End If
    End Select
    
    Set RutaLugarGrupo = Nothing
End Sub

Private Sub ResizeControls(ByVal CoolBarHeight As Single)
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
    
    lvwData.Top = CoolBarHeight + CONTROL_SPACE
    lvwData.Left = CONTROL_SPACE
    lvwData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lvwData.Height = ScaleHeight - lvwData.Top - CONTROL_SPACE - stbMain.Height
End Sub
