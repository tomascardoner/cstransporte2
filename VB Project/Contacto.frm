VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmContacto 
   Caption         =   "Contactos"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12990
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Contacto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   12990
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   60
      TabIndex        =   7
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
      Height          =   630
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12990
      _ExtentX        =   22913
      _ExtentY        =   1111
      FixedOrder      =   -1  'True
      _CBWidth        =   12990
      _CBHeight       =   630
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain"
      MinWidth1       =   5505
      MinHeight1      =   570
      Width1          =   5505
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picFilterGrupo"
      MinWidth2       =   3225
      MinHeight2      =   360
      Width2          =   3225
      Key2            =   "FilterGrupo"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Child3          =   "picFilterActivo"
      MinWidth3       =   1605
      MinHeight3      =   330
      Width3          =   1830
      Key3            =   "FilterActivo"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Begin VB.PictureBox picFilterGrupo 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   7815
         ScaleHeight     =   360
         ScaleWidth      =   3225
         TabIndex        =   8
         Top             =   135
         Width           =   3225
         Begin VB.ComboBox cboFilterGrupo 
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   0
            Width           =   2610
         End
         Begin VB.Label lblGrupo 
            AutoSize        =   -1  'True
            Caption         =   "Grupo:"
            Height          =   210
            Left            =   0
            TabIndex        =   10
            Top             =   60
            Width           =   495
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
         Left            =   11265
         ScaleHeight     =   330
         ScaleWidth      =   1635
         TabIndex        =   5
         Top             =   150
         Width           =   1635
         Begin VB.ComboBox cboFilterActivo 
            Height          =   330
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   0
            Width           =   990
         End
         Begin VB.Label lblFilterActivo 
            AutoSize        =   -1  'True
            Caption         =   "Activo:"
            Height          =   210
            Left            =   0
            TabIndex        =   6
            Top             =   60
            Width           =   510
         End
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   570
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   7560
         _ExtentX        =   13335
         _ExtentY        =   1005
         ButtonWidth     =   2170
         ButtonHeight    =   1005
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
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
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   5445
      Width           =   12990
      _ExtentX        =   22913
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
            Object.Width           =   21696
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
      Height          =   3735
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6588
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Apellido"
         Text            =   "Apellido"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Nombre"
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Compania"
         Text            =   "Compañía"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Grupo"
         Text            =   "Grupo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Activo"
         Text            =   "Activo"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmContacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean

Public FormWaitingForSelect As String
Public FormKeepOpenOnSelect As Boolean

Public Sub FillListView(ByVal IDContacto As Long)
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    
    If mLoading Then
        Exit Sub
    End If
        
    Screen.MousePointer = vbHourglass
    
    If IDContacto = 0 Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & IDContacto
    End If
    
    SQL_Where = ""
    
    If cboFilterGrupo.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Contacto.IDContactoGrupo = " & cboFilterGrupo.ItemData(cboFilterGrupo.ListIndex)
    End If
    
    If cboFilterActivo.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Contacto.Activo = " & IIf(cboFilterActivo.ListIndex = 1, 1, 0)
    End If
    
    Select Case lvwData.SortKey
        Case 0  'APELLIDO
            SQL_OrderBy = " ORDER BY Contacto.Apellido" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Contacto.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Contacto.Compania" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 1  'NOMBRE
            SQL_OrderBy = " ORDER BY Contacto.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Contacto.Apellido" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Contacto.Compania" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 2  'COMPANIA
            SQL_OrderBy = " ORDER BY Contacto.Compania" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Contacto.Apellido" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Contacto.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 3  'GRUPO
            SQL_OrderBy = " ORDER BY ContactoGrupo.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Contacto.Apellido" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Contacto.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Contacto.Compania" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 4  'ACTIVO
            SQL_OrderBy = " ORDER BY Contacto.Activo" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Contacto.Apellido" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Contacto.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", Contacto.Compania" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT Contacto.IDContacto, Contacto.Apellido, Contacto.Nombre, Contacto.Compania, ContactoGrupo.Nombre AS Grupo, Contacto.Activo FROM Contacto INNER JOIN ContactoGrupo ON Contacto.IDContactoGrupo = ContactoGrupo.IDContactoGrupo" & SQL_Where & SQL_OrderBy
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("IDContacto").Value, .Fields("Apellido").Value & "")
                ListItem.SubItems(1) = .Fields("Nombre").Value & ""
                ListItem.SubItems(2) = .Fields("Compania").Value & ""
                ListItem.SubItems(3) = .Fields("Grupo").Value
                ListItem.SubItems(4) = GetBooleanString(.Fields("Activo").Value)
                .MoveNext
            Loop
            
            stbMain.Panels("TEXT").Text = .RecordCount & " items"
        Else
            stbMain.Panels("TEXT").Text = "No hay items."
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
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.Contacto.FillListView", "Error al leer la Lista de Contactos."
End Sub

Private Sub cboFilterGrupo_Click()
    FillListView 0
End Sub

Private Sub cboFilterActivo_Click()
    FillListView 0
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
    tlbMain.Buttons("SELECT").Image = "SELECT"
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL LISTVIEW
    Set lvwData.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL PIN
    Set tlbPin.ImageList = frmMDI.ilsFormPin
    '//////////////////////////////////////////////////////////
    
    FillComboBoxContactoGrupo
    
    cboFilterActivo.AddItem ITEM_ALL_MALE
    cboFilterActivo.AddItem "Sí"
    cboFilterActivo.AddItem "No"
    cboFilterActivo.ListIndex = FILTER_ACTIVO_LIST_INDEX
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "Contacto", cbrMain
    pParametro.GetListViewSettings "Contacto", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
    tlbPin.Buttons("PIN").Value = pParametro.Usuario_LeerNumero("Contacto_Pin", tlbPin.Buttons("PIN").Value)
    If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If

    mLoading = False

    FillListView 0
    
    FormKeepOpenOnSelect = False
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "Contacto", cbrMain
    pParametro.SaveListViewSettings "Contacto", lvwData
    pParametro.Usuario_GuardarNumero "Contacto_Pin", tlbPin.Buttons("PIN").Value
End Sub

Private Sub lvwData_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = 0
    If ColumnHeader.Index - 1 = lvwData.SortKey Then
        lvwData.SortOrder = IIf(lvwData.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwData.SortKey = ColumnHeader.Index - 1
        lvwData.SortOrder = lvwAscending
    End If
    ColumnHeader.Icon = lvwData.SortOrder + 1
    FillListView 0
End Sub

Private Sub lvwData_DblClick()
    If GetFormIndex(FormWaitingForSelect) > 0 Then
        tlbMain_ButtonClick tlbMain.Buttons.Item("SELECT")
    Else
        tlbMain_ButtonClick tlbMain.Buttons.Item("PROPERTIES")
    End If
End Sub

Private Sub lvwData_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        lvwData_DblClick
    End If
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim FormIndex As Long
    Dim Contacto As Contacto
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_CONTACTO_ADD) Then
                Screen.MousePointer = vbHourglass
                
                Set Contacto = New Contacto
                frmContactoPropiedad.LoadDataAndShow Me, Contacto
                Set Contacto = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_CONTACTO_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                
                Set Contacto = New Contacto
                Contacto.IDContacto = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                If Contacto.Load() Then
                    frmContactoPropiedad.LoadDataAndShow Me, Contacto
                End If
                Set Contacto = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_CONTACTO_DELETE) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If MsgBox("¿Desea eliminar el Contacto seleccionado?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    Set Contacto = New Contacto
                    Contacto.IDContacto = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                    If Contacto.Load() Then
                        Call Contacto.Delete
                    End If
                    Set Contacto = Nothing
                End If
            End If
            lvwData.SetFocus
        Case "SELECT"
            FormIndex = GetFormIndex(FormWaitingForSelect)
            FormWaitingForSelect = ""
            If FormIndex >= 0 Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Set Contacto = New Contacto
                Contacto.IDContacto = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                If Not Contacto.Load() Then
                    Set Contacto = Nothing
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                If Not Contacto.Activo Then
                    Set Contacto = Nothing
                    MsgBox "No puede seleccionar este Item ya que está inactivo.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                
                Forms(FormIndex).ContactoSelected Contacto.IDContacto
                Forms(FormIndex).SetFocus
                If tlbPin.Buttons("PIN").Value = tbrUnpressed And Not FormKeepOpenOnSelect Then
                    Unload Me
                End If
                FormKeepOpenOnSelect = False
                Set Contacto = Nothing
                
                Screen.MousePointer = vbDefault
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

Public Sub FillComboBoxContactoGrupo()
    Dim KeySave As Long
    
    If cboFilterGrupo.ListCount > 0 Then
        KeySave = cboFilterGrupo.ItemData(cboFilterGrupo.ListIndex)
    End If
    Call CSM_Control_ComboBox.FillFromSQL(cboFilterGrupo, "(SELECT 0 AS IDContactoGrupo, '<Todos>' AS Nombre, 1 AS Orden) UNION (SELECT IDContactoGrupo, Nombre, 2 AS Orden FROM ContactoGrupo WHERE Activo = 1) ORDER BY Orden, Nombre", "IDContactoGrupo", "Nombre", "Grupos de Contactos", cscpItemOrfirst, KeySave)
End Sub
