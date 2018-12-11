VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmPersonaPrepago 
   Caption         =   "Prepagos de la Persona"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PersonaPrepago.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   10170
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   1111
      BandCount       =   2
      FixedOrder      =   -1  'True
      _CBWidth        =   10170
      _CBHeight       =   630
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain"
      MinWidth1       =   4410
      MinHeight1      =   570
      Width1          =   4410
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picFilterRutaGrupo"
      MinWidth2       =   3495
      MinHeight2      =   360
      Width2          =   3495
      FixedBackground2=   0   'False
      Key2            =   "RutaGrupo"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin VB.PictureBox picFilterRutaGrupo 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   6585
         ScaleHeight     =   360
         ScaleWidth      =   3495
         TabIndex        =   3
         Top             =   135
         Width           =   3495
         Begin VB.ComboBox cboRutaGrupo 
            Height          =   330
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   0
            Width           =   2250
         End
         Begin VB.Label lblRutaGrupo 
            AutoSize        =   -1  'True
            Caption         =   "Grupo de Rutas:"
            Height          =   210
            Left            =   0
            TabIndex        =   5
            Top             =   60
            Width           =   1185
         End
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   570
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   1005
         ButtonWidth     =   2170
         ButtonHeight    =   1005
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
   Begin MSComctlLib.ListView lvwData 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   900
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   7223
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
         Key             =   "IDRuta"
         Text            =   "Grupo de Rutas"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "FechaInicio"
         Text            =   "FechaInicio"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "ListaPrecio"
         Text            =   "Lista de Precios"
         Object.Width           =   6174
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   5355
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   635
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17410
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
Attribute VB_Name = "frmPersonaPrepago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FormWaitingForSelect As String

Private mPersona As Persona
Private mLoading As Boolean

Public Sub LoadDataAndShow(ByRef Persona As Persona)
    Set mPersona = Persona
    
    Load Me
    
    If Not FillListView(mPersona.IDPersona, 0, DATE_TIME_FIELD_NULL_VALUE) Then
        Unload Me
        Exit Sub
    End If

    Caption = "Prepagos del Pasajero: " & mPersona.ApellidoNombre

    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Sub ForceRefresh()
    FillListView mPersona.IDPersona, 0, DATE_TIME_FIELD_NULL_VALUE
End Sub

Public Function FillListView(ByVal IDPersona As Long, ByVal IDRutaGrupo As Long, ByVal FechaInicio As Date) As Boolean
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As Variant
    Dim CKeySave As Collection
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    
    If mPersona.IDPersona <> IDPersona Then
        Exit Function
    End If
    
    If mLoading Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If IDRutaGrupo = 0 Then
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
        KeySave = KEY_STRINGER & IDRutaGrupo & KEY_DELIMITER & FechaInicio
    End If
    
    SQL_Where = " WHERE PersonaPrepago.IDPersona = " & IDPersona
    
    If cboRutaGrupo.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "PersonaPrepago.IDRutaGrupo = " & cboRutaGrupo.ItemData(cboRutaGrupo.ListIndex)
    End If
        
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Select Case lvwData.SortKey
        Case 0  'GRUPO DE RUTAS
            SQL_OrderBy = " ORDER BY PersonaPrepago.IDRutaGrupo" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 1  'FECHA DE INICIO
            SQL_OrderBy = " ORDER BY PersonaPrepago.FechaInicio" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 2  'LISTA DE PRECIOS
            SQL_OrderBy = " ORDER BY ListaPrecio.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT PersonaPrepago.IDRutaGrupo, RutaGrupo.Nombre AS RutaGrupoNombre, PersonaPrepago.FechaInicio, ListaPrecio.Nombre AS ListaPrecio FROM (PersonaPrepago INNER JOIN RutaGrupo ON PersonaPrepago.IDRutaGrupo = RutaGrupo.IDRutaGrupo) INNER JOIN ListaPrecio ON PersonaPrepago.IDListaPrecio = ListaPrecio.IDListaPrecio" & SQL_Where & SQL_OrderBy
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & (.Fields("IDRutaGrupo").Value) & KEY_DELIMITER & .Fields("FechaInicio").Value, .Fields("RutaGrupoNombre").Value)
                ListItem.SubItems(1) = Format(.Fields("FechaInicio").Value, "Short Date")
                ListItem.SubItems(2) = .Fields("ListaPrecio").Value
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
    lvwData.SelectedItem.Selected = False
    Set lvwData.SelectedItem = lvwData.ListItems(KeySave)
    lvwData.SelectedItem.EnsureVisible
    
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
    
    Screen.MousePointer = vbDefault
    FillListView = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.PersonaPrepago.FillListView", "Error al obtener la Lista de Prepagos de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona
End Function

Private Sub cboRutaGrupo_Click()
    FillListView mPersona.IDPersona, 0, DATE_TIME_FIELD_NULL_VALUE
End Sub

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    ResizeControls NewHeight
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyN
                tlbMain_ButtonClick tlbMain.Buttons.Item("NEW")
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
    
    FillComboBoxRutaGrupo
    cboRutaGrupo.ListIndex = 0
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "PersonaPrepago", cbrMain
    pParametro.GetListViewSettings "PersonaPrepago", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1

    mLoading = False
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "PersonaPrepago", cbrMain
    pParametro.SaveListViewSettings "PersonaPrepago", lvwData
    Set mPersona = Nothing
    Set frmPersonaPrepago = Nothing
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
    FillListView mPersona.IDPersona, "", DATE_TIME_FIELD_NULL_VALUE
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
    Dim ItemIndex As Long
    Dim SelectedItemCount As Long
    Dim PersonaPrepago As PersonaPrepago
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_PERSONA_PREPAGO_ADD) Then
                Screen.MousePointer = vbHourglass
                
                Set PersonaPrepago = New PersonaPrepago
                PersonaPrepago.IDPersona = mPersona.IDPersona
                frmPersonaPrepagoPropiedad.LoadDataAndShow Me, PersonaPrepago
                Set PersonaPrepago = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_PERSONA_PREPAGO_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                
                Set PersonaPrepago = New PersonaPrepago
                PersonaPrepago.IDPersona = mPersona.IDPersona
                PersonaPrepago.IDRutaGrupo = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)
                PersonaPrepago.FechaInicio = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                If PersonaPrepago.Load() Then
                    frmPersonaPrepagoPropiedad.LoadDataAndShow Me, PersonaPrepago
                End If
                Set PersonaPrepago = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_PERSONA_PREPAGO_DELETE) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                If MsgBox("¿Desea eliminar el Prepago seleccionado?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                    Set PersonaPrepago = New PersonaPrepago
                    PersonaPrepago.IDPersona = mPersona.IDPersona
                    PersonaPrepago.IDRutaGrupo = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)
                    PersonaPrepago.FechaInicio = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                    If PersonaPrepago.Load() Then
                        Call PersonaPrepago.Delete
                    End If
                    
                    LogAccionAdd ENTIDAD_TIPO_PERSONA_PREPAGO, "ELIMINACIÓN: " & PersonaPrepago.Persona.ApellidoNombre & " || Grupo Ruta: " & PersonaPrepago.RutaGrupo.Nombre & " || Fecha Inicio: " & PersonaPrepago.FechaInicio_Formatted
                    
                    Set PersonaPrepago = Nothing
                End If
                lvwData.SetFocus
            End If
        Case "SELECT"
            FormIndex = GetFormIndex(FormWaitingForSelect)
            If FormIndex >= 0 Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                Forms(FormIndex).PersonaPrepagoSelected mPersona.IDPersona, CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)
                Forms(FormIndex).SetFocus
                Unload Me
                Screen.MousePointer = vbDefault
            End If
            FormWaitingForSelect = ""
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
End Sub

Public Sub FillComboBoxRutaGrupo()
    Dim recRutaGrupo As ADODB.Recordset
    Dim KeySave As String
    
    KeySave = cboRutaGrupo.Text

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recRutaGrupo = New ADODB.Recordset
    recRutaGrupo.Source = "SELECT IDRutaGrupo, Nombre FROM RutaGrupo WHERE Activo = 1 ORDER BY Nombre"
    recRutaGrupo.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    cboRutaGrupo.Clear
    cboRutaGrupo.AddItem CSM_Constant.ITEM_ALL_MALE
    Do While Not recRutaGrupo.EOF
        cboRutaGrupo.AddItem recRutaGrupo("Nombre").Value
        cboRutaGrupo.ItemData(cboRutaGrupo.NewIndex) = recRutaGrupo("IDRutaGrupo").Value
        recRutaGrupo.MoveNext
    Loop
    recRutaGrupo.Close
    Set recRutaGrupo = Nothing

    cboRutaGrupo.ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboRutaGrupo, KeySave, cscpItemOrfirst)
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.PersonaPrepago.FillComboBoxRutaGrupo", "Error al leer la lista de Grupos de Rutas."
End Sub

