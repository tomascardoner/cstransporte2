VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmPersonaHorario 
   Caption         =   "Horarios de la Persona"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   Icon            =   "PersonaHorario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   9915
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   15
      TabIndex        =   6
      Top             =   6645
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
      TabIndex        =   2
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   1111
      BandCount       =   2
      FixedOrder      =   -1  'True
      _CBWidth        =   9915
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
      Child2          =   "picFilterRuta"
      MinWidth2       =   3015
      MinHeight2      =   360
      Width2          =   1095
      Key2            =   "FilterRuta"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin VB.PictureBox picFilterRuta 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   6810
         ScaleHeight     =   360
         ScaleWidth      =   3015
         TabIndex        =   4
         Top             =   135
         Width           =   3015
         Begin VB.ComboBox cboRuta 
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
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   0
            Width           =   2550
         End
         Begin VB.Label lblRuta 
            AutoSize        =   -1  'True
            Caption         =   "Ruta:"
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
            Left            =   0
            TabIndex        =   5
            Top             =   60
            Width           =   375
         End
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   570
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   6555
         _ExtentX        =   11562
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
      TabIndex        =   1
      Top             =   6615
      Width           =   9915
      _ExtentX        =   17489
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
            Object.Width           =   16272
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
      Height          =   4095
      Left            =   60
      TabIndex        =   0
      Top             =   2100
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   7223
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "DiaSemana"
         Text            =   "Día"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Hora"
         Text            =   "Hora"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "IDRuta"
         Text            =   "Ruta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "FechaDesde"
         Text            =   "Inicio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "FechaHasta"
         Text            =   "Fin"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Origen"
         Text            =   "Origen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "Destino"
         Text            =   "Destino"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmPersonaHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FormWaitingForSelect As String

Private mLoading As Boolean
Private mIDPersona As Long

Public Sub LoadDataAndShow(ByVal IDPersona As Long)
    Dim Persona As Persona

    mIDPersona = IDPersona
    
    Load Me
    
    If Not FillListView(mIDPersona, 0, Date, "") Then
        Unload Me
        Exit Sub
    End If

    Set Persona = New Persona
    Persona.IDPersona = mIDPersona
    If Persona.Load() Then
        Caption = "Horarios del Pasajero: " & Persona.ApellidoNombre
    End If
    Set Persona = Nothing

    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Sub ForceRefresh()
    FillListView mIDPersona, 0, Date, ""
End Sub

Public Function FillListView(ByVal IDPersona As Long, ByVal DiaSemana As Long, ByVal Hora As Date, ByVal IDRuta As String) As Boolean
    Dim MousePointerSave As Integer
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As Variant
    Dim CKeySave As Collection
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    
    If mIDPersona <> IDPersona Then
        Exit Function
    End If
    
    If mLoading Then
        Exit Function
    End If
    
    MousePointerSave = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    If DiaSemana = 0 Then
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
        KeySave = KEY_STRINGER & DiaSemana & KEY_DELIMITER & Hora & KEY_DELIMITER & IDRuta
    End If
        
    SQL_Where = ""
    
    'PERSONAL
    If pPersonal Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Horario.Personal = 0"
    End If
    
    SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "PersonaHorario.IDPersona = " & IDPersona
    
    If cboRuta.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "PersonaHorario.IDRuta = '" & ReplaceQuote(cboRuta.Text) & "'"
    Else
        If pCPermiso.RutaWhere <> "" Then
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "PersonaHorario")
        End If
    End If
    
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Select Case lvwData.SortKey
        Case 0  'DIA SEMANA
            SQL_OrderBy = " ORDER BY PersonaHorario.DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", PersonaHorario.Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", PersonaHorario.IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 1  'HORA
            SQL_OrderBy = " ORDER BY PersonaHorario.Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", PersonaHorario.DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", PersonaHorario.IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 2  'RUTA
            SQL_OrderBy = " ORDER BY PersonaHorario.IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", PersonaHorario.DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", PersonaHorario.Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 3  'FECHA DESDE
            SQL_OrderBy = " ORDER BY PersonaHorario.FechaDesde" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", PersonaHorario.DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", PersonaHorario.Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", PersonaHorario.IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 3  'FECHA HASTA
            SQL_OrderBy = " ORDER BY PersonaHorario.FechaHasta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", PersonaHorario.DiaSemana" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", PersonaHorario.Hora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", PersonaHorario.IDRuta" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT PersonaHorario.DiaSemana, PersonaHorario.Hora, PersonaHorario.IDRuta, PersonaHorario.FechaDesde, PersonaHorario.FechaHasta, Lugar_Origen.Nombre AS Origen, PersonaHorario.Sube, Lugar_Destino.Nombre AS Destino, PersonaHorario.Baja FROM ((Horario INNER JOIN PersonaHorario ON Horario.DiaSemana = PersonaHorario.DiaSemana AND Horario.Hora = PersonaHorario.Hora AND Horario.IDRuta = PersonaHorario.IDRuta) LEFT JOIN Lugar AS Lugar_Origen ON PersonaHorario.IDOrigen = Lugar_Origen.IDLugar) LEFT JOIN Lugar AS Lugar_Destino ON PersonaHorario.IDDestino = Lugar_Destino.IDLugar" & SQL_Where & SQL_OrderBy
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("DiaSemana").Value & KEY_DELIMITER & Format(.Fields("Hora").Value, "hh:nn:ss") & KEY_DELIMITER & RTrim(.Fields("IDRuta").Value), WeekdayName(.Fields("DiaSemana").Value))
                ListItem.SubItems(1) = Format(.Fields("Hora").Value, "Short Time")
                ListItem.SubItems(2) = RTrim(.Fields("IDRuta").Value)
                ListItem.SubItems(3) = IIf(IsNull(.Fields("FechaDesde").Value), "", Format(.Fields("FechaDesde").Value, "Short Date"))
                ListItem.SubItems(4) = IIf(IsNull(.Fields("FechaHasta").Value), "", Format(.Fields("FechaHasta").Value, "Short Date"))
                ListItem.SubItems(5) = IIf(IsNull(.Fields("Sube").Value), .Fields("Origen").Value & "", .Fields("Sube").Value)
                ListItem.SubItems(6) = IIf(IsNull(.Fields("Baja").Value), .Fields("Destino").Value & "", .Fields("Baja").Value)
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
    
    Screen.MousePointer = MousePointerSave
    FillListView = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.PersonaHorario.FillListView", "Error al obtener la Lista de Horarios de la Persona." & vbCr & vbCr & "IDPersona: " & IDPersona
End Function

Public Sub FillComboBoxRuta()
    Dim recRuta As ADODB.Recordset
    Dim KeySave As String
    
    KeySave = cboRuta.Text

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recRuta = New ADODB.Recordset
    recRuta.Source = "SELECT IDRuta FROM Ruta WHERE IDRuta <> '" & ReplaceQuote(pParametro.Ruta_ID_Otra) & "' AND Activo = 1" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & " ORDER BY IDRuta"
    recRuta.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    cboRuta.Clear
    cboRuta.AddItem "<Todas>"
    Do While Not recRuta.EOF
        cboRuta.AddItem RTrim(recRuta("IDRuta").Value)
        recRuta.MoveNext
    Loop
    recRuta.Close
    Set recRuta = Nothing

    cboRuta.ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboRuta, KeySave, cscpItemOrfirst)
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.PersonaHorario.FillComboBoxRuta", "Error al leer la lista de Rutas."
End Sub

Private Sub cboRuta_Click()
    FillListView mIDPersona, 0, Date, ""
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
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL PIN
    Set tlbPin.ImageList = frmMDI.ilsFormPin
    '//////////////////////////////////////////////////////////
    
    FillComboBoxRuta
    cboRuta.ListIndex = 0
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "PersonaHorario", cbrMain
    pParametro.GetListViewSettings "PersonaHorario", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
    tlbPin.Buttons("PIN").Value = pParametro.Usuario_LeerNumero("PersonaHorario_Pin", tlbPin.Buttons("PIN").Value)
    If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If

    mLoading = False
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "PersonaHorario", cbrMain
    pParametro.SaveListViewSettings "PersonaHorario", lvwData
    pParametro.Usuario_GuardarNumero "PersonaHorario_Pin", tlbPin.Buttons("PIN").Value
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
    FillListView mIDPersona, 0, Date, ""
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
    Dim PersonaHorario As PersonaHorario
    Dim CPersonaHorario As Collection
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_PERSONA_HORARIO_ADD) Then
                Screen.MousePointer = vbHourglass
                frmHorario.Show
                If frmHorario.WindowState = vbMinimized Then
                    frmHorario.WindowState = vbNormal
                End If
                frmHorario.FormWaitingForSelect = Me.Name
                frmHorario.AllowMultipleSelect = True
                frmHorario.AllowMultipleRuta = False
                frmHorario.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_PERSONA_HORARIO_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If

                SelectedItemCount = 0
                For ItemIndex = 1 To lvwData.ListItems.Count
                    If lvwData.ListItems(ItemIndex).Selected Then
                        SelectedItemCount = SelectedItemCount + 1
                        If SelectedItemCount = 1 Then
                            Set PersonaHorario = New PersonaHorario
                            PersonaHorario.IDPersona = mIDPersona
                            PersonaHorario.DiaSemana = Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                            PersonaHorario.Hora = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                            PersonaHorario.IDRuta = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)
                            If PersonaHorario.Load() Then
                                Set CPersonaHorario = New Collection
                                
                                CPersonaHorario.Add PersonaHorario
                                
                                frmPersonaHorarioPropiedad.LoadDataAndShow Me, CPersonaHorario
                                With frmPersonaHorarioPropiedad
                                    .dtpFechaDesde.Enabled = False
                                    .cmdHoyDesde.Enabled = False
                                    .dtpFechaHasta.Enabled = False
                                    .cmdHoyHasta.Enabled = False
                                    .datcboOrigen.Enabled = False
                                    .txtSube.Enabled = False
                                    .datcboDestino.Enabled = False
                                    .txtBaja.Enabled = False
                                    .cmdOK.Enabled = False
                                End With
                                
                                Set CPersonaHorario = Nothing
                            End If
                            Set PersonaHorario = Nothing
                        Else
                            MsgBox "No se puede Modificar más de un Horario a la vez.", vbExclamation, App.Title
                            Exit Sub
                        End If
                    End If
                Next ItemIndex
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_PERSONA_HORARIO_DELETE) Then
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
                
                If MsgBox(IIf(SelectedItemCount = 1, "¿Desea eliminar el Horario seleccionado?" & vbCr & "Se eliminarán también las Reservas Fijas generadas a Futuro, basadas en este Horario.", "¿Desea eliminar los " & SelectedItemCount & " Horarios seleccionados?" & vbCr & "Se eliminarán también las Reservas Fijas generadas a Futuro, basadas en estos Horarios."), vbExclamation + vbYesNo, App.Title) = vbYes Then
                    Set PersonaHorario = New PersonaHorario
                    If SelectedItemCount = 1 Then
                        PersonaHorario.IDPersona = mIDPersona
                        PersonaHorario.DiaSemana = Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                        PersonaHorario.Hora = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                        PersonaHorario.IDRuta = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)
                        If PersonaHorario.Load() Then
                            Call PersonaHorario.Delete
                        End If
                    Else
                        PersonaHorario.RefreshList = False
                        For ItemIndex = 1 To lvwData.ListItems.Count
                            If lvwData.ListItems(ItemIndex).Selected Then
                                PersonaHorario.IDPersona = mIDPersona
                                PersonaHorario.DiaSemana = Val(GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
                                PersonaHorario.Hora = CSM_String.GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
                                PersonaHorario.IDRuta = CSM_String.GetSubString(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)
                                If PersonaHorario.Load() Then
                                    Call PersonaHorario.Delete
                                End If
                            End If
                        Next ItemIndex
                        RefreshList_RefreshPersonaHorario mIDPersona, 1, Time, ""
                    End If
                    Set PersonaHorario = Nothing
                    lvwData.SetFocus
                End If
            End If
        Case "SELECT"
            FormIndex = GetFormIndex(FormWaitingForSelect)
            If FormIndex >= 0 Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                SelectedItemCount = 0
                For ItemIndex = 1 To lvwData.ListItems.Count
                    If lvwData.ListItems(ItemIndex).Selected Then
                        SelectedItemCount = SelectedItemCount + 1
                        If SelectedItemCount > 1 Then
                            MsgBox "No se puede Seleccionar más de un Horario a la vez.", vbExclamation, App.Title
                            Exit Sub
                        End If
                    End If
                Next ItemIndex
                
                Screen.MousePointer = vbHourglass
                Forms(FormIndex).PersonaHorarioSelected mIDPersona, Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)), CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)), CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)
                Forms(FormIndex).SetFocus
                If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
                    Unload Me
                End If
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
    
    tlbPin.Top = ScaleHeight - 330
    tlbPin.Left = 15
End Sub

Public Sub MultipleHorarioSelected(ByVal Horarios As Collection)
    Dim PersonaHorario As PersonaHorario
    Dim CPersonaHorario As Collection
    Dim Horario As Variant
    
    Screen.MousePointer = vbHourglass
    
    Set CPersonaHorario = New Collection
    For Each Horario In Horarios
        Set PersonaHorario = New PersonaHorario
        
        PersonaHorario.IDPersona = mIDPersona
        PersonaHorario.DiaSemana = Val(GetSubString(Mid(Horario, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
        PersonaHorario.Hora = CDate(GetSubString(Mid(Horario, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER))
        PersonaHorario.IDRuta = CSM_String.GetSubString(Mid(Horario, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)
        
        CPersonaHorario.Add PersonaHorario
    Next Horario
    
    frmPersonaHorarioPropiedad.LoadDataAndShow Me, CPersonaHorario
    Screen.MousePointer = vbDefault
    
    Set PersonaHorario = Nothing
    Set CPersonaHorario = Nothing
End Sub
