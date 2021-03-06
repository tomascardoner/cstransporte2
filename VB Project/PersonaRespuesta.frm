VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmPersonaRespuesta 
   Caption         =   "Respuestas de la Persona"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   Icon            =   "PersonaRespuesta.frx":0000
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
      MinWidth1       =   6450
      MinHeight1      =   570
      Width1          =   6450
      FixedBackground1=   0   'False
      Key1            =   "Toolbar"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picFilterActivo"
      MinWidth2       =   1605
      MinHeight2      =   330
      Width2          =   1605
      Key2            =   "FilterActivo"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin VB.PictureBox picFilterActivo 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   7500
         ScaleHeight     =   330
         ScaleWidth      =   1605
         TabIndex        =   4
         Top             =   150
         Width           =   1605
         Begin VB.ComboBox cboFilterActivo 
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
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   0
            Width           =   990
         End
         Begin VB.Label lblFilterActivo 
            AutoSize        =   -1  'True
            Caption         =   "Activo:"
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
            TabIndex        =   6
            Top             =   60
            Width           =   510
         End
      End
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   570
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   1005
         ButtonWidth     =   2434
         ButtonHeight    =   1005
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
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
               Caption         =   "Activar Todo"
               Key             =   "ACTIVATE"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Desactivar Todo"
               Key             =   "DEACTIVATE"
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
      Left            =   60
      TabIndex        =   0
      Top             =   1200
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
         Key             =   "FechaHora"
         Text            =   "Fecha/Hora"
         Object.Width           =   2417
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Respuesta"
         Text            =   "Respuesta"
         Object.Width           =   8114
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Activo"
         Text            =   "Activo"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmPersonaRespuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean

Private mPersona As Persona

Public Sub LoadDataAndShow(ByRef Persona As Persona)
    
    Set mPersona = Persona
    
    Load Me
    
    Caption = "Respuestas de la Persona: " & Persona.IDPersona & " - " & Persona.ApellidoNombre
    
    If Not FillListView(mPersona.IDPersona, DATE_TIME_FIELD_NULL_VALUE) Then
        Unload Me
        Exit Sub
    End If

    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Function FillListView(ByVal IDPersona As Long, ByVal FechaHora As Date) As Boolean
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    Dim SQL_Where As String
    
    If mLoading Or IDPersona <> mPersona.IDPersona Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If FechaHora = DATE_TIME_FIELD_NULL_VALUE Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & Format(FechaHora, "Short Date") & " " & Format(FechaHora, "Short Time")
    End If
    
    SQL_Where = " WHERE IDPersona = " & mPersona.IDPersona
    
    If cboFilterActivo.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Activo = " & IIf(cboFilterActivo.ListIndex = 1, 1, 0)
    End If
    
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT FechaHora, Respuesta, Activo FROM PersonaRespuesta" & SQL_Where & " ORDER BY FechaHora DESC"
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & Format(.Fields("FechaHora").Value, "Short Date") & " " & Format(.Fields("FechaHora").Value, "Short Time"), Format(.Fields("FechaHora").Value, "Short Date") & " " & Format(.Fields("FechaHora").Value, "Short Time"))
                ListItem.SubItems(1) = .Fields("Respuesta").Value
                ListItem.SubItems(2) = GetBooleanString(.Fields("Activo").Value)
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
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hWnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    
    FillListView = True
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.PersonaRespuesta.FillListView", "Error al obtener la lista de las Respuestas de la Persona."
End Function

Private Sub cboFilterActivo_Click()
    FillListView mPersona.IDPersona, DATE_TIME_FIELD_NULL_VALUE
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
    tlbMain.Buttons("ACTIVATE").Image = "ACTIVATE"
    tlbMain.Buttons("DEACTIVATE").Image = "DEACTIVATE"
    
    cbrMain.Bands("Toolbar").MinWidth = CSM_Control_Toolbar.GetTotalWidth(tlbMain)
    '//////////////////////////////////////////////////////////
    
    cboFilterActivo.AddItem ITEM_ALL_MALE
    cboFilterActivo.AddItem "S�"
    cboFilterActivo.AddItem "No"
    cboFilterActivo.ListIndex = FILTER_ACTIVO_LIST_INDEX
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "PersonaRespuesta", cbrMain
    pParametro.GetListViewSettings "PersonaRespuesta", lvwData
    
    mLoading = False
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "PersonaRespuesta", cbrMain
    pParametro.SaveListViewSettings "PersonaRespuesta", lvwData
    Set mPersona = Nothing
    Set frmPersonaRespuesta = Nothing
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
    Dim PersonaRespuesta As PersonaRespuesta
    
    Set PersonaRespuesta = New PersonaRespuesta
    PersonaRespuesta.IDPersona = mPersona.IDPersona
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_PERSONA_RESPUESTA_ADD) Then
                Screen.MousePointer = vbHourglass
                frmPersonaRespuestaPropiedad.LoadDataAndShow Me, PersonaRespuesta
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_PERSONA_RESPUESTA_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ning�n Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                PersonaRespuesta.FechaHora = CDate(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                If Not PersonaRespuesta.Load() Then
                    Set PersonaRespuesta = Nothing
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                frmPersonaRespuestaPropiedad.LoadDataAndShow Me, PersonaRespuesta
                Screen.MousePointer = vbDefault
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_PERSONA_RESPUESTA_DELETE) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ning�n Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If MsgBox("�Desea eliminar la Respuesta seleccionada?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                    
                    PersonaRespuesta.FechaHora = CDate(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                    If Not PersonaRespuesta.Load() Then
                        Set PersonaRespuesta = Nothing
                        lvwData.SetFocus
                        Exit Sub
                    End If
                    If Not PersonaRespuesta.Delete() Then
                        Set PersonaRespuesta = Nothing
                        lvwData.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        Case "ACTIVATE"
            If pCPermiso.GotPermission(PERMISO_PERSONA_RESPUESTA_ACTIVATE_ALL) Then
                If MsgBox("Se Activar�n todas las Respuestas de esta Persona." & vbCr & vbCr & "�Desea continuar?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                    Call PersonaRespuesta.ActivateAll(True)
                End If
            End If
            lvwData.SetFocus
        Case "DEACTIVATE"
            If pCPermiso.GotPermission(PERMISO_PERSONA_RESPUESTA_ACTIVATE_ALL) Then
                If MsgBox("Se Desactivar�n todas las Respuestas de esta Persona." & vbCr & vbCr & "�Desea continuar?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                    Call PersonaRespuesta.ActivateAll(False)
                End If
            End If
            lvwData.SetFocus
    End Select
    
    Set PersonaRespuesta = Nothing
End Sub

Private Sub ResizeControls(ByVal CoolBarHeight As Single)
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
    
    lvwData.Top = CoolBarHeight + CONTROL_SPACE
    lvwData.Left = CONTROL_SPACE
    lvwData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lvwData.Height = ScaleHeight - lvwData.Top - CONTROL_SPACE - stbMain.Height
End Sub
