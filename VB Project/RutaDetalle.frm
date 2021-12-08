VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmRutaDetalle 
   Caption         =   "Detalle de Ruta"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   Icon            =   "RutaDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   8955
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
      Height          =   1020
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   1799
      BandCount       =   2
      FixedOrder      =   -1  'True
      _CBWidth        =   8955
      _CBHeight       =   1020
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain"
      MinWidth1       =   8670
      MinHeight1      =   570
      Width1          =   8670
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
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   8700
         TabIndex        =   4
         Top             =   630
         Width           =   8700
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
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   1005
         ButtonWidth     =   2170
         ButtonHeight    =   1005
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
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
               Caption         =   "Subir"
               Key             =   "UP"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Bajar"
               Key             =   "DOWN"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exclusiones"
               Key             =   "EXCLUSION"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   5040
      Width           =   8955
      _ExtentX        =   15796
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
            Object.Width           =   14579
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
      Height          =   2955
      Left            =   180
      TabIndex        =   0
      Top             =   1560
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   5212
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Lugar"
         Text            =   "Lugar"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Grupo"
         Text            =   "Grupo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "Duracion"
         Text            =   "Duración"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "HoraInicio"
         Text            =   "Excluído desde"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "HoraFin"
         Text            =   "Excluído hasta"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Key             =   "Exclusiones"
         Text            =   "Exclusiones"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "frmRutaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FormWaitingForSelect As String

Private mLoading As Boolean

Public Sub ForceRefresh()
    FillListView cboRuta.Text, 0
End Sub

Public Sub FillListView(ByVal IDRuta As String, ByVal IDLugar As Long)
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    
    Dim SQLStatement As String
    
    If IDRuta <> cboRuta.Text Then
        Exit Sub
    End If
    
    If mLoading Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    If IDRuta = "" Then
        If Not lvwData.SelectedItem Is Nothing Then
            KeySave = lvwData.SelectedItem.Key
        End If
    Else
        KeySave = KEY_STRINGER & IDLugar
    End If
        
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    SQLStatement = "SELECT rd.IDLugar, l.Nombre AS Lugar, lg.Nombre AS LugarGrupo, rd.Duracion, rd.HoraInicio, rd.HoraFin, COUNT(rdh.IDRutaDetalleHorario) AS Exclusiones" & vbCr
    SQLStatement = SQLStatement & "FROM RutaDetalle AS rd" & vbCr
    SQLStatement = SQLStatement & "INNER JOIN Lugar AS l ON rd.IDLugar = l.IDLugar" & vbCr
    SQLStatement = SQLStatement & "INNER JOIN LugarGrupo AS lg ON rd.IDLugarGrupo = lg.IDLugarGrupo" & vbCr
    SQLStatement = SQLStatement & "LEFT JOIN RutaDetalleHorario AS rdh ON rd.IDRuta = rdh.IDRuta AND rd.IDLugar = rdh.IDLugar" & vbCr
    SQLStatement = SQLStatement & "WHERE rd.IDRuta = '" & ReplaceQuote(cboRuta.Text) & "'" & vbCr
    SQLStatement = SQLStatement & "GROUP BY rd.Indice, rd.IDLugar, l.Nombre, lg.Nombre, rd.Duracion, rd.HoraInicio, rd.HoraFin" & vbCr
    SQLStatement = SQLStatement & "ORDER BY rd.Indice" & vbCr
    
    Set recData = New ADODB.Recordset
    recData.Source = SQLStatement
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("IDLugar").value, .Fields("Lugar").value)
                ListItem.SubItems(1) = .Fields("LugarGrupo").value
                ListItem.SubItems(2) = CSM_Function.IfIsNull_Space(.Fields("Duracion").value)
                ListItem.SubItems(3) = IIf(IsNull(.Fields("HoraInicio").value), "", Format(.Fields("HoraInicio").value, "Short Time"))
                ListItem.SubItems(4) = IIf(IsNull(.Fields("HoraFin").value), "", Format(.Fields("HoraFin").value, "Short Time"))
                ListItem.SubItems(5) = CSM_Function.IfIsZero_Space(.Fields("Exclusiones").value)
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
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.RutaDetalle.FillListView", "Error al obtener el Detalle de la Ruta."
End Sub

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
    Do While Not recRuta.EOF
        cboRuta.AddItem RTrim(recRuta("IDRuta").value)
        recRuta.MoveNext
    Loop
    recRuta.Close
    Set recRuta = Nothing

    cboRuta.ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboRuta, IDRutaSave, cscpItemOrFirst)
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.RutaDetalle.FillComboBoxRuta", "Error al leer la lista de Rutas."
End Sub

Private Sub cboRuta_Click()
    FillListView cboRuta.Text, 0
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
    tlbMain.Buttons("UP").Image = "UP"
    tlbMain.Buttons("DOWN").Image = "DOWN"
    tlbMain.Buttons("EXCLUSION").Image = "HORARIO"
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
    pParametro.GetCoolBarSettings "RutaDetalle", cbrMain
    pParametro.GetListViewSettings "RutaDetalle", lvwData
    tlbPin.Buttons("PIN").value = pParametro.Usuario_LeerNumero("RutaDetalle_Pin", tlbPin.Buttons("PIN").value)
    If tlbPin.Buttons("PIN").value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If

    mLoading = False

    FillListView cboRuta.Text, 0
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "RutaDetalle", cbrMain
    pParametro.SaveListViewSettings "RutaDetalle", lvwData
    pParametro.Usuario_GuardarNumero "RutaDetalle_Pin", tlbPin.Buttons("PIN").value
    Set frmRutaDetalle = Nothing
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
    Dim RutaDetalle As RutaDetalle
    Dim FormIndex As Long
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_RUTA_DETALLE_ADD) Then
                Screen.MousePointer = vbHourglass
                
                Set RutaDetalle = New RutaDetalle
                RutaDetalle.IDRuta = cboRuta.Text
                frmRutaDetallePropiedad.LoadDataAndShow Me, RutaDetalle
                Set RutaDetalle = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_RUTA_DETALLE_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                
                Set RutaDetalle = New RutaDetalle
                RutaDetalle.IDRuta = cboRuta.Text
                RutaDetalle.IDLugar = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                If RutaDetalle.Load() Then
                    frmRutaDetallePropiedad.LoadDataAndShow Me, RutaDetalle
                End If
                Set RutaDetalle = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_RUTA_DETALLE_DELETE) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                If MsgBox("¿Desea eliminar el Detalle de Ruta seleccionado?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    Set RutaDetalle = New RutaDetalle
                    RutaDetalle.IDRuta = cboRuta.Text
                    RutaDetalle.IDLugar = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                    If RutaDetalle.Load() Then
                        Call RutaDetalle.Delete
                    End If
                    Set RutaDetalle = Nothing
                    
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
                
                Screen.MousePointer = vbHourglass
                Forms(FormIndex).HorarioSelected Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER)), CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)), CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)
                Forms(FormIndex).SetFocus
                If tlbPin.Buttons("PIN").value = tbrUnpressed Then
                    Unload Me
                End If
                Screen.MousePointer = vbDefault
            End If
            FormWaitingForSelect = ""
        Case "UP"
            If pCPermiso.GotPermission(PERMISO_RUTA_DETALLE_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If lvwData.SelectedItem.Index = 1 Then
                    Exit Sub
                End If
            
                Screen.MousePointer = vbHourglass
                
                Set RutaDetalle = New RutaDetalle
                RutaDetalle.IDRuta = cboRuta.Text
                RutaDetalle.IDLugar = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                If RutaDetalle.Load() Then
                    RutaDetalle.Indice = RutaDetalle.Indice - 15
                    If RutaDetalle.Update() Then
                        Call RutaDetalle.ReIndex(False)
                    End If
                End If
                Set RutaDetalle = Nothing
                
                Screen.MousePointer = vbDefault
            
            End If
        Case "DOWN"
            If pCPermiso.GotPermission(PERMISO_RUTA_DETALLE_MODIFY) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                If lvwData.SelectedItem.Index = lvwData.ListItems.Count Then
                    Exit Sub
                End If
            
                Screen.MousePointer = vbHourglass
                
                Set RutaDetalle = New RutaDetalle
                RutaDetalle.IDRuta = cboRuta.Text
                RutaDetalle.IDLugar = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                If RutaDetalle.Load() Then
                    RutaDetalle.Indice = RutaDetalle.Indice + 15
                    If RutaDetalle.Update() Then
                        Call RutaDetalle.ReIndex(True)
                    End If
                End If
                Set RutaDetalle = Nothing
                
                Screen.MousePointer = vbDefault
            End If
        Case "EXCLUSION"
            If pCPermiso.GotPermission(PERMISO_RUTA_DETALLE_HORARIO) Then
                If lvwData.SelectedItem Is Nothing Then
                    MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
                    lvwData.SetFocus
                    Exit Sub
                End If
                
                Screen.MousePointer = vbHourglass
                frmRutaDetalleHorario.LoadDataAndShow cboRuta.Text, Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1)), lvwData.SelectedItem.Text
                Screen.MousePointer = vbDefault
            End If
    End Select
End Sub

Private Sub tlbPin_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.value = tbrUnpressed Then
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
