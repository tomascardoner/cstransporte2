VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmRutaDetalleHorario 
   Caption         =   "Horarios del Detalle de Ruta"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   Icon            =   "RutaDetalleHorario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   8505
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   1111
      BandCount       =   1
      FixedOrder      =   -1  'True
      _CBWidth        =   8505
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
      Begin MSComctlLib.Toolbar tlbMain 
         Height          =   570
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   8385
         _ExtentX        =   14790
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
      Top             =   5055
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14473
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
      Left            =   120
      TabIndex        =   0
      Top             =   780
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "DiaSemana"
         Text            =   "Día"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "HoraInicio"
         Text            =   "Hora de inicio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "HoraFin"
         Text            =   "Hora de fin"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmRutaDetalleHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FormWaitingForSelect As String

Private mLoading As Boolean

Private mIDRuta As String
Private mIDLugar As Long
Private mLugarNombre As String

Public Sub LoadDataAndShow(ByVal IDRuta As String, ByVal IDLugar As Integer, ByVal LugarNombre As String)
    mIDRuta = IDRuta
    mIDLugar = IDLugar
    mLugarNombre = LugarNombre
    
    Load Me
    
    If Not FillListView(mIDRuta, IDLugar, 0) Then
        Unload Me
        Exit Sub
    End If
    
    Caption = "Horarios del Detalle de Ruta: " & mIDRuta & " - Lugar: " & mLugarNombre

    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Sub ForceRefresh()
    FillListView mIDRuta, mIDLugar, 0
End Sub

Public Function FillListView(ByVal IDRuta As String, ByVal IDLugar As Long, ByVal IDRutaDetalleHorario As Long) As Boolean
    Dim MousePointerSave As Integer
    Dim ListItem As MSComctlLib.ListItem
    Dim cmdData As ADODB.command
    Dim recData As ADODB.Recordset
    Dim KeySave As Variant
    Dim CKeySave As Collection
    
    If mIDRuta <> IDRuta Or mIDLugar <> IDLugar Then
        Exit Function
    End If
    
    If mLoading Then
        Exit Function
    End If
    
    MousePointerSave = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    If IDRutaDetalleHorario = 0 Then
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
        KeySave = KEY_STRINGER & IDRutaDetalleHorario
    End If
        
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = "usp_RutaDetalleHorario_List"
    cmdData.CommandType = adCmdStoredProc
    cmdData.Parameters.Append cmdData.CreateParameter("IDRuta", adChar, adParamInput, 20, IDRuta)
    cmdData.Parameters.Append cmdData.CreateParameter("IDLugar", adInteger, adParamInput, , IDLugar)
    
    Set recData = New ADODB.Recordset
    recData.Open cmdData, , adOpenForwardOnly, adLockReadOnly
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Select Case .Fields("DiaSemanaNumero").value
                    Case 0
                        Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("IDRutaDetalleHorario").value, CSM_Constant.ITEM_ALL_MALE)
                    Case 1 To 6
                        Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("IDRutaDetalleHorario").value, WeekdayName(.Fields("DiaSemanaNumero").value + 1))
                    Case 7
                        Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("IDRutaDetalleHorario").value, WeekdayName(1))
                    Case Else
                        Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("IDRutaDetalleHorario").value, CSM_Constant.ITEM_NOTSPECIFIED)
                End Select
                ListItem.SubItems(1) = Format(.Fields("HoraInicio").value, "Short Time")
                ListItem.SubItems(2) = Format(.Fields("HoraFin").value, "Short Time")
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
    ShowErrorMessage "Forms.RutaDetalleHorario.FillListView", "Error al obtener la Lista de Horarios del Detalle de la Ruta." & vbCr & vbCr & "IDRuta: " & IDRuta & " - IDLugar: " & IDLugar
End Function

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
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "RutaDetalleHorario", cbrMain
    pParametro.GetListViewSettings "RutaDetalleHorario", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1

    mLoading = False
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "RutaDetalleHorario", cbrMain
    pParametro.SaveListViewSettings "RutaDetalleHorario", lvwData
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
    Dim RutaDetalleHorario As RutaDetalleHorario
    Dim CRutaDetalleHorario As Collection
    
    Select Case Button.Key
        Case "NEW"
            If pCPermiso.GotPermission(PERMISO_RUTA_DETALLE_HORARIO_ADD) Then
                Screen.MousePointer = vbHourglass
                frmRutaDetalleHorarioAgregar.LoadDataAndShow Me, mIDRuta, mIDLugar, mLugarNombre
                If frmRutaDetalleHorarioAgregar.WindowState = vbMinimized Then
                    frmRutaDetalleHorarioAgregar.WindowState = vbNormal
                End If
                frmRutaDetalleHorarioAgregar.SetFocus
                Screen.MousePointer = vbDefault
            End If
        Case "PROPERTIES"
            If pCPermiso.GotPermission(PERMISO_RUTA_DETALLE_HORARIO_MODIFY) Then
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
                            Set RutaDetalleHorario = New RutaDetalleHorario
                            RutaDetalleHorario.IDRutaDetalleHorario = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                            If RutaDetalleHorario.Load() Then
                                frmRutaDetalleHorarioPropiedad.LoadDataAndShow Me, RutaDetalleHorario, mLugarNombre
                            End If
                            Set RutaDetalleHorario = Nothing
                        Else
                            MsgBox "No se puede Modificar más de un Horario a la vez.", vbExclamation, App.Title
                            Exit Sub
                        End If
                    End If
                Next ItemIndex
            End If
        Case "DELETE"
            If pCPermiso.GotPermission(PERMISO_RUTA_DETALLE_HORARIO_DELETE) Then
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
                
                If SelectedItemCount = 1 Then
                    If MsgBox("¿Desea eliminar el Horario seleccionado?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                        Set RutaDetalleHorario = New RutaDetalleHorario
                        RutaDetalleHorario.IDRutaDetalleHorario = Val(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1))
                        If RutaDetalleHorario.Load() Then
                            Call RutaDetalleHorario.Delete
                        End If
                        Set RutaDetalleHorario = Nothing
                    End If
                Else
                    If MsgBox("¿Desea eliminar los " & SelectedItemCount & " Horarios seleccionados?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                        Set RutaDetalleHorario = New RutaDetalleHorario
                        RutaDetalleHorario.RefreshListSkip = True
                        For ItemIndex = 1 To lvwData.ListItems.Count
                            If lvwData.ListItems(ItemIndex).Selected Then
                                RutaDetalleHorario.IDRutaDetalleHorario = Val(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1))
                                If RutaDetalleHorario.Load() Then
                                    Call RutaDetalleHorario.Delete
                                End If
                            End If
                        Next ItemIndex
                        Set RutaDetalleHorario = Nothing
                        RefreshList_RefreshRutaDetalleHorario mIDRuta, mIDLugar, 0
                    End If
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
                Forms(FormIndex).RutaDetalleHorarioSelected Val(Mid(lvwData.ListItems(ItemIndex).Key, Len(KEY_STRINGER) + 1))
                Forms(FormIndex).SetFocus
                Screen.MousePointer = vbDefault
            End If
            FormWaitingForSelect = ""
    End Select
End Sub

Private Sub ResizeControls(ByVal CoolBarHeight As Single)
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
    
    lvwData.Top = CoolBarHeight + CONTROL_SPACE
    lvwData.Left = CONTROL_SPACE
    lvwData.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lvwData.Height = ScaleHeight - lvwData.Top - CONTROL_SPACE - stbMain.Height
End Sub

