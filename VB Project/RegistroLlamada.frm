VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistroLlamada 
   Caption         =   "Registro de Llamadas Telefónicas"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RegistroLlamada.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   7035
   Begin MSComctlLib.Toolbar tlbPin 
      Height          =   330
      Left            =   60
      TabIndex        =   3
      Top             =   5220
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
      Height          =   810
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   1429
      BandCount       =   4
      FixedOrder      =   -1  'True
      _CBWidth        =   7035
      _CBHeight       =   810
      _Version        =   "6.7.9782"
      Child1          =   "picFecha"
      MinWidth1       =   2955
      MinHeight1      =   360
      Width1          =   2955
      Key1            =   "Fecha"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "picSucursal"
      MinWidth2       =   3615
      MinHeight2      =   360
      Width2          =   9465
      Key2            =   "Sucursal"
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Child3          =   "picShowTop"
      MinWidth3       =   2580
      MinHeight3      =   360
      Width3          =   2580
      Key3            =   "ShowTop"
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Child4          =   "picPersona"
      MinWidth4       =   3795
      MinHeight4      =   360
      Width4          =   3795
      Key4            =   "Persona"
      NewRow4         =   0   'False
      AllowVertical4  =   0   'False
      Begin VB.PictureBox picPersona 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   2970
         ScaleHeight     =   360
         ScaleWidth      =   3975
         TabIndex        =   16
         Top             =   420
         Width           =   3975
         Begin VB.CommandButton cmdMostrarPersona 
            Caption         =   "Mostrar Persona"
            Height          =   315
            Left            =   0
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   1395
         End
         Begin VB.CommandButton cmdMostrarPersonas 
            Caption         =   "Mostrar todas las Personas"
            Height          =   315
            Left            =   1500
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.PictureBox picShowTop 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   2580
         TabIndex        =   13
         Top             =   420
         Width           =   2580
         Begin VB.ComboBox cboShowTop 
            Height          =   330
            Left            =   660
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblShowTop 
            AutoSize        =   -1  'True
            Caption         =   "Ultimas                          llamadas."
            Height          =   210
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   2355
         End
      End
      Begin VB.PictureBox picFecha 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   30
         ScaleHeight     =   360
         ScaleWidth      =   2955
         TabIndex        =   7
         Top             =   30
         Width           =   2955
         Begin VB.CommandButton cmdHoyDesde 
            Height          =   315
            Left            =   2640
            Picture         =   "RegistroLlamada.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Hoy"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton cmdSiguienteDesde 
            Height          =   315
            Left            =   2340
            Picture         =   "RegistroLlamada.frx":06D4
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Siguiente"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton cmdAnteriorDesde 
            Height          =   315
            Left            =   600
            Picture         =   "RegistroLlamada.frx":0C5E
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Anterior"
            Top             =   0
            Width           =   300
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   315
            Left            =   900
            TabIndex        =   11
            Top             =   0
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   108920833
            CurrentDate     =   36950
         End
         Begin VB.Label lblFecha 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   210
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   495
         End
      End
      Begin VB.PictureBox picSucursal 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   3210
         ScaleHeight     =   360
         ScaleWidth      =   3735
         TabIndex        =   4
         Top             =   30
         Width           =   3735
         Begin VB.ComboBox cboSucursal 
            Height          =   330
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label lblSucursal 
            AutoSize        =   -1  'True
            Caption         =   "Sucursal:"
            Height          =   210
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   690
         End
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   5295
      Width           =   7035
      _ExtentX        =   12409
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
            Object.Width           =   11183
            MinWidth        =   176
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
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   6376
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "IDRegistroLlamada"
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Fecha"
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Hora"
         Text            =   "Hora"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Sucursal"
         Text            =   "Sucursal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "SucursalTelefonoNumero"
         Text            =   "Línea Telefónica"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Telefono"
         Text            =   "Número de Teléfono"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "Tipo"
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "Entidad"
         Text            =   "Entidad"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmRegistroLlamada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean
Private mCSucursal As Collection

Public FormWaitingForSelect As String

Public Sub LoadDataAndShow()
    Load Me
    
    If Not FillListView() Then
        Unload Me
        Exit Sub
    End If

    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Function FillListView() As Boolean
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As String
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    Dim NumeroTelefono As String
    
    If mLoading Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    If Not lvwData.SelectedItem Is Nothing Then
        KeySave = lvwData.SelectedItem.Key
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    SQL_Where = ""
    
    'FILTRO DE FECHA
    SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "RegistroLlamada.FechaHora BETWEEN '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & " 00:00:00' AND '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & " 23:59:59'"
    
    'FILTRO DE SUCURSAL
    If cboSucursal.ListIndex > 0 Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Sucursal.IDSucursal = '" & mCSucursal(cboSucursal.ListIndex) & "'"
    End If
    
    Select Case lvwData.SortKey
        Case 0  'ID
            SQL_OrderBy = " ORDER BY RegistroLlamada.IDRegistroLlamada" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 1, 6, 7 'FECHA + HORA
            SQL_OrderBy = " ORDER BY RegistroLlamada.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 2  'HORA + FECHA
            SQL_OrderBy = " ORDER BY convert(char(8), RegistroLlamada.FechaHora, 108)" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", RegistroLlamada.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 3  'SUCURSAL + SUCURSALTELEFONONUMERO + FECHA + HORA
            SQL_OrderBy = " ORDER BY Sucursal.Nombre" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", SucursalTelefono.TelefonoNumero" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", RegistroLlamada.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 4  'SUCURSALTELEFONONUMERO + FECHA + HORA
            SQL_OrderBy = " ORDER BY SucursalTelefono.TelefonoNumero" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", RegistroLlamada.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
        Case 5  'NUMERO TELEFONO + FECHA + HORA
            SQL_OrderBy = " ORDER BY RegistroLlamada.TelefonoNumero" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC") & ", RegistroLlamada.FechaHora" & IIf(lvwData.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    lvwData.ListItems.Clear
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT TOP " & cboShowTop.Text & " RegistroLlamada.IDRegistroLlamada, RegistroLlamada.FechaHora, Sucursal.Nombre AS SucursalNombre, RTRIM(SucursalTelefono.TelefonoNumero) AS SucursalTelefonoNumero, RegistroLlamada.TelefonoNumero "
    recData.Source = recData.Source & "FROM (RegistroLlamada INNER JOIN SucursalTelefono ON RegistroLlamada.IDSucursalTelefono = SucursalTelefono.IDSucursalTelefono) INNER JOIN Sucursal ON SucursalTelefono.IDSucursal = Sucursal.IDSucursal" & SQL_Where & SQL_OrderBy
    recData.MaxRecords = pParametro.Recordset_MaxRecords
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("IDRegistroLlamada").Value, .Fields("IDRegistroLlamada").Value)
                ListItem.SubItems(1) = Format(.Fields("FechaHora").Value, "Short Date")
                ListItem.SubItems(2) = Format(.Fields("FechaHora").Value, "Short Time")
                ListItem.SubItems(3) = .Fields("SucursalNombre").Value
                ListItem.SubItems(4) = .Fields("SucursalTelefonoNumero").Value
                
                NumeroTelefono = .Fields("TelefonoNumero").Value
                If NumeroTelefono <> CALLERID_NOTAVAILABLE And NumeroTelefono <> CALLERID_PRIVATE And Len(NumeroTelefono) > Len(pTelephony.LocationCityCode) Then
                    If Left(NumeroTelefono, Len(pTelephony.LocationCityCode)) = pTelephony.LocationCityCode Then
                        NumeroTelefono = Mid(NumeroTelefono, Len(pTelephony.LocationCityCode) + 1)
                    End If
                End If
                ListItem.SubItems(5) = NumeroTelefono
                'Call ResolverEntidad(ListItem)

                .MoveNext
            Loop
            
            stbMain.Panels("TEXT").Text = .RecordCount & " items" & IIf(.RecordCount = .MaxRecords, " (Limitados)", "") & " "
        Else
            stbMain.Panels("TEXT").Text = "No hay items. "
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
    ShowErrorMessage "Forms.RegistroLlamada.FillListView", "Error al obtener la lista de Llamadas Telefónicas."
End Function

Private Sub cboShowTop_Click()
    FillListView
End Sub

Private Sub cboSucursal_Click()
    FillListView
End Sub

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    ResizeControls NewHeight
End Sub

Private Sub cmdMostrarPersona_Click()
    If lvwData.SelectedItem Is Nothing Then
        MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
        lvwData.SetFocus
        Exit Sub
    End If
    
    Call ResolverEntidad(lvwData.SelectedItem)
    
    lvwData.SetFocus
End Sub

Private Sub cmdMostrarPersonas_Click()
    Dim ListViewItem As MSComctlLib.ListItem
    
    For Each ListViewItem In lvwData.ListItems
        If Not ResolverEntidad(ListViewItem) Then
            Exit For
        End If
    Next ListViewItem
    
    lvwData.SetFocus
End Sub

Private Sub dtpFechaDesde_Change()
    FillListView
End Sub

Private Sub cmdAnteriorDesde_Click()
    dtpFechaDesde.Value = DateAdd("d", -1, dtpFechaDesde.Value)
    dtpFechaDesde.SetFocus
    dtpFechaDesde_Change
End Sub

Private Sub cmdSiguienteDesde_Click()
    dtpFechaDesde.Value = DateAdd("d", 1, dtpFechaDesde.Value)
    dtpFechaDesde.SetFocus
    dtpFechaDesde_Change
End Sub

Private Sub cmdHoyDesde_Click()
    Dim OldValue As Date
    
    OldValue = dtpFechaDesde.Value
    dtpFechaDesde.Value = Date
    dtpFechaDesde.SetFocus
    If OldValue <> dtpFechaDesde.Value Then
        dtpFechaDesde_Change
    End If
End Sub

Private Sub Form_Load()
    mLoading = True
    
    lvwData.GridLines = pParametro.ListView_GridLines
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL LISTVIEW
    Set lvwData.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    '//////////////////////////////////////////////////////////
    
    '//////////////////////////////////////////////////////////
    'ASIGNO LOS ICONOS AL PIN
    Set tlbPin.ImageList = frmMDI.ilsFormPin
    '//////////////////////////////////////////////////////////
    
    dtpFechaDesde.Value = Date
    
    Call FillComboBoxSucursal
    cboSucursal.ListIndex = 0
    
    cboShowTop.AddItem "10"
    cboShowTop.AddItem "50"
    cboShowTop.AddItem "100"
    cboShowTop.AddItem "500"
    cboShowTop.AddItem "1000"
    cboShowTop.AddItem "5000"
    cboShowTop.ListIndex = 1
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    pParametro.GetCoolBarSettings "RegistroLlamada", cbrMain
    pParametro.GetListViewSettings "RegistroLlamada", lvwData
    lvwData.ColumnHeaders(lvwData.SortKey + 1).Icon = lvwData.SortOrder + 1
    tlbPin.Buttons("PIN").Value = pParametro.Usuario_LeerNumero("RegistroLlamada_Pin", tlbPin.Buttons("PIN").Value)
    If tlbPin.Buttons("PIN").Value = tbrUnpressed Then
        tlbPin.Buttons("PIN").Image = 1
    Else
        tlbPin.Buttons("PIN").Image = 2
    End If
    
    mLoading = False
    
    FillListView
End Sub

Private Sub Form_Resize()
    ResizeControls cbrMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveCoolBarSettings "RegistroLlamada", cbrMain
    pParametro.SaveListViewSettings "RegistroLlamada", lvwData
    pParametro.Usuario_GuardarNumero "RegistroLlamada_Pin", tlbPin.Buttons("PIN").Value
    
    Set mCSucursal = Nothing
    
    Set frmRegistroLlamada = Nothing
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
    FillListView
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

Private Sub FillComboBoxSucursal()
    Dim KeySave As String
    
    If cboSucursal.ListIndex > 0 Then
        KeySave = mCSucursal(cboSucursal.ListIndex)
    End If
    
    cboSucursal.Clear
    Set mCSucursal = New Collection
    cboSucursal.AddItem "« Todas »"
    
    If CSM_Control_ComboBox.AndCollection_FillFromSQL(cboSucursal, mCSucursal, False, "SELECT IDSucursal, Nombre FROM Sucursal WHERE Activo = 1 ORDER BY Nombre", "IDSucursal", "Nombre", "Sucursales", cscpFirstIfUnique) Then
        cboSucursal.ListIndex = CSM_Control_ComboBox.AndCollection_GetListIndexByCollectionItem(cboSucursal, mCSucursal, KeySave, cscpFirstIfUnique)
    Else
        cboSucursal.ListIndex = 0
    End If
End Sub

Private Function ResolverEntidad(ByRef ListViewItem As MSComctlLib.ListItem) As Boolean
    Dim cmdEntidad As ADODB.command
    Dim recEntidad As ADODB.Recordset
    Dim NumeroCompleto As String
    
    Dim errorMessage As String
    Dim TelefonoIndex As Long
    Dim IDTelefonoTipo As Long
    
    ListViewItem.SubItems(6) = ""
    ListViewItem.SubItems(7) = ""
    
    If ListViewItem.SubItems(3) = CALLERID_NOTAVAILABLE Or ListViewItem.SubItems(2) = CALLERID_PRIVATE Then
        ResolverEntidad = True
        Exit Function
    End If

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    NumeroCompleto = IIf(Left(ListViewItem.SubItems(5), 1) <> "0", pTelephony.LocationCityCode & ListViewItem.SubItems(5), ListViewItem.SubItems(5))
    
    errorMessage = "Error al obtener la Lista de Entidades por el Número de Teléfono."
    
    Set cmdEntidad = New ADODB.command
    Set cmdEntidad.ActiveConnection = pDatabase.Connection
    cmdEntidad.CommandText = "sp_Persona_CallerID_Search"
    cmdEntidad.CommandType = adCmdStoredProc
    cmdEntidad.Parameters.Append cmdEntidad.CreateParameter("@TelefonoAreaLocal", adVarChar, adParamInput, 5, pTelephony.LocationCityCode)
    cmdEntidad.Parameters.Append cmdEntidad.CreateParameter("@TelefonoNumero", adVarChar, adParamInput, 21, NumeroCompleto)
    Set recEntidad = New ADODB.Recordset
    recEntidad.Open cmdEntidad, , adOpenForwardOnly, adLockReadOnly
    Set cmdEntidad = Nothing
    
    errorMessage = "Error al leer las Entidades según el Número de Teléfono."
    Do While Not recEntidad.EOF
        ListViewItem.SubItems(7) = ListViewItem.SubItems(7) & IIf(ListViewItem.SubItems(7) = "", "", " // ") & recEntidad("Persona").Value
        
        'Busco el Tipo de Teléfono
        Select Case NumeroCompleto
            Case IIf(IsNull(recEntidad("Telefono1Area").Value), pTelephony.LocationCityCode, recEntidad("Telefono1Area").Value) & recEntidad("Telefono1Numero").Value
                TelefonoIndex = 1
            Case IIf(IsNull(recEntidad("Telefono2Area").Value), pTelephony.LocationCityCode, recEntidad("Telefono2Area").Value) & recEntidad("Telefono2Numero").Value
                TelefonoIndex = 2
            Case IIf(IsNull(recEntidad("Telefono3Area").Value), pTelephony.LocationCityCode, recEntidad("Telefono3Area").Value) & recEntidad("Telefono3Numero").Value
                TelefonoIndex = 3
            Case IIf(IsNull(recEntidad("Telefono4Area").Value), pTelephony.LocationCityCode, recEntidad("Telefono4Area").Value) & recEntidad("Telefono4Numero").Value
                TelefonoIndex = 4
            Case IIf(IsNull(recEntidad("Telefono5Area").Value), pTelephony.LocationCityCode, recEntidad("Telefono5Area").Value) & recEntidad("Telefono5Numero").Value
                TelefonoIndex = 5
        End Select
        errorMessage = "Ha ocurrido un error al leer el Tipo de Teléfono."
        
        IDTelefonoTipo = Val(recEntidad("IDTelefono" & TelefonoIndex & "Tipo").Value & "")
        If IDTelefonoTipo <> 0 Then
            If IsNull(recEntidad("Telefono" & TelefonoIndex & "TipoOtro").Value) Then
                ListViewItem.SubItems(6) = ListViewItem.SubItems(6) & IIf(ListViewItem.SubItems(6) = "", "", " // ") & recEntidad("Telefono" & TelefonoIndex & "TipoNombre").Value & ""
            Else
                ListViewItem.SubItems(6) = ListViewItem.SubItems(6) & IIf(ListViewItem.SubItems(6) = "", "", " // ") & recEntidad("Telefono" & TelefonoIndex & "TipoOtro").Value & ""
            End If
        End If
        
        errorMessage = "Error al leer las Entidades según el Número de Teléfono."
        recEntidad.MoveNext
    Loop
    recEntidad.Close
    Set recEntidad = Nothing
    
    ResolverEntidad = True
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Forms.RegistroLlamada.ResolverEntidad", errorMessage
End Function
