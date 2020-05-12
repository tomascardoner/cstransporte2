VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListaPasajero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Lista de Pasajeros"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ListaPasajero.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   6990
   Begin VB.CheckBox chkBorrarMarcasPersonas 
      Caption         =   "Quitar marcas actuales"
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   5640
      Value           =   1  'Checked
      Width           =   2235
   End
   Begin VB.CheckBox chkMarcarPersonas 
      Caption         =   "Marcar las Personas"
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   5940
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelectNone 
      Caption         =   "&Ninguno"
      Height          =   375
      Left            =   4620
      TabIndex        =   15
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Todos"
      Height          =   375
      Left            =   3540
      TabIndex        =   14
      Top             =   1140
      Width           =   975
   End
   Begin VB.ListBox lstWeekday 
      Height          =   1740
      ItemData        =   "ListaPasajero.frx":000C
      Left            =   180
      List            =   "ListaPasajero.frx":000E
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   1500
      Width           =   1275
   End
   Begin MSComCtl2.UpDown udCantidadPasajero 
      Height          =   315
      Left            =   3435
      TabIndex        =   2
      Top             =   180
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtCantidadPasajero"
      BuddyDispid     =   196614
      OrigLeft        =   3180
      OrigTop         =   120
      OrigRight       =   3420
      OrigBottom      =   855
      Max             =   600
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtCantidadPasajero 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2940
      MaxLength       =   3
      TabIndex        =   1
      Top             =   180
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4260
      TabIndex        =   12
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5580
      TabIndex        =   13
      Top             =   5700
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFechaDesde 
      Height          =   315
      Left            =   2940
      TabIndex        =   4
      Top             =   660
      Width           =   1635
      _ExtentX        =   2884
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
      Format          =   127270913
      CurrentDate     =   36950
      MaxDate         =   73050
      MinDate         =   36526
   End
   Begin MSComCtl2.DTPicker dtpFechaHasta 
      Height          =   315
      Left            =   5160
      TabIndex        =   5
      Top             =   660
      Width           =   1635
      _ExtentX        =   2884
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
      Format          =   127270913
      CurrentDate     =   36950
      MaxDate         =   73050
      MinDate         =   36526
   End
   Begin MSComctlLib.ListView lvwRutaHorario 
      Height          =   3915
      Left            =   2100
      TabIndex        =   9
      Top             =   1560
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6906
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "IDRuta"
         Text            =   "Ruta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Hora"
         Text            =   "Hora"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblHorario 
      AutoSize        =   -1  'True
      Caption         =   "Rutas - Horarios:"
      Height          =   210
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   1230
   End
   Begin VB.Label lblWeekday 
      AutoSize        =   -1  'True
      Caption         =   "Días de la Semana:"
      Height          =   210
      Left            =   180
      TabIndex        =   6
      Top             =   1200
      Width           =   1380
   End
   Begin VB.Label lblRange 
      AutoSize        =   -1  'True
      Caption         =   "Pasajeros que más viajaron entre el                                            y el "
      Height          =   210
      Left            =   180
      TabIndex        =   3
      Top             =   720
      Width           =   4860
   End
   Begin VB.Label lblCantidadPasajero 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad de Pasajeros a Listar:"
      Height          =   210
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   2250
   End
End
Attribute VB_Name = "frmListaPasajero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub FillListView()
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim KeySave As Variant
    Dim CKeySave As Collection
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    
    Screen.MousePointer = vbHourglass
    
    If Not lvwRutaHorario.SelectedItem Is Nothing Then
        KeySave = lvwRutaHorario.SelectedItem.Key
        Set CKeySave = New Collection
        For Each ListItem In lvwRutaHorario.ListItems
            If ListItem.Checked Then
                CKeySave.Add ListItem.Key
            End If
        Next ListItem
    End If
    
    lvwRutaHorario.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    SQL_Where = ""
    
    If pPersonal Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Horario.Personal = 0"
    End If
    
    If pCPermiso.RutaWhere <> "" Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Horario")
    End If
    
    Select Case lvwRutaHorario.SortKey
        Case 0  'RUTA
            SQL_OrderBy = " ORDER BY IDRuta" & IIf(lvwRutaHorario.SortOrder = lvwAscending, "", " DESC") & ", Hora" & IIf(lvwRutaHorario.SortOrder = lvwAscending, "", " DESC")
        Case 1  'HORA
            SQL_OrderBy = " ORDER BY Hora" & IIf(lvwRutaHorario.SortOrder = lvwAscending, "", " DESC") & ", IDRuta" & IIf(lvwRutaHorario.SortOrder = lvwAscending, "", " DESC")
    End Select
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT DISTINCT IDRuta, Hora FROM Horario" & SQL_Where & SQL_OrderBy
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwRutaHorario.ListItems.Add(, KEY_STRINGER & RTrim(.Fields("IDRuta").Value) & KEY_DELIMITER & Format(.Fields("Hora").Value, "hh:nn:ss"), RTrim(.Fields("IDRuta").Value))
                ListItem.SubItems(1) = Format(.Fields("Hora").Value, "Short Time")
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set recData = Nothing
    
    On Error Resume Next
    lvwRutaHorario.SelectedItem.Selected = False
    Set lvwRutaHorario.SelectedItem = lvwRutaHorario.ListItems(KeySave)
    lvwRutaHorario.SelectedItem.EnsureVisible
    
    If Not CKeySave Is Nothing Then
        If CKeySave.Count > 1 Then
            For Each KeySave In CKeySave
                lvwRutaHorario.ListItems(KeySave).Checked = True
            Next KeySave
        End If
    End If
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hwnd And frmMDI.WindowState <> vbMinimized Then
        lvwRutaHorario.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.ListaPasajero.FillListView", "Error al obtener la lista de Horarios."
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim ExcelApplication As Object
    Dim ExcelWorkbook As Object
    Dim ExcelWorksheet As Object
    Dim recData As ADODB.Recordset
    Dim Index As Integer
    Dim ListItem As MSComctlLib.ListItem
    Dim AllSelected As Boolean
    Dim WeekdayWhere As String
    Dim RutaHorarioWhere As String
    Dim Persona As Persona
    
    If DateDiff("d", dtpFechaDesde.Value, dtpFechaHasta.Value) < 0 Then
        MsgBox "La Fecha Desde debe ser menor o igual a la Fecha Hasta.", vbInformation, App.Title
        dtpFechaHasta.SetFocus
        Exit Sub
    End If
    
    AllSelected = True
    For Index = 1 To 7
        If lstWeekday.Selected(Index - 1) Then
            WeekdayWhere = WeekdayWhere & IIf(WeekdayWhere = "", "", " OR ") & "DATEPART(weekday, Viaje.FechaHora) = " & Index
        Else
            AllSelected = False
        End If
    Next Index
    If WeekdayWhere = "" Then
        MsgBox "Debe seleccionar al menos un Día de la Semana.", vbInformation, App.Title
        lstWeekday.SetFocus
        Exit Sub
    End If
    If AllSelected Then
        WeekdayWhere = ""
    Else
        WeekdayWhere = "AND (" & WeekdayWhere & ") "
    End If
    
    AllSelected = True
    For Each ListItem In lvwRutaHorario.ListItems
        If ListItem.Checked Then
            RutaHorarioWhere = RutaHorarioWhere & IIf(RutaHorarioWhere = "", "", " OR ") & "(Viaje.IDRuta = '" & ReplaceQuote(ListItem.Text) & "' AND CONVERT(char(8), Viaje.FechaHora, 108) = '" & ListItem.SubItems(1) & ":00')"
        Else
            AllSelected = False
        End If
    Next ListItem
    If RutaHorarioWhere = "" Then
        MsgBox "Debe seleccionar al menos una Ruta y Horario.", vbInformation, App.Title
        lvwRutaHorario.SetFocus
        Exit Sub
    Else
        RutaHorarioWhere = "AND (" & RutaHorarioWhere & ") "
    End If
    If AllSelected Then
        RutaHorarioWhere = ""
    Else
        
    End If
    
    If MsgBox("¿Desea generar la Lista de Pasajeros?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If chkBorrarMarcasPersonas.Value = vbChecked Then
        pDatabase.Connection.Execute "UPDATE Persona SET ListaPasajero = 0 WHERE ListaPasajero = 1", 0
    End If
    
    'INICIO UNA SESION DE EXCEL
    Set ExcelApplication = CreateObject("Excel.Application")
    ExcelApplication.Visible = Not pIsCompiled
    
    'ABRO EL ARCHIVO ESPECIFICADO
    Set ExcelWorkbook = ExcelApplication.Workbooks.Open(pParametro.Report_Path & pParametro.ListaPasajero_ArchivoNombre, , True)
    Set ExcelWorksheet = ExcelWorkbook.Worksheets(1)
    
    'ABRO EL RECORDSET
    Set recData = New ADODB.Recordset
    Set recData.ActiveConnection = pDatabase.Connection
    With recData
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "SELECT Persona.IDPersona, Persona.Apellido + (CASE ISNULL(Persona.Nombre, '') WHEN '' THEN '' ELSE ', ' + Persona.Nombre END) AS Persona, DocumentoTipo.Nombre + ': ' + Persona.DocumentoNumero AS Documento" & vbCr
        .Source = .Source & "FROM (Persona INNER JOIN DocumentoTipo ON Persona.IDDocumentoTipo = DocumentoTipo.IDDocumentoTipo) INNER JOIN" & vbCr
        .Source = .Source & "(SELECT TOP " & Val(txtCantidadPasajero.Text) & " ViajeDetalle.IDPersona, COUNT(ViajeDetalle.FechaHora) AS CantidadViajes" & vbCr
        .Source = .Source & "FROM (Persona INNER JOIN ViajeDetalle ON Persona.IDPersona = ViajeDetalle.IDPersona) INNER JOIN Viaje ON ViajeDetalle.FechaHora = Viaje.FechaHora AND ViajeDetalle.IDRuta = Viaje.IDRuta" & vbCr
        .Source = .Source & "WHERE Viaje.Estado <> 'CA' AND ViajeDetalle.Estado = '1CO' AND ViajeDetalle.OcupanteTipo = 'PA' AND Persona.EntidadTipo = 'PC' AND ViajeDetalle.Realizado = 1 AND Persona.IDDocumentoTipo IS NOT NULL AND Persona.DocumentoNumero IS NOT NULL" & vbCr
        .Source = .Source & "AND convert(char(10), Viaje.FechaHora, 111) >= '" & Format(dtpFechaDesde.Value, "yyyy/mm/dd") & "'" & vbCr
        .Source = .Source & "AND convert(char(10), Viaje.FechaHora, 111) <= '" & Format(dtpFechaHasta.Value, "yyyy/mm/dd") & "'" & vbCr
        If pPersonal Then
            .Source = .Source & "AND Viaje.Personal = 0" & vbCr
        End If
        .Source = .Source & WeekdayWhere & RutaHorarioWhere
        .Source = .Source & "GROUP BY ViajeDetalle.IDPersona" & vbCr
        .Source = .Source & "ORDER BY CantidadViajes DESC) AS Viajes ON Persona.IDPersona = Viajes.IDPersona" & vbCr
        .Source = .Source & "ORDER BY Persona" & vbCr
        .Open , , , , adCmdText
        
        Do While Not .EOF
            ExcelWorksheet.Range(pParametro.ListaPasajero_ColumnPasajero & (pParametro.ListaPasajero_RowStart + .AbsolutePosition - 1)).Value = .Fields("Persona").Value
            ExcelWorksheet.Range(pParametro.ListaPasajero_ColumnDocumento & (pParametro.ListaPasajero_RowStart + .AbsolutePosition - 1)).Value = .Fields("Documento").Value
            
            If chkMarcarPersonas.Value = vbChecked Then
                Set Persona = New Persona
                Persona.IDPersona = .Fields("IDPersona").Value
                If Persona.Load() Then
                    Persona.ListaPasajero = True
                    Persona.Update
                End If
                Set Persona = Nothing
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    Set recData = Nothing
    
    ExcelApplication.Visible = True
    
    'LIBERO LOS OBJETOS
    Set ExcelWorksheet = Nothing
    Set ExcelWorkbook = Nothing
    Set ExcelApplication = Nothing
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Se ha Generado la Lista de Pasajeros.", vbInformation, App.Title
    
    Unload Me
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case 429
            Screen.MousePointer = vbDefault
            MsgBox "No se puede iniciar una sesión de Microsoft Excel." & vbCr & "Reinstale Microsoft Excel.", vbCritical, App.Title
        Case 1004
            Screen.MousePointer = vbDefault
            MsgBox "Error interno de Microsoft Excel.", vbCritical, App.Title
        Case Else
            ShowErrorMessage "Forms.ListaPasajero.OK", "Error al Generar la Lista de Pasajeros."
    End Select
    If Not ExcelWorksheet Is Nothing Then
        Set ExcelWorksheet = Nothing
    End If
    If Not ExcelWorkbook Is Nothing Then
        Set ExcelWorkbook = Nothing
    End If
    If Not ExcelApplication Is Nothing Then
        ExcelApplication.Quit
        Set ExcelApplication = Nothing
    End If
    If Not recData Is Nothing Then
        If recData.State = adStateOpen Then
            recData.Close
        End If
        Set recData = Nothing
    End If
End Sub

Private Sub cmdSelectAll_Click()
    Call SelectListViewItems(True)
End Sub

Private Sub cmdSelectNone_Click()
    Call SelectListViewItems(False)
End Sub

Private Sub Form_Load()
    Dim Index As Integer
    Dim ListItem As MSComctlLib.ListItem
    
    txtCantidadPasajero.Text = 1
    dtpFechaDesde.Value = DateAdd("m", -3, Date)
    dtpFechaHasta.Value = Date
    
    For Index = 1 To 7
        lstWeekday.AddItem WeekdayName(Index)
        lstWeekday.Selected(Index - 1) = True
    Next Index
    lstWeekday.ListIndex = -1
    
    lvwRutaHorario.GridLines = pParametro.ListView_GridLines
    Set lvwRutaHorario.ColumnHeaderIcons = frmMDI.ilsFormSortColumn
    pParametro.GetListViewSettings "ListaPasajero", lvwRutaHorario
    lvwRutaHorario.ColumnHeaders(lvwRutaHorario.SortKey + 1).Icon = lvwRutaHorario.SortOrder + 1
    
    Call FillListView
    For Each ListItem In lvwRutaHorario.ListItems
        ListItem.Checked = True
    Next ListItem
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveListViewSettings "ListaPasajero", lvwRutaHorario
    Set frmListaPasajero = Nothing
End Sub

Private Sub lvwRutaHorario_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwRutaHorario.ColumnHeaders(lvwRutaHorario.SortKey + 1).Icon = 0
    If ColumnHeader.Index - 1 = lvwRutaHorario.SortKey Then
        lvwRutaHorario.SortOrder = IIf(lvwRutaHorario.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwRutaHorario.SortKey = ColumnHeader.Index - 1
        lvwRutaHorario.SortOrder = lvwAscending
    End If
    ColumnHeader.Icon = lvwRutaHorario.SortOrder + 1
    Call FillListView
End Sub

Private Sub txtCantidadPasajero_GotFocus()
    CSM_Control_TextBox.SelAllText txtCantidadPasajero
End Sub

Private Sub txtCantidadPasajero_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCantidadPasajero_LostFocus()
    txtCantidadPasajero.Text = Val(txtCantidadPasajero.Text)
    If txtCantidadPasajero.Text > 600 Then
        txtCantidadPasajero.Text = 600
    End If
    If txtCantidadPasajero.Text = 0 Then
        txtCantidadPasajero.Text = ""
    End If
End Sub

Private Sub udCantidadPasajero_Change()
    txtCantidadPasajero_GotFocus
End Sub

Private Sub SelectListViewItems(ByVal Value As Boolean)
    Dim ListItem As MSComctlLib.ListItem
    
    lvwRutaHorario.Visible = False
    
    For Each ListItem In lvwRutaHorario.ListItems
        ListItem.Checked = Value
    Next ListItem
    
    lvwRutaHorario.Visible = True
End Sub
