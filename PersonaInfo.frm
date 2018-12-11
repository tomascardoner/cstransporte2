VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonaInfo 
   Caption         =   "Información"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PersonaInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   9570
   Begin VB.CommandButton cmdAbrirReporte 
      Caption         =   "Abrir Reporte"
      Height          =   330
      Left            =   6600
      TabIndex        =   3
      Top             =   60
      Width           =   1575
   End
   Begin VB.ComboBox cboDiasPasado 
      Height          =   330
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin MSComctlLib.ListView lvwData 
      Height          =   5355
      Left            =   240
      TabIndex        =   1
      Top             =   540
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   9446
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
         Key             =   "DiaSemana"
         Text            =   "Día"
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
         Key             =   "IDRuta"
         Text            =   "Ruta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "ListaPrecio"
         Text            =   "Lista de Precios"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Estado"
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   6195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   10927
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Viajes Pasados"
            Key             =   "PASADO"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Viajes del Día"
            Key             =   "PRESENTE"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Viajes Futuros"
            Key             =   "FUTURO"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   6375
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16351
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
Attribute VB_Name = "frmPersonaInfo"
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
    
    If Not FillListView(mPersona.IDPersona) Then
        Unload Me
        Exit Sub
    End If
    
    Caption = "Información de la Persona: " & mPersona.IDPersona & " - " & mPersona.ApellidoNombre

    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Sub ForceRefresh()
    FillListView mPersona.IDPersona
End Sub
    
Public Function FillListView(ByVal IDPersona As Long) As Boolean
    Dim MousePointerSave As Integer
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim SQL_Where As String
    Dim SQL_OrderBy As String
    Dim Viaje As Viaje
    Dim ViajeDetalle As ViajeDetalle
    
    If mLoading Then
        Exit Function
    End If
    
    MousePointerSave = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    SQL_Where = ""
    
    If pPersonal Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Viaje.Personal = 0"
    End If
    
    SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "ViajeDetalle.IDPersona = " & IDPersona & " AND ViajeDetalle.OcupanteTipo = '" & OCUPANTE_TIPO_PASAJERO & "'"
    
    Select Case tabMain.SelectedItem.Key
        Case "PASADO"
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "convert(char(10), Viaje.FechaHora, 111) + ' ' + convert(char(10), Viaje.FechaHora, 108) >= '" & Format(DateAdd("d", (cboDiasPasado.ListIndex + 1) * -30, Date), "yyyy/mm/dd") & " 00:00:00' AND convert(char(10), Viaje.FechaHora, 111) + ' ' + convert(char(10), Viaje.FechaHora, 108) < '" & Format(Date, "yyyy/mm/dd") & " 00:00:00' AND Viaje.Estado <> '" & VIAJE_ESTADO_CANCELADO & "' AND ViajeDetalle.Estado = '" & VIAJE_DETALLE_ESTADO_CONFIRMADO & "'"
            SQL_OrderBy = " ORDER BY Viaje.FechaHora DESC"
        Case "PRESENTE"
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "convert(char(10), Viaje.FechaHora, 111) + ' ' + convert(char(10), Viaje.FechaHora, 108) >= '" & Format(Date, "yyyy/mm/dd") & " 00:00:00' AND convert(char(10), Viaje.FechaHora, 111) + ' ' + convert(char(10), Viaje.FechaHora, 108) <= '" & Format(Date, "yyyy/mm/dd") & " 23:59:00' AND Viaje.Estado <> '" & VIAJE_ESTADO_CANCELADO & "'"
            SQL_OrderBy = " ORDER BY Viaje.FechaHora"
        Case "FUTURO"
            SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "convert(char(10), Viaje.FechaHora, 111) + ' ' + convert(char(10), Viaje.FechaHora, 108) > '" & Format(Date, "yyyy/mm/dd") & " 23:59:00'"
            SQL_OrderBy = " ORDER BY Viaje.FechaHora"
    End Select
        
    lvwData.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT Viaje.FechaHora, Viaje.IDRuta, ViajeDetalle.Indice, Viaje.Estado AS ViajeEstado, ViajeDetalle.Estado AS ViajeDetalleEstado, ViajeDetalle.Realizado, ListaPrecio.Nombre AS ListaPrecioNombre FROM (Viaje INNER JOIN ViajeDetalle ON Viaje.FechaHora = ViajeDetalle.FechaHora AND Viaje.IDRuta = ViajeDetalle.IDRuta) INNER JOIN ListaPrecio ON ViajeDetalle.IDListaPrecio = ListaPrecio.IDListaPrecio" & SQL_Where & SQL_OrderBy
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set Viaje = New Viaje
    Set ViajeDetalle = New ViajeDetalle
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwData.ListItems.Add(, KEY_STRINGER & .Fields("FechaHora").Value & KEY_DELIMITER & RTrim(.Fields("IDRuta").Value) & KEY_DELIMITER & .Fields("Indice").Value, WeekdayName(Weekday(.Fields("FechaHora").Value)))
                ListItem.SubItems(1) = Format(.Fields("FechaHora").Value, "Short Date")
                ListItem.SubItems(2) = Format(.Fields("FechaHora").Value, "Short Time")
                ListItem.SubItems(3) = RTrim(.Fields("IDRuta").Value)
                ListItem.SubItems(4) = .Fields("ListaPrecioNombre").Value
                Select Case tabMain.SelectedItem.Key
                    Case "PASADO"
                        ListItem.SubItems(5) = IIf(IsNull(.Fields("Realizado").Value), "", IIf(.Fields("Realizado").Value, "Realizado", "Ausente"))
                    Case "PRESENTE", "FUTURO"
                        ViajeDetalle.Estado = .Fields("ViajeDetalleEstado").Value & ""
                        ListItem.SubItems(5) = ViajeDetalle.Estado_ToString
                End Select
                Viaje.Estado = .Fields("ViajeEstado").Value
                ListItem.ForeColor = Viaje.Estado_ToColor
                ListItem.Bold = Viaje.Estado_ToBold
                .MoveNext
            Loop
            
            stbMain.Panels("TEXT").Text = .RecordCount & " items"
        Else
            stbMain.Panels("TEXT").Text = "No hay items."
        End If
        .Close
    End With
    Set recData = Nothing
    
    Set Viaje = Nothing
    Set ViajeDetalle = Nothing
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hWnd And frmMDI.WindowState <> vbMinimized Then
        lvwData.SetFocus
    End If
    
    Screen.MousePointer = MousePointerSave
    
    FillListView = True
    Exit Function
    
ErrorHandler:
    Screen.MousePointer = MousePointerSave
    CSM_Error.ShowErrorMessage "Forms.PersonaInfo.FillListView", "Error al obtener la información de la Persona." & vbCr & vbCr & "IDPersona: " & mPersona.IDPersona
End Function

Private Sub cboDiasPasado_Click()
    FillListView mPersona.IDPersona
End Sub

Private Sub cmdAbrirReporte_Click()
    Dim ReporteParametro As ReporteParametro
    
    If pCPermiso.GotPermission(PERMISO_REPORTE) Then
        If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & "Persona_Viaje_Listado") Then
            Screen.MousePointer = vbHourglass
            
            'SELECCIONO EL REPORTE
            Set frmReporte.lvwReport.SelectedItem = frmReporte.lvwReport.ListItems(KEY_STRINGER & "Persona_Viaje_Listado")
            frmReporte.cmdNext_Click
            
            'CARGO EL PARÁMETRO DE PERSONA
            Set ReporteParametro = frmReporte.mReporte.Parametros("IDPersona")
            ReporteParametro.Valor = mPersona.IDPersona
            ReporteParametro.ValorLeyenda = mPersona.ApellidoNombre
            frmReporte.lvwParameter.ListItems(KEY_STRINGER & "IDPersona").SubItems(2) = ReporteParametro.ValorLeyenda
            Set ReporteParametro = Nothing
            
            'CARGO LOS PARÁMETROS DE FECHA DESDE Y HASTA
            Select Case tabMain.SelectedItem.Key
                Case "PASADO"
                    Set ReporteParametro = frmReporte.mReporte.Parametros("FechaDesde")
                    ReporteParametro.Valor = Format(DateAdd("d", (cboDiasPasado.ListIndex + 1) * -30, Date), "Short Date")
                    ReporteParametro.ValorLeyenda = ReporteParametro.Valor
                    frmReporte.lvwParameter.ListItems(KEY_STRINGER & "FechaDesde").SubItems(2) = ReporteParametro.ValorLeyenda
                    Set ReporteParametro = Nothing
                    
                    Set ReporteParametro = frmReporte.mReporte.Parametros("FechaHasta")
                    ReporteParametro.Valor = Format(DateAdd("d", -1, Date), "Short Date")
                    ReporteParametro.ValorLeyenda = ReporteParametro.Valor
                    frmReporte.lvwParameter.ListItems(KEY_STRINGER & "FechaHasta").SubItems(2) = ReporteParametro.ValorLeyenda
                    Set ReporteParametro = Nothing
                Case "PRESENTE"
                    Set ReporteParametro = frmReporte.mReporte.Parametros("FechaDesde")
                    ReporteParametro.Valor = Format(Date, "Short Date")
                    ReporteParametro.ValorLeyenda = ReporteParametro.Valor
                    frmReporte.lvwParameter.ListItems(KEY_STRINGER & "FechaDesde").SubItems(2) = ReporteParametro.ValorLeyenda
                    Set ReporteParametro = Nothing
                    
                    Set ReporteParametro = frmReporte.mReporte.Parametros("FechaHasta")
                    ReporteParametro.Valor = Format(Date, "Short Date")
                    ReporteParametro.ValorLeyenda = ReporteParametro.Valor
                    frmReporte.lvwParameter.ListItems(KEY_STRINGER & "FechaHasta").SubItems(2) = ReporteParametro.ValorLeyenda
                    Set ReporteParametro = Nothing
                Case "FUTURO"
                    Set ReporteParametro = frmReporte.mReporte.Parametros("FechaDesde")
                    ReporteParametro.Valor = Format(DateAdd("d", 1, Date), "Short Date")
                    ReporteParametro.ValorLeyenda = ReporteParametro.Valor
                    frmReporte.lvwParameter.ListItems(KEY_STRINGER & "FechaDesde").SubItems(2) = ReporteParametro.ValorLeyenda
                    Set ReporteParametro = Nothing
            End Select
    
            'MUESTRO LA PANTALLA
            frmReporte.Show
            If frmReporte.WindowState = vbMinimized Then
                frmReporte.WindowState = vbNormal
            End If
            frmReporte.lvwParameter.SetFocus
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub Form_Load()
    lvwData.GridLines = pParametro.ListView_GridLines
    
    CSM_Forms.ResizeAndPosition frmMDI, Me
    
    mLoading = True
    
    tabMain.SelectedItem = tabMain.Tabs("PRESENTE")
    
    cboDiasPasado.AddItem "Ultimos 30 días (1 mes)"
    cboDiasPasado.AddItem "Ultimos 60 días (2 meses)"
    cboDiasPasado.AddItem "Ultimos 90 días (3 meses)"
    cboDiasPasado.AddItem "Ultimos 120 días (4 meses)"
    cboDiasPasado.AddItem "Ultimos 150 días (5 meses)"
    cboDiasPasado.AddItem "Ultimos 180 días (6 meses)"
    cboDiasPasado.AddItem "Ultimos 210 días (7 meses)"
    cboDiasPasado.AddItem "Ultimos 240 días (8 meses)"
    cboDiasPasado.AddItem "Ultimos 270 días (9 meses)"
    cboDiasPasado.AddItem "Ultimos 300 días (10 meses)"
    cboDiasPasado.AddItem "Ultimos 330 días (11 meses)"
    cboDiasPasado.AddItem "Ultimos 360 días (12 meses)"
    cboDiasPasado.ListIndex = 0
    
    mLoading = False
End Sub

Private Sub ResizeControls()
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
    
    tabMain.Top = CONTROL_SPACE
    tabMain.Left = CONTROL_SPACE
    tabMain.Width = ScaleWidth - (CONTROL_SPACE * 2)
    tabMain.Height = ScaleHeight - tabMain.Top - CONTROL_SPACE - stbMain.Height
    
    cboDiasPasado.Top = CONTROL_SPACE
    
    lvwData.Top = tabMain.ClientTop + CONTROL_SPACE
    lvwData.Left = tabMain.ClientLeft + CONTROL_SPACE
    lvwData.Width = tabMain.ClientWidth - (CONTROL_SPACE * 2)
    lvwData.Height = tabMain.ClientHeight - (CONTROL_SPACE * 2)
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
    WindowState = vbNormal
    pParametro.SaveListViewSettings "PersonaInformacion", lvwData
    Set mPersona = Nothing
    Set frmPersonaInfo = Nothing
End Sub

Private Sub lvwData_DblClick()
    Dim Viaje As Viaje
    
    If pCPermiso.GotPermission(PERMISO_VIAJE) Then
        If lvwData.SelectedItem Is Nothing Then
            MsgBox "No hay ningún Item seleccionado.", vbInformation, App.Title
            lvwData.SetFocus
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        
        Set Viaje = New Viaje
        Viaje.FechaHora = CDate(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 1, KEY_DELIMITER))
        Viaje.IDRuta = CSM_String.GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 2, KEY_DELIMITER)
        If Viaje.Load() Then
            frmViajeDetalle.LoadDataAndShow Viaje
            On Error Resume Next
            Set frmViajeDetalle.lvwData.SelectedItem = frmViajeDetalle.lvwData.ListItems(KEY_STRINGER & Val(GetSubString(Mid(lvwData.SelectedItem.Key, Len(KEY_STRINGER) + 1), 3, KEY_DELIMITER)))
        Else
            lvwData.SetFocus
        End If
        Set Viaje = Nothing
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub tabMain_Click()
    FillListView mPersona.IDPersona
    cboDiasPasado.Visible = (tabMain.SelectedItem.Key = "PASADO")
End Sub
