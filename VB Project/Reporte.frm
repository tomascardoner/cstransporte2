VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReporte 
   Caption         =   "Reportes"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Reporte.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   8655
   Begin VB.CheckBox chkPrinterSetupBeforeShow 
      Caption         =   "Configurar Impresión"
      Height          =   210
      Left            =   3540
      TabIndex        =   7
      Top             =   5460
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtReportTitle 
      Height          =   315
      Left            =   660
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   7635
   End
   Begin VB.CommandButton cmdShow 
      Cancel          =   -1  'True
      Caption         =   "&Mostrar"
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Anterior"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Siguiente"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwReport 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   8281
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Tipo"
         Text            =   "Tipo"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Descripcion"
         Text            =   "Descripcion"
         Object.Width           =   10583
      EndProperty
   End
   Begin MSComctlLib.ListView lvwParameter 
      Height          =   4695
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   8281
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
         Key             =   "Parametro"
         Text            =   "Parametro"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Tipo"
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Valor"
         Text            =   "Valor"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblReportTitle 
      AutoSize        =   -1  'True
      Caption         =   "Título:"
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   4980
      Visible         =   0   'False
      Width           =   420
   End
End
Attribute VB_Name = "frmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mReporte As Reporte

Public Sub FillListViewReport()
    Dim ListItem As MSComctlLib.ListItem
    Dim recData As ADODB.Recordset
    Dim SQL_Where As String
    
    Screen.MousePointer = vbHourglass
    
    lvwReport.ListItems.Clear
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    SQL_Where = " WHERE MostrarEnVisor = 1"
    
    'PERSONAL
    If pPersonal Then
        SQL_Where = SQL_Where & IIf(SQL_Where = "", " WHERE ", " AND ") & "Personal = 0"
    End If
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT IDReporte, Tipo, Nombre FROM Reporte" & SQL_Where & " ORDER BY Tipo, Nombre"
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With recData
        If Not .EOF Then
            Do While Not .EOF
                Set ListItem = lvwReport.ListItems.Add(, KEY_STRINGER & RTrim(.Fields("IDReporte").Value), .Fields("Tipo").Value)
                ListItem.SubItems(1) = .Fields("Nombre").Value
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set recData = Nothing
    
    On Error Resume Next
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hWnd And frmMDI.WindowState <> vbMinimized Then
        lvwReport.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.Report.FillListView", "Error al obtener la Lista de Reportes."
End Sub

Private Function FillListViewParameter() As Boolean
    Dim ListItem As MSComctlLib.ListItem
    Dim ReporteParametro As ReporteParametro
    
    Screen.MousePointer = vbHourglass
    
    lvwParameter.ListItems.Clear
    
    For Each ReporteParametro In mReporte.Parametros
        If ReporteParametro.Tipo <> REPORTE_PARAMETRO_TIPO_COMPANY And ReporteParametro.Tipo <> REPORTE_PARAMETRO_TIPO_TITLE And ReporteParametro.Tipo <> REPORTE_PARAMETRO_TIPO_CONDITION_TEXT And ReporteParametro.Tipo <> REPORTE_PARAMETRO_TIPO_PERSONAL Then
            Set ListItem = lvwParameter.ListItems.Add(, KEY_STRINGER & ReporteParametro.IDParametro, ReporteParametro.Nombre & ":")
            ListItem.SubItems(1) = IIf(ReporteParametro.Requerido, "Requerido", "Opcional")
        End If
    Next ReporteParametro
    
    On Error Resume Next
    
    If frmMDI.ActiveForm.Name = Me.Name And GetForegroundWindow() = frmMDI.hWnd And frmMDI.WindowState <> vbMinimized Then
        lvwParameter.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
    FillListViewParameter = True
End Function

Public Sub cmdNext_Click()
    If lvwReport.SelectedItem Is Nothing Then
        MsgBox "No hay ningún Reporte seleccionado.", vbInformation, App.Title
        lvwReport.SetFocus
        Exit Sub
    End If
    If pCPermiso.GotPermission(PERMISO_REPORTE_REPORTE & Mid(lvwReport.SelectedItem.Key, Len(KEY_STRINGER) + 1)) Then
        mReporte.IDReporte = Mid(lvwReport.SelectedItem.Key, Len(KEY_STRINGER) + 1)
        If mReporte.Load() Then
            If FillListViewParameter() Then
                Caption = "Reportes: " & mReporte.Tipo & " - " & mReporte.Nombre
            
                lvwReport.Visible = False
                lvwParameter.Visible = True
                
                cmdNext.Visible = False
                chkPrinterSetupBeforeShow.Visible = True
                cmdBack.Visible = True
                cmdShow.Visible = True
                cmdShow.Default = True
                
                txtReportTitle.Text = mReporte.Titulo
                lblReportTitle.Visible = (mReporte.Titulo <> "")
                txtReportTitle.Visible = (mReporte.Titulo <> "")
            End If
        End If
    End If
End Sub

Private Sub cmdBack_Click()
    Caption = "Reportes"
    
    lvwReport.Visible = True
    lvwParameter.Visible = False
    
    chkPrinterSetupBeforeShow.Visible = False
    cmdNext.Visible = True
    cmdNext.Default = True
    cmdBack.Visible = False
    cmdShow.Visible = False
    
    lblReportTitle.Visible = False
    txtReportTitle.Visible = False
End Sub

Private Sub cmdShow_Click()
    Dim ReporteParametro As ReporteParametro
    Dim ListItem As MSComctlLib.ListItem
    
    For Each ListItem In lvwParameter.ListItems
        Set ReporteParametro = mReporte.Parametros(Mid(ListItem.Key, Len(KEY_STRINGER) + 1))
        If ReporteParametro.Requerido And IsEmpty(ReporteParametro.Valor) Then
            MsgBox ReporteParametro.RequeridoLeyenda, vbInformation, App.Title
            Set lvwParameter.SelectedItem = ListItem
            lvwParameter.SetFocus
            Exit Sub
        End If
    Next ListItem
    Set ReporteParametro = Nothing
    Set ListItem = Nothing

    mReporte.Titulo = txtReportTitle.Text
    mReporte.PrinterSetupBeforeShow = (chkPrinterSetupBeforeShow.Value = vbChecked)
    If mReporte.OpenReport() Then
        mReporte.PrintReport True
    End If
End Sub

Private Sub Form_Load()
    lvwReport.GridLines = pParametro.ListView_GridLines
    lvwParameter.GridLines = pParametro.ListView_GridLines

    CSM_Forms.ResizeAndPosition frmMDI, Me

    pParametro.GetListViewSettings "Report", lvwReport
    pParametro.GetListViewSettings "ReportParameter", lvwParameter
    
    FillListViewReport
    
    Set mReporte = New Reporte
End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub

Private Sub ResizeControls()
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
    
    lvwReport.Top = CONTROL_SPACE
    lvwReport.Left = CONTROL_SPACE
    lvwReport.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lvwReport.Height = ScaleHeight - lvwReport.Top - CONTROL_SPACE - cmdBack.Height - CONTROL_SPACE
    
    lvwParameter.Top = CONTROL_SPACE
    lvwParameter.Left = CONTROL_SPACE
    lvwParameter.Width = ScaleWidth - (CONTROL_SPACE * 2)
    lvwParameter.Height = lvwReport.Height - CONTROL_SPACE - txtReportTitle.Height
    
    lblReportTitle.Top = lvwParameter.Top + lvwParameter.Height + (CONTROL_SPACE * 2)
    lblReportTitle.Left = CONTROL_SPACE
    txtReportTitle.Top = lvwParameter.Top + lvwParameter.Height + CONTROL_SPACE
    txtReportTitle.Left = lblReportTitle.Left + lblReportTitle.Width + (CONTROL_SPACE * 2)
    txtReportTitle.Width = ScaleWidth - txtReportTitle.Left - CONTROL_SPACE
    
    cmdBack.Top = lvwReport.Top + lvwReport.Height + CONTROL_SPACE
    cmdBack.Left = ScaleWidth - CONTROL_SPACE - cmdNext.Width - CONTROL_SPACE - cmdBack.Width
    cmdNext.Top = cmdBack.Top
    cmdNext.Left = cmdBack.Left + cmdBack.Width + CONTROL_SPACE
    cmdShow.Top = cmdBack.Top
    cmdShow.Left = cmdNext.Left
    
    chkPrinterSetupBeforeShow.Top = cmdBack.Top + CONTROL_SPACE
    chkPrinterSetupBeforeShow.Left = cmdBack.Left - CONTROL_SPACE - chkPrinterSetupBeforeShow.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pParametro.SaveListViewSettings "Report", lvwReport
    pParametro.SaveListViewSettings "ReportParameter", lvwParameter
    Set mReporte = Nothing
    Set frmReporte = Nothing
End Sub

Private Sub lvwParameter_DblClick()
    Dim ReporteParametro As ReporteParametro
    
    If lvwParameter.SelectedItem Is Nothing Then
        MsgBox "No hay ningún Parámetro seleccionado.", vbInformation, App.Title
        lvwParameter.SetFocus
        Exit Sub
    End If
    
    Set ReporteParametro = mReporte.Parametros(Mid(lvwParameter.SelectedItem.Key, Len(KEY_STRINGER) + 1))
    frmReporteParametro.LoadDataAndShow mReporte.Tipo & " - " & mReporte.Nombre, ReporteParametro
    If frmReporteParametro.Tag = "OK" Then
        lvwParameter.SelectedItem.SubItems(2) = ReporteParametro.ValorLeyenda
    End If
    Unload frmReporteParametro
    Set ReporteParametro = Nothing
End Sub

Private Sub lvwParameter_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ReporteParametro As ReporteParametro
    
    If KeyCode = vbKeyDelete And Not lvwReport.SelectedItem Is Nothing Then
        Set ReporteParametro = mReporte.Parametros(Mid(lvwParameter.SelectedItem.Key, Len(KEY_STRINGER) + 1))
        ReporteParametro.Valor = Empty
        lvwParameter.SelectedItem.SubItems(2) = ""
        Set ReporteParametro = Nothing
    End If
End Sub

Private Sub lvwReport_DblClick()
    cmdNext_Click
End Sub

Private Sub txtReportTitle_GotFocus()
    CSM_Control_TextBox.SelAllText txtReportTitle
End Sub
