VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVehiculoMantenimientoPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "VehiculoMantenimientoPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   3660
      Picture         =   "VehiculoMantenimientoPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   60
      Width           =   975
   End
   Begin VB.OptionButton optTipoFecha 
      Caption         =   "Fecha"
      Height          =   195
      Left            =   2940
      TabIndex        =   7
      Top             =   1920
      Width           =   795
   End
   Begin VB.CheckBox chkActivo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Activo"
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
      Left            =   120
      TabIndex        =   15
      Top             =   4380
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox txtPreaviso 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1020
      MaxLength       =   7
      TabIndex        =   12
      Top             =   2700
      Width           =   1035
   End
   Begin VB.OptionButton optTipoNinguno 
      Caption         =   "Ninguno"
      Height          =   195
      Left            =   3780
      TabIndex        =   8
      Top             =   1920
      Width           =   915
   End
   Begin VB.OptionButton optTipoDias 
      Caption         =   "Días"
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   1920
      Width           =   675
   End
   Begin VB.OptionButton optTipoKilometraje 
      Caption         =   "Kilometros"
      Height          =   195
      Left            =   1020
      TabIndex        =   5
      Top             =   1920
      Value           =   -1  'True
      Width           =   1035
   End
   Begin VB.TextBox txtLapso 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1020
      MaxLength       =   7
      TabIndex        =   10
      Top             =   2280
      Width           =   1035
   End
   Begin VB.CommandButton cmdVehiculoMantenimientoGrupo 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4380
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Grupos"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton cmdVehiculo 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4380
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Vehículos"
      Top             =   1020
      Width           =   255
   End
   Begin VB.TextBox txtNotas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   3300
      Width           =   3615
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   21
      Top             =   780
      Width           =   4515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3420
      TabIndex        =   17
      Top             =   4980
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      TabIndex        =   16
      Top             =   4980
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo datcboVehiculo 
      Height          =   330
      Left            =   1020
      TabIndex        =   1
      Top             =   1020
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
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
   Begin MSDataListLib.DataCombo datcboVehiculoMantenimientoGrupo 
      Height          =   330
      Left            =   1020
      TabIndex        =   3
      Top             =   1440
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
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
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   1020
      TabIndex        =   24
      Top             =   2280
      Visible         =   0   'False
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
      Format          =   16580609
      CurrentDate     =   36950
   End
   Begin VB.Label lblPreavisoUnidad 
      AutoSize        =   -1  'True
      Caption         =   "kilómetros antes."
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
      Left            =   2220
      TabIndex        =   23
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblPreaviso 
      AutoSize        =   -1  'True
      Caption         =   "Aviso:"
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
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   465
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   345
   End
   Begin VB.Label lblLapsoUnidad 
      AutoSize        =   -1  'True
      Caption         =   "kilómetros."
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
      Left            =   2220
      TabIndex        =   22
      Top             =   2340
      Width           =   765
   End
   Begin VB.Label lblLapso 
      AutoSize        =   -1  'True
      Caption         =   "Cada:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   2340
      Width           =   420
   End
   Begin VB.Label lblVehiculoMantenimientoGrupo 
      AutoSize        =   -1  'True
      Caption         =   "Grupo:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1500
      Width           =   495
   End
   Begin VB.Label lblVehiculo 
      AutoSize        =   -1  'True
      Caption         =   "&Vehículo:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label lblNotas 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
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
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   465
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Datos del Mantenimiento del Vehículo"
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
      Left            =   780
      TabIndex        =   20
      Top             =   300
      Width           =   2670
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "VehiculoMantenimientoPropiedad.frx":0B74
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmVehiculoMantenimientoPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mVehiculoMantenimiento As VehiculoMantenimiento
Private mNew As Boolean


Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mVehiculoMantenimiento
End Sub

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef VehiculoMantenimiento As VehiculoMantenimiento)
    Set mVehiculoMantenimiento = VehiculoMantenimiento
    mNew = (mVehiculoMantenimiento.IDVehiculo = 0)
    
    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mVehiculoMantenimiento
        If Not CSM_Control_DataCombo.FillFromSQL(datcboVehiculo, "SELECT IDVehiculo, Nombre FROM Vehiculo WHERE Activo = 1 OR IDVehiculo = " & .IDVehiculo & " ORDER BY Nombre", "IDVehiculo", "Nombre", "Vehículos", cscpItemOrFirst, .IDVehiculo) Then
            Unload Me
            Exit Sub
        End If
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboVehiculoMantenimientoGrupo, "SELECT IDVehiculoMantenimientoGrupo, Nombre FROM VehiculoMantenimientoGrupo WHERE Activo = 1 OR IDVehiculoMantenimientoGrupo = " & .IDVehiculoMantenimientoGrupo & " ORDER BY Nombre", "IDVehiculoMantenimientoGrupo", "Nombre", "Grupos de Mantenimiento de Vehículos", cscpItemOrFirst, .IDVehiculoMantenimientoGrupo) Then
            Unload Me
            Exit Sub
        End If
        
        dtpFecha.Value = Date
        Select Case .Tipo
            Case VEHICULO_MATENIMIENTO_TIPO_KILOMETRAJE
                optTipoKilometraje.Value = True
                ShowControls
                txtLapso.Text = .KilometrajeLapso
                dtpFecha.Value = Date
                txtPreaviso.Text = .KilometrajePreaviso
            Case VEHICULO_MATENIMIENTO_TIPO_DIAS
                optTipoDias.Value = True
                ShowControls
                txtLapso.Text = .DiasLapso
                dtpFecha.Value = Date
                txtPreaviso.Text = .DiasPreaviso
            Case VEHICULO_MATENIMIENTO_TIPO_FECHA
                optTipoFecha.Value = True
                ShowControls
                txtLapso.Text = ""
                dtpFecha.Value = .FechaFecha
                txtPreaviso.Text = .FechaPreaviso
            Case VEHICULO_MATENIMIENTO_TIPO_NINGUNO
                optTipoNinguno.Value = True
                ShowControls
                txtLapso.Text = ""
                dtpFecha.Value = Date
                txtPreaviso.Text = ""
        End Select
        txtNotas.Text = .Notas
        chkActivo.Value = IIf(.Activo, vbChecked, vbUnchecked)
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Sub FillComboBoxVehiculo()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboVehiculo.BoundText)
    Set recData = datcboVehiculo.RowSource
    recData.Requery
    Set recData = Nothing
    datcboVehiculo.BoundText = KeySave
End Sub

Public Sub FillComboBoxVehiculoMantenimientoGrupo()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboVehiculoMantenimientoGrupo.BoundText)
    Set recData = datcboVehiculoMantenimientoGrupo.RowSource
    recData.Requery
    Set recData = Nothing
    datcboVehiculoMantenimientoGrupo.BoundText = KeySave
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Val(datcboVehiculo.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Vehículo.", vbInformation, App.Title
        datcboVehiculo.SetFocus
        Exit Sub
    End If
    If Val(datcboVehiculoMantenimientoGrupo.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo.", vbInformation, App.Title
        datcboVehiculoMantenimientoGrupo.SetFocus
        Exit Sub
    End If
    
    If optTipoKilometraje.Value Or optTipoDias.Value Then
        If Not IsNumeric(txtLapso.Text) Then
            MsgBox "El Lapso debe ser un valor numérico.", vbInformation, App.Title
            txtLapso.SetFocus
            Exit Sub
        End If
        If CLng(txtLapso.Text) <= 0 Then
            MsgBox "El Lapso debe ser mayor a cero.", vbInformation, App.Title
            txtLapso.SetFocus
            Exit Sub
        End If
    End If
    If optTipoKilometraje.Value Or optTipoDias.Value Or optTipoFecha.Value Then
        If Not IsNumeric(txtPreaviso.Text) Then
            MsgBox "El Aviso debe ser un valor numérico.", vbInformation, App.Title
            txtPreaviso.SetFocus
            Exit Sub
        End If
        If CLng(txtPreaviso.Text) < 0 Then
            MsgBox "El Aviso debe ser mayor o igual a cero.", vbInformation, App.Title
            txtPreaviso.SetFocus
            Exit Sub
        End If
    End If
    If optTipoKilometraje.Value Or optTipoDias.Value Then
        If CLng(txtPreaviso.Text) >= CLng(txtLapso.Text) Then
            MsgBox "El Aviso debe ser menor al Lapso.", vbInformation, App.Title
            txtPreaviso.SetFocus
            Exit Sub
        End If
    End If
    
    With mVehiculoMantenimiento
        .IDVehiculo = Val(datcboVehiculo.BoundText)
        .IDVehiculoMantenimientoGrupo = Val(datcboVehiculoMantenimientoGrupo.BoundText)
        If optTipoKilometraje.Value Then
            .Tipo = VEHICULO_MATENIMIENTO_TIPO_KILOMETRAJE
            .KilometrajeLapso = CLng(txtLapso.Text)
            .KilometrajePreaviso = CLng(txtPreaviso.Text)
        ElseIf optTipoDias.Value Then
            .Tipo = VEHICULO_MATENIMIENTO_TIPO_DIAS
            .DiasLapso = CLng(txtLapso.Text)
            .DiasPreaviso = CLng(txtPreaviso.Text)
        ElseIf optTipoFecha.Value Then
            .Tipo = VEHICULO_MATENIMIENTO_TIPO_FECHA
            .FechaFecha = dtpFecha.Value
            .FechaPreaviso = CLng(txtPreaviso.Text)
        ElseIf optTipoNinguno.Value Then
            .Tipo = VEHICULO_MATENIMIENTO_TIPO_NINGUNO
        End If
        .Notas = txtNotas.Text
        .Activo = (chkActivo.Value = vbChecked)
        If mNew Then
            If Not .AddNew() Then
                Exit Sub
            End If
        Else
            If Not .Update Then
                Exit Sub
            End If
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdVehiculo_Click()
    If pCPermiso.GotPermission(PERMISO_VEHICULO) Then
        Screen.MousePointer = vbHourglass
        frmVehiculo.Show
        On Error Resume Next
        Set frmVehiculo.lvwData.SelectedItem = frmVehiculo.lvwData.ListItems(KEY_STRINGER & datcboVehiculo.BoundText)
        frmVehiculo.lvwData.SelectedItem.EnsureVisible
        If frmVehiculo.WindowState = vbMinimized Then
            frmVehiculo.WindowState = vbNormal
        End If
        frmVehiculo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdVehiculoMantenimientoGrupo_Click()
    If pCPermiso.GotPermission(PERMISO_VEHICULO_MANTENIMIENTO_GRUPO) Then
        Screen.MousePointer = vbHourglass
        frmVehiculoMantenimientoGrupo.Show
        On Error Resume Next
        Set frmVehiculoMantenimientoGrupo.lvwData.SelectedItem = frmVehiculoMantenimientoGrupo.lvwData.ListItems(KEY_STRINGER & datcboVehiculoMantenimientoGrupo.BoundText)
        frmVehiculoMantenimientoGrupo.lvwData.SelectedItem.EnsureVisible
        If frmVehiculoMantenimientoGrupo.WindowState = vbMinimized Then
            frmVehiculoMantenimientoGrupo.WindowState = vbNormal
        End If
        frmVehiculoMantenimientoGrupo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mVehiculoMantenimiento = Nothing
    Set frmVehiculoMantenimientoPropiedad = Nothing
End Sub

Private Sub optTipoKilometraje_Click()
    ShowControls
End Sub

Private Sub optTipoDias_Click()
    ShowControls
End Sub

Private Sub optTipoFecha_Click()
    ShowControls
End Sub

Private Sub optTipoNinguno_Click()
    ShowControls
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
End Sub

Private Sub txtLapso_GotFocus()
    CSM_Control_TextBox.SelAllText txtLapso
End Sub

Private Sub txtLapso_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLapso_LostFocus()
    txtLapso.Text = Val(txtLapso.Text)
    If txtLapso.Text = 0 Then
        txtLapso.Text = ""
    End If
End Sub

Private Sub txtPreaviso_GotFocus()
    CSM_Control_TextBox.SelAllText txtPreaviso
End Sub

Private Sub txtPreaviso_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPreaviso_LostFocus()
    txtPreaviso.Text = Val(txtPreaviso.Text)
End Sub

Private Sub ShowControls()
    lblLapso.Visible = (optTipoKilometraje.Value Or optTipoDias.Value Or optTipoFecha.Value)
    txtLapso.Visible = (optTipoKilometraje.Value Or optTipoDias.Value)
    txtLapso.Text = ""
    lblLapsoUnidad.Visible = (optTipoKilometraje.Value Or optTipoDias.Value)
    
    dtpFecha.Visible = (optTipoFecha.Value)
    
    lblPreaviso.Visible = (optTipoKilometraje.Value Or optTipoDias.Value Or optTipoFecha.Value)
    txtPreaviso.Visible = (optTipoKilometraje.Value Or optTipoDias.Value Or optTipoFecha.Value)
    txtPreaviso.Text = ""
    lblPreavisoUnidad.Visible = (optTipoKilometraje.Value Or optTipoDias.Value Or optTipoFecha.Value)
    
    If optTipoKilometraje.Value Then
        lblLapso.Caption = "Cada:"
        lblLapsoUnidad.Caption = "kilómetros."
        lblPreavisoUnidad.Caption = "kilómetros antes."
    ElseIf optTipoDias.Value Then
        lblLapso.Caption = "Cada:"
        lblLapsoUnidad.Caption = "días."
        lblPreavisoUnidad.Caption = "días antes."
    ElseIf optTipoFecha.Value Then
        lblLapso.Caption = "Fecha:"
        lblPreavisoUnidad.Caption = "días antes."
    End If
End Sub
