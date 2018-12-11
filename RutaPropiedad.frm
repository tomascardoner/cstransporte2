VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRutaPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RutaPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   9390
   Begin VB.CheckBox chkPermite2Conductores 
      Alignment       =   1  'Right Justify
      Caption         =   "Permite 2 Conductores:"
      Height          =   210
      Left            =   4860
      TabIndex        =   25
      Top             =   1020
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.Frame fraConductorImporte 
      Caption         =   "Importe a acreditar al conductor:"
      Height          =   1695
      Left            =   4860
      TabIndex        =   26
      Top             =   1380
      Width           =   4395
      Begin VB.TextBox txtConductorImporteTramo2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   32
         Tag             =   "CURRENCY|EMPTY|ZERO|POSITIVE"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtConductorImporteTramo1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   30
         Tag             =   "CURRENCY|EMPTY|ZERO|POSITIVE"
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox txtConductorImporteTramoCompleto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   28
         Tag             =   "CURRENCY|EMPTY|ZERO|POSITIVE"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblConductorImporteTramo2 
         AutoSize        =   -1  'True
         Caption         =   "Tramo 2:"
         Height          =   210
         Left            =   180
         TabIndex        =   31
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label lblConductorImporteTramo1 
         AutoSize        =   -1  'True
         Caption         =   "Tramo 1:"
         Height          =   210
         Left            =   180
         TabIndex        =   29
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblConductorImporteTramoCompleto 
         AutoSize        =   -1  'True
         Caption         =   "Tramo completo:"
         Height          =   210
         Left            =   180
         TabIndex        =   27
         Top             =   420
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdRutaGrupo 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4260
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Grupos de Rutas"
      Top             =   2700
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fraLimiteCancelacion 
      Caption         =   "Límite de Cancelación de Reservas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   120
      TabIndex        =   18
      Top             =   4020
      Width           =   4395
      Begin VB.TextBox txtLimiteCancelacionDuracion 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   900
         MaxLength       =   5
         TabIndex        =   23
         Top             =   720
         Width           =   795
      End
      Begin VB.CommandButton cmdLimiteCancelacionLugar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4020
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Lugares"
         Top             =   300
         Width           =   255
      End
      Begin MSDataListLib.DataCombo datcboLimiteCancelacionLugar 
         Height          =   330
         Left            =   900
         TabIndex        =   20
         Top             =   300
         Width           =   3075
         _ExtentX        =   5424
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
      Begin VB.Label lblLimiteCancelacionDuracion 
         AutoSize        =   -1  'True
         Caption         =   "Duración:"
         Height          =   210
         Left            =   120
         TabIndex        =   22
         Top             =   780
         Width           =   690
      End
      Begin VB.Label lblLimiteCancelacionDuracionMinutos 
         AutoSize        =   -1  'True
         Caption         =   "minutos"
         Height          =   210
         Left            =   1800
         TabIndex        =   24
         Top             =   780
         Width           =   555
      End
      Begin VB.Label lblLimiteCancelacionLugar 
         AutoSize        =   -1  'True
         Caption         =   "Lugar:"
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   900
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1380
      Width           =   3615
   End
   Begin VB.TextBox txtDuracion 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   900
      MaxLength       =   5
      TabIndex        =   16
      Top             =   3540
      Width           =   795
   End
   Begin VB.CommandButton cmdLugarDestino 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4260
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Lugares"
      Top             =   2220
      Width           =   255
   End
   Begin VB.CommandButton cmdLugarOrigen 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4260
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Lugares"
      Top             =   1860
      Width           =   255
   End
   Begin VB.TextBox txtKilometro 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   900
      MaxLength       =   4
      TabIndex        =   14
      Top             =   3180
      Width           =   795
   End
   Begin VB.TextBox txtNotas 
      Height          =   945
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   34
      Top             =   3240
      Width           =   3615
   End
   Begin VB.CheckBox chkActivo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Activo:"
      Height          =   195
      Left            =   4860
      TabIndex        =   35
      Top             =   4500
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox txtIDRuta 
      Height          =   315
      Left            =   900
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1020
      Width           =   2775
   End
   Begin VB.Frame fraLine 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   120
      TabIndex        =   39
      Top             =   780
      Width           =   9135
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8040
      TabIndex        =   37
      Top             =   4980
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6720
      TabIndex        =   36
      Top             =   4980
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo datcboOrigen 
      Height          =   330
      Left            =   900
      TabIndex        =   5
      Top             =   1860
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
   Begin MSDataListLib.DataCombo datcboDestino 
      Height          =   330
      Left            =   900
      TabIndex        =   8
      Top             =   2220
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
   Begin MSDataListLib.DataCombo datcboRutaGrupo 
      Height          =   330
      Left            =   900
      TabIndex        =   11
      Top             =   2700
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
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   4680
      Y1              =   900
      Y2              =   5160
   End
   Begin VB.Label lblRutaGrupo 
      AutoSize        =   -1  'True
      Caption         =   "&Grupo:"
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblDuracionMinutos 
      AutoSize        =   -1  'True
      Caption         =   "minutos"
      Height          =   210
      Left            =   1800
      TabIndex        =   17
      Top             =   3600
      Width           =   555
   End
   Begin VB.Label lblDuracion 
      AutoSize        =   -1  'True
      Caption         =   "Duración:"
      Height          =   210
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   690
   End
   Begin VB.Label lblKilometro 
      AutoSize        =   -1  'True
      Caption         =   "&Kms.:"
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label lblDestino 
      AutoSize        =   -1  'True
      Caption         =   "&Destino:"
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   585
   End
   Begin VB.Label lblOrigen 
      AutoSize        =   -1  'True
      Caption         =   "&Origen:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   525
   End
   Begin VB.Label lblNotas 
      AutoSize        =   -1  'True
      Caption         =   "Notas:"
      Height          =   210
      Left            =   4860
      TabIndex        =   33
      Top             =   3300
      Width           =   465
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos de la Ruta"
      Height          =   210
      Left            =   780
      TabIndex        =   38
      Top             =   300
      Width           =   2370
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "RutaPropiedad.frx":054A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblIDLugar 
      AutoSize        =   -1  'True
      Caption         =   "&ID:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   180
   End
End
Attribute VB_Name = "frmRutaPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRuta As Ruta
Private mNew As Boolean
Private mKeyDecimal As Boolean

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef Ruta As Ruta)
    Set mRuta = Ruta
    mNew = (mRuta.IDRuta = "")

    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mRuta
        If Not CSM_Control_DataCombo.FillFromSQL(datcboOrigen, "SELECT IDLugar, Nombre FROM Lugar WHERE IDLugar <> " & pParametro.Lugar_ID_Otro & " AND (Activo = 1 OR IDLugar = " & .IDOrigen & ") ORDER BY Nombre", "IDLugar", "Nombre", "Orígenes") Then
            Unload Me
            Exit Sub
        End If
        If Not CSM_Control_DataCombo.FillFromSQL(datcboDestino, "SELECT IDLugar, Nombre FROM Lugar WHERE IDLugar <> " & pParametro.Lugar_ID_Otro & " AND (Activo = 1 OR IDLugar = " & .IDDestino & ") ORDER BY Nombre", "IDLugar", "Nombre", "Destinos") Then
            Unload Me
            Exit Sub
        End If
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboRutaGrupo, "SELECT IDRutaGrupo, Nombre FROM RutaGrupo WHERE Activo = 1 OR IDRutaGrupo = " & .IDRutaGrupo & " ORDER BY Nombre", "IDRutaGrupo", "Nombre", "Grupos de Rutas") Then
            Unload Me
            Exit Sub
        End If
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboLimiteCancelacionLugar, "(SELECT 1 AS Orden, 0 AS Indice, 0 AS IDLugar, '----------' AS Nombre) UNION (SELECT 2 AS Orden, RutaDetalle.Indice, RutaDetalle.IDLugar, Lugar.Nombre FROM RutaDetalle INNER JOIN Lugar ON RutaDetalle.IDLugar = Lugar.IDLugar WHERE RutaDetalle.IDRuta = '" & ReplaceQuote(mRuta.IDRuta) & "' AND Lugar.Activo = 1) ORDER BY Orden, Indice, Nombre", "IDLugar", "Nombre", "Orígenes", cscpItemOrfirst, Val(datcboLimiteCancelacionLugar.BoundText)) Then
            Unload Me
            Exit Sub
        End If
        
        txtIDRuta.Text = .IDRuta
        txtIDRuta.Enabled = mNew
        txtNombre.Text = .Nombre
        datcboOrigen.BoundText = .IDOrigen
        datcboDestino.BoundText = .IDDestino
        datcboRutaGrupo.BoundText = .IDRutaGrupo
        txtKilometro.Text = IIf(.Kilometro = 0, "", .Kilometro)
        txtDuracion.Text = IIf(.Duracion = 0, "", .Duracion)
        datcboLimiteCancelacionLugar.BoundText = .LimiteCancelacionIDLugar
        txtLimiteCancelacionDuracion.Text = IIf(.LimiteCancelacionDuracion = 0, "", .LimiteCancelacionDuracion)
        
        chkPermite2Conductores.Value = IIf(.Permite2Conductores, vbChecked, vbUnchecked)
        Call chkPermite2Conductores_Click
        txtConductorImporteTramoCompleto.Text = .ConductorImporteTramoCompleto_FormattedAsString
        If pParametro.Viaje_Permite_2_Conductores And chkPermite2Conductores.Value = vbChecked Then
            txtConductorImporteTramo1.Text = .ConductorImporteTramo1_FormattedAsString
            txtConductorImporteTramo2.Text = .ConductorImporteTramo2_FormattedAsString
        End If
        
        txtNotas.Text = .Notas
        chkActivo.Value = IIf(.Activo, vbChecked, vbUnchecked)
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub Form_Load()
    chkPermite2Conductores.Visible = pParametro.Viaje_Permite_2_Conductores
    chkPermite2Conductores.Value = IIf(pParametro.Viaje_Permite_2_Conductores, vbChecked, vbUnchecked)
    Call chkPermite2Conductores_Click
End Sub

Private Sub txtNombre_GotFocus()
    CSM_Control_TextBox.SelAllText txtNombre
End Sub

Private Sub cmdLimiteCancelacionLugar_Click()
    If pCPermiso.GotPermission(PERMISO_LUGAR) Then
        Screen.MousePointer = vbHourglass
        frmLugar.Show
        On Error Resume Next
        Set frmLugar.lvwData.SelectedItem = frmLugar.lvwData.ListItems(KEY_STRINGER & datcboLimiteCancelacionLugar.BoundText)
        frmLugar.lvwData.SelectedItem.EnsureVisible
        If frmLugar.WindowState = vbMinimized Then
            frmLugar.WindowState = vbNormal
        End If
        frmLugar.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdLugarOrigen_Click()
    If pCPermiso.GotPermission(PERMISO_LUGAR) Then
        Screen.MousePointer = vbHourglass
        frmLugar.Show
        On Error Resume Next
        Set frmLugar.lvwData.SelectedItem = frmLugar.lvwData.ListItems(KEY_STRINGER & datcboOrigen.BoundText)
        frmLugar.lvwData.SelectedItem.EnsureVisible
        If frmLugar.WindowState = vbMinimized Then
            frmLugar.WindowState = vbNormal
        End If
        frmLugar.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdLugarDestino_Click()
    If pCPermiso.GotPermission(PERMISO_LUGAR) Then
        Screen.MousePointer = vbHourglass
        frmLugar.Show
        On Error Resume Next
        Set frmLugar.lvwData.SelectedItem = frmLugar.lvwData.ListItems(KEY_STRINGER & datcboDestino.BoundText)
        frmLugar.lvwData.SelectedItem.EnsureVisible
        If frmLugar.WindowState = vbMinimized Then
            frmLugar.WindowState = vbNormal
        End If
        frmLugar.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtIDRuta_Change()
    Caption = "Propiedades" & IIf(Trim(txtIDRuta.Text) = "", "", " de " & txtIDRuta.Text)
End Sub

Private Sub txtIDRuta_GotFocus()
    CSM_Control_TextBox.SelAllText txtIDRuta
End Sub

Private Sub txtKilometro_GotFocus()
    CSM_Control_TextBox.SelAllText txtKilometro
End Sub

Private Sub txtKilometro_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtKilometro_LostFocus()
    txtKilometro.Text = Val(txtKilometro.Text)
    If txtKilometro.Text = 0 Then
        txtKilometro.Text = ""
    End If
End Sub

Private Sub txtDuracion_GotFocus()
    CSM_Control_TextBox.SelAllText txtDuracion
End Sub

Private Sub txtDuracion_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDuracion_LostFocus()
    txtDuracion.Text = Val(txtDuracion.Text)
    If txtDuracion.Text = 0 Then
        txtDuracion.Text = ""
    End If
End Sub

Private Sub txtLimiteCancelacionDuracion_GotFocus()
    CSM_Control_TextBox.SelAllText txtLimiteCancelacionDuracion
End Sub

Private Sub txtLimiteCancelacionDuracion_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLimiteCancelacionDuracion_LostFocus()
    txtLimiteCancelacionDuracion.Text = Val(txtLimiteCancelacionDuracion.Text)
    If txtLimiteCancelacionDuracion.Text = 0 Then
        txtLimiteCancelacionDuracion.Text = ""
    End If
End Sub

Private Sub chkPermite2Conductores_Click()
    Call ShowControls
End Sub

Private Sub txtConductorImporteTramoCompleto_GotFocus()
    CSM_Control_TextBox.SelAllText txtConductorImporteTramoCompleto
End Sub

Private Sub txtConductorImporteTramoCompleto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtConductorImporteTramoCompleto_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(txtConductorImporteTramoCompleto, KeyAscii, mKeyDecimal)
End Sub

Private Sub txtConductorImporteTramoCompleto_LostFocus()
    Call CSM_Control_TextBox.FormatValue_ByTag(txtConductorImporteTramoCompleto)
End Sub

Private Sub txtConductorImporteTramo1_GotFocus()
    CSM_Control_TextBox.SelAllText txtConductorImporteTramo1
End Sub

Private Sub txtConductorImporteTramo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtConductorImporteTramo1_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(txtConductorImporteTramo1, KeyAscii, mKeyDecimal)
End Sub

Private Sub txtConductorImporteTramo1_LostFocus()
    Call CSM_Control_TextBox.FormatValue_ByTag(txtConductorImporteTramo1)
End Sub

Private Sub txtConductorImporteTramo2_GotFocus()
    CSM_Control_TextBox.SelAllText txtConductorImporteTramo2
End Sub

Private Sub txtConductorImporteTramo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDecimal Then
        mKeyDecimal = True
    End If
End Sub

Private Sub txtConductorImporteTramo2_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(txtConductorImporteTramo2, KeyAscii, mKeyDecimal)
End Sub

Private Sub txtConductorImporteTramo2_LostFocus()
    Call CSM_Control_TextBox.FormatValue_ByTag(txtConductorImporteTramo2)
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
End Sub

Private Sub cmdOK_Click()
    If Trim(txtIDRuta.Text) = "" Then
        MsgBox "Debe ingresar el ID de la Ruta.", vbInformation, App.Title
        txtIDRuta.SetFocus
        Exit Sub
    End If
    If Trim(txtNombre.Text) = "" Then
        MsgBox "Debe ingresar el Nombre de la Ruta.", vbInformation, App.Title
        txtNombre.SetFocus
        Exit Sub
    End If
    If Val(datcboOrigen.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Origen de la Ruta.", vbInformation, App.Title
        datcboOrigen.SetFocus
        Exit Sub
    End If
    If Val(datcboDestino.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Destino de la Ruta.", vbInformation, App.Title
        datcboDestino.SetFocus
        Exit Sub
    End If
    If (datcboOrigen.BoundText = datcboDestino.BoundText) Then
        MsgBox "El Origen debe ser diferente del Destino de la Ruta.", vbInformation, App.Title
        datcboDestino.SetFocus
        Exit Sub
    End If
    If Val(datcboRutaGrupo.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo de la Ruta.", vbInformation, App.Title
        datcboRutaGrupo.SetFocus
        Exit Sub
    End If
    
    Call txtConductorImporteTramoCompleto_LostFocus
    If pParametro.Viaje_Permite_2_Conductores Then
        Call txtConductorImporteTramo1_LostFocus
        Call txtConductorImporteTramo2_LostFocus
    End If
    
    With mRuta
        .IDRuta = txtIDRuta.Text
        .Nombre = txtNombre.Text
        .IDOrigen = Val(datcboOrigen.BoundText)
        .IDDestino = Val(datcboDestino.BoundText)
        .IDRutaGrupo = Val(datcboRutaGrupo.BoundText)
        .Kilometro = Val(txtKilometro.Text)
        .Duracion = Val(txtDuracion.Text)
        .LimiteCancelacionIDLugar = Val(datcboLimiteCancelacionLugar.BoundText)
        .LimiteCancelacionDuracion = Val(txtLimiteCancelacionDuracion.Text)
        .ConductorImporteTramoCompleto_FormattedAsString = txtConductorImporteTramoCompleto.Text
        .Permite2Conductores = (chkPermite2Conductores.Value = vbChecked)
        If pParametro.Viaje_Permite_2_Conductores And chkPermite2Conductores.Value = vbChecked Then
            .ConductorImporteTramo1_FormattedAsString = txtConductorImporteTramo1.Text
            .ConductorImporteTramo2_FormattedAsString = txtConductorImporteTramo2.Text
        Else
            .ConductorImporteTramo1 = -1
            .ConductorImporteTramo2 = -1
        End If
        .Notas = txtNotas.Text
        .Activo = (chkActivo.Value = vbChecked)
        If Not .Update() Then
            Exit Sub
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mRuta = Nothing
    Set frmRutaPropiedad = Nothing
End Sub

Public Sub FillComboBoxLugar()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboOrigen.BoundText)
    Set recData = datcboOrigen.RowSource
    recData.Requery
    Set recData = Nothing
    datcboOrigen.BoundText = KeySave

    KeySave = Val(datcboDestino.BoundText)
    Set recData = datcboDestino.RowSource
    recData.Requery
    Set recData = Nothing
    datcboDestino.BoundText = KeySave
End Sub

Private Sub ShowControls()
    lblConductorImporteTramo1.Visible = (pParametro.Viaje_Permite_2_Conductores And chkPermite2Conductores.Value = vbChecked)
    txtConductorImporteTramo1.Visible = (pParametro.Viaje_Permite_2_Conductores And chkPermite2Conductores.Value = vbChecked)
    lblConductorImporteTramo2.Visible = (pParametro.Viaje_Permite_2_Conductores And chkPermite2Conductores.Value = vbChecked)
    txtConductorImporteTramo2.Visible = (pParametro.Viaje_Permite_2_Conductores And chkPermite2Conductores.Value = vbChecked)
End Sub
