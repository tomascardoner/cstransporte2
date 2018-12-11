VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRutaDetallePropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RutaDetallePropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   4680
   Begin VB.TextBox txtEspera 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   900
      MaxLength       =   5
      TabIndex        =   12
      Top             =   3060
      Width           =   795
   End
   Begin VB.TextBox txtDuracion 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   900
      MaxLength       =   5
      TabIndex        =   9
      Top             =   2640
      Width           =   795
   End
   Begin VB.CommandButton cmdRuta 
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
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Rutas"
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdLugarGrupo 
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
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Lugares"
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdLugar 
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
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Lugares"
      Top             =   1380
      Width           =   255
   End
   Begin VB.TextBox txtKilometro 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   900
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2220
      Width           =   795
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
      TabIndex        =   20
      Top             =   780
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3300
      TabIndex        =   15
      Top             =   3780
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1980
      TabIndex        =   14
      Top             =   3780
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo datcboLugar 
      Height          =   330
      Left            =   900
      TabIndex        =   3
      Top             =   1380
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
   Begin MSDataListLib.DataCombo datcboLugarGrupo 
      Height          =   330
      Left            =   900
      TabIndex        =   5
      Top             =   1800
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
   Begin MSDataListLib.DataCombo datcboRuta 
      Height          =   330
      Left            =   900
      TabIndex        =   1
      Top             =   960
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
   Begin VB.Label lblEsperaMinutos 
      AutoSize        =   -1  'True
      Caption         =   "minutos"
      Height          =   210
      Left            =   1800
      TabIndex        =   13
      Top             =   3120
      Width           =   555
   End
   Begin VB.Label lblEspera 
      AutoSize        =   -1  'True
      Caption         =   "Espera:"
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   555
   End
   Begin VB.Label lblDuracion 
      AutoSize        =   -1  'True
      Caption         =   "Duración:"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   2700
      Width           =   690
   End
   Begin VB.Label lblDuracionMinutos 
      AutoSize        =   -1  'True
      Caption         =   "minutos"
      Height          =   210
      Left            =   1800
      TabIndex        =   10
      Top             =   2700
      Width           =   555
   End
   Begin VB.Label lblRuta 
      AutoSize        =   -1  'True
      Caption         =   "&Ruta:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   375
   End
   Begin VB.Label lblKilometro 
      AutoSize        =   -1  'True
      Caption         =   "&Kms.:"
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   405
   End
   Begin VB.Label lblLugarGrupo 
      AutoSize        =   -1  'True
      Caption         =   "Grupo:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   1860
      Width           =   495
   End
   Begin VB.Label lblLugar 
      AutoSize        =   -1  'True
      Caption         =   "Lugar:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos del Detalle de la Ruta"
      Height          =   210
      Left            =   780
      TabIndex        =   19
      Top             =   300
      Width           =   3150
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "RutaDetallePropiedad.frx":054A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmRutaDetallePropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRutaDetalle As RutaDetalle
Private mNew As Boolean

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef RutaDetalle As RutaDetalle)
    Set mRutaDetalle = RutaDetalle
    mNew = (mRutaDetalle.IDLugar = 0)

    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mRutaDetalle
        If Not CSM_Control_DataCombo.FillFromSQL(datcboRuta, "SELECT RTRIM(IDRuta) AS IDRuta FROM Ruta WHERE IDRuta <> '" & ReplaceQuote(pParametro.Ruta_ID_Otra) & "' AND Activo = 1" & IIf(pCPermiso.RutaWhere <> "", " AND " & Replace(pCPermiso.RutaWhere, "%TABLENAME%", "Ruta"), "") & " ORDER BY IDRuta", "IDRuta", "IDRuta", "Rutas", cscpItemOrNone, .IDRuta) Then
            Unload Me
            Exit Sub
        End If
        If Not CSM_Control_DataCombo.FillFromSQL(datcboLugar, "SELECT IDLugar, Nombre FROM Lugar WHERE IDLugar <> " & pParametro.Lugar_ID_Otro & " AND (Activo = 1 OR IDLugar = " & .IDLugar & ") ORDER BY Nombre", "IDLugar", "Nombre", "Lugares", cscpItemOrNone, .IDLugar) Then
            Unload Me
            Exit Sub
        End If
        If Not CSM_Control_DataCombo.FillFromSQL(datcboLugarGrupo, "SELECT IDLugarGrupo, Nombre FROM LugarGrupo WHERE IDLugarGrupo <> " & pParametro.LugarGrupo_ID_Otro & " AND (Activo = 1 OR IDLugarGrupo = " & .IDLugarGrupo & ") ORDER BY Nombre", "IDLugarGrupo", "Nombre", "Grupos de Lugares", cscpItemOrNone, .IDLugarGrupo) Then
            Unload Me
            Exit Sub
        End If
        
        txtKilometro.Text = IIf(.Kilometro = -1, "", .Kilometro)
        txtDuracion.Text = IIf(.Duracion = -1, "", .Duracion)
        txtEspera.Text = IIf(.Espera = -1, "", .Espera)
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Public Sub FillComboBoxRuta()
    Dim KeySave As String
    Dim recData As ADODB.Recordset
    
    KeySave = datcboRuta.BoundText
    Set recData = datcboRuta.RowSource
    recData.Requery
    Set recData = Nothing
    datcboRuta.BoundText = KeySave
End Sub

Public Sub FillComboBoxLugar()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboLugar.BoundText)
    Set recData = datcboLugar.RowSource
    recData.Requery
    Set recData = Nothing
    datcboLugar.BoundText = KeySave
End Sub

Public Sub FillComboBoxLugarGrupo()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboLugarGrupo.BoundText)
    Set recData = datcboLugarGrupo.RowSource
    recData.Requery
    Set recData = Nothing
    datcboLugarGrupo.BoundText = KeySave
End Sub

Private Sub cmdRuta_Click()
    If pCPermiso.GotPermission(PERMISO_RUTA) Then
        Screen.MousePointer = vbHourglass
        frmRuta.Show
        On Error Resume Next
        Set frmRuta.lvwData.SelectedItem = frmRuta.lvwData.ListItems(KEY_STRINGER & datcboRuta.BoundText)
        frmRuta.lvwData.SelectedItem.EnsureVisible
        If frmRuta.WindowState = vbMinimized Then
            frmRuta.WindowState = vbNormal
        End If
        frmRuta.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdLugar_Click()
    If pCPermiso.GotPermission(PERMISO_LUGAR) Then
        Screen.MousePointer = vbHourglass
        frmLugar.Show
        On Error Resume Next
        Set frmLugar.lvwData.SelectedItem = frmLugar.lvwData.ListItems(KEY_STRINGER & datcboLugar.BoundText)
        frmLugar.lvwData.SelectedItem.EnsureVisible
        If frmLugar.WindowState = vbMinimized Then
            frmLugar.WindowState = vbNormal
        End If
        frmLugar.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdLugarGrupo_Click()
    If pCPermiso.GotPermission(PERMISO_LUGAR_GRUPO) Then
        Screen.MousePointer = vbHourglass
        frmLugarGrupo.Show
        On Error Resume Next
        Set frmLugarGrupo.lvwData.SelectedItem = frmLugarGrupo.lvwData.ListItems(KEY_STRINGER & datcboLugarGrupo.BoundText)
        frmLugarGrupo.lvwData.SelectedItem.EnsureVisible
        If frmLugarGrupo.WindowState = vbMinimized Then
            frmLugarGrupo.WindowState = vbNormal
        End If
        frmLugarGrupo.SetFocus
        Screen.MousePointer = vbDefault
    End If
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
    If Trim(txtKilometro.Text) <> "" Then
        txtKilometro.Text = Val(txtKilometro.Text)
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
    If Trim(txtDuracion.Text) <> "" Then
        txtDuracion.Text = Val(txtDuracion.Text)
    End If
End Sub

Private Sub txtEspera_GotFocus()
    CSM_Control_TextBox.SelAllText txtEspera
End Sub

Private Sub txtEspera_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 31 And KeyAscii < 48) Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEspera_LostFocus()
    If Trim(txtEspera.Text) <> "" Then
        txtEspera.Text = Val(txtEspera.Text)
    End If
End Sub

Private Sub cmdOK_Click()
    If datcboRuta.BoundText = "" Then
        MsgBox "Debe seleccionar la Ruta.", vbInformation, App.Title
        datcboRuta.SetFocus
        Exit Sub
    End If
    If Val(datcboLugar.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Lugar.", vbInformation, App.Title
        datcboLugar.SetFocus
        Exit Sub
    End If
    If Val(datcboLugarGrupo.BoundText) = 0 Then
        MsgBox "Debe seleccionar el Grupo.", vbInformation, App.Title
        datcboLugarGrupo.SetFocus
        Exit Sub
    End If
    
    With mRutaDetalle
        .IDRuta = datcboRuta.BoundText
        .IDLugar = Val(datcboLugar.BoundText)
        .IDLugarGrupo = Val(datcboLugarGrupo.BoundText)
        .Kilometro = IIf(Trim(txtKilometro.Text) = "", -1, Val(txtKilometro.Text))
        .Duracion = IIf(Trim(txtDuracion.Text) = "", -1, Val(txtDuracion.Text))
        .Espera = IIf(Trim(txtEspera.Text) = "", -1, Val(txtEspera.Text))
        If mNew Then
            If Not .AddNew() Then
                Exit Sub
            End If
        Else
            If Not .Update() Then
                Exit Sub
            End If
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mRutaDetalle = Nothing
    Set frmRutaDetallePropiedad = Nothing
End Sub
