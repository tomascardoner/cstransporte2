VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMensajePropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MensajePropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   5820
   Begin VB.ComboBox cboRepetirVeces 
      Height          =   330
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4560
      Width           =   810
   End
   Begin VB.CheckBox chkActivo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Activo"
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   5100
      Value           =   1  'Checked
      Width           =   1755
   End
   Begin VB.CommandButton cmdGrupo 
      Caption         =   "..."
      Height          =   315
      Left            =   5400
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Grupos"
      Top             =   3180
      Width           =   255
   End
   Begin VB.TextBox txtMensaje 
      Height          =   1725
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1260
      Width           =   5535
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
      TabIndex        =   15
      Top             =   780
      Width           =   5595
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   5700
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFechaInicio 
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   3660
      Width           =   1755
      _ExtentX        =   3096
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
      CheckBox        =   -1  'True
      Format          =   16646145
      CurrentDate     =   36950
      MaxDate         =   73050
      MinDate         =   36526
   End
   Begin MSDataListLib.DataCombo datcboGrupo 
      Height          =   330
      Left            =   1680
      TabIndex        =   3
      Top             =   3180
      Width           =   3675
      _ExtentX        =   6482
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
   Begin MSComCtl2.DTPicker dtpFechaFin 
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   4080
      Width           =   1755
      _ExtentX        =   3096
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
      CheckBox        =   -1  'True
      Format          =   16646145
      CurrentDate     =   36950
      MaxDate         =   73050
      MinDate         =   36526
   End
   Begin VB.Label lblRepetirVeces 
      AutoSize        =   -1  'True
      Caption         =   "Repeticiones:"
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   4620
      Width           =   975
   End
   Begin VB.Label lblFechaFin 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Fin:"
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   4140
      Width           =   750
   End
   Begin VB.Label lblGrupo 
      AutoSize        =   -1  'True
      Caption         =   "&Grupo de Usuarios:"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1410
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      Caption         =   "Texto del Mensaje:"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   1350
   End
   Begin VB.Label lblFechaInicio 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicio:"
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   900
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese aquí los Datos del Mensaje"
      Height          =   210
      Left            =   780
      TabIndex        =   14
      Top             =   300
      Width           =   2505
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "MensajePropiedad.frx":054A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMensajePropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMensaje As Mensaje
Private mNew As Boolean

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef Mensaje As Mensaje)
    Set mMensaje = Mensaje

    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mMensaje
        txtMensaje.Text = .Mensaje
        
        If Not CSM_Control_DataCombo.FillFromSQL(datcboGrupo, "(SELECT 0 AS IDUsuarioGrupo, '" & ITEM_NONE_CHARS20 & "' AS Nombre FROM UsuarioGrupo) UNION (SELECT IDUsuarioGrupo, Nombre FROM UsuarioGrupo WHERE IDUsuarioGrupo <> " & USUARIO_GRUPO_ID_ADMINISTRADORES & " AND (Activo = 1 OR IDUsuarioGrupo = " & .IDUsuarioGrupo & ")) ORDER BY Nombre", "IDUsuarioGrupo", "Nombre", "Grupos de Usuarios", cscpItemOrFirst, .IDUsuarioGrupo) Then
            Unload Me
            Exit Sub
        End If
        
        dtpFechaInicio.Value = IIf(.FechaInicio = DATE_TIME_FIELD_NULL_VALUE, Null, .FechaInicio)
        dtpFechaFin.Value = IIf(.FechaFin = DATE_TIME_FIELD_NULL_VALUE, Null, .FechaFin)
        
        cboRepetirVeces.ListIndex = .RepetirVeces - 1
        
        chkActivo.Value = IIf(.Activo, vbChecked, vbUnchecked)
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub Form_Load()
    Dim Index As Byte
    
    dtpFechaInicio.Value = Date
    dtpFechaInicio.Value = Null
    dtpFechaFin.Value = Date
    dtpFechaFin.Value = Null
    
    For Index = 1 To 250
        cboRepetirVeces.AddItem Index
    Next Index
End Sub

Private Sub txtMensaje_GotFocus()
    CSM_Control_TextBox.SelAllText txtMensaje
    cmdOK.Default = False
End Sub

Private Sub txtMensaje_LostFocus()
    cmdOK.Default = True
End Sub

Private Sub cmdGrupo_Click()
    If pCPermiso.GotPermission(PERMISO_USUARIO_GRUPO) Then
        Screen.MousePointer = vbHourglass
        frmUsuarioGrupo.Show
        On Error Resume Next
        Set frmUsuarioGrupo.lvwData.SelectedItem = frmUsuarioGrupo.lvwData.ListItems(KEY_STRINGER & datcboGrupo.BoundText)
        frmUsuarioGrupo.lvwData.SelectedItem.EnsureVisible
        If frmUsuarioGrupo.WindowState = vbMinimized Then
            frmUsuarioGrupo.WindowState = vbNormal
        End If
        frmUsuarioGrupo.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdOK_Click()
    If Trim(txtMensaje.Text) = "" Then
        MsgBox "Debe ingresar el Texto del Mensaje.", vbInformation, App.Title
        txtMensaje.SetFocus
        Exit Sub
    End If
    If (Not IsNull(dtpFechaInicio.Value)) And (Not IsNull(dtpFechaFin.Value)) Then
        If dtpFechaInicio.Value > dtpFechaFin.Value Then
            MsgBox "La Fecha de Inicio debe ser menor o igual a la Fecha de Fin.", vbInformation, App.Title
            dtpFechaFin.SetFocus
            Exit Sub
        End If
    End If
    
    With mMensaje
        .Mensaje = txtMensaje.Text
        .IDUsuarioGrupo = Val(datcboGrupo.BoundText & "")
        .FechaInicio = IIf(IsNull(dtpFechaInicio.Value), DATE_TIME_FIELD_NULL_VALUE, dtpFechaInicio.Value)
        .FechaFin = IIf(IsNull(dtpFechaFin.Value), DATE_TIME_FIELD_NULL_VALUE, dtpFechaFin.Value)
        .RepetirVeces = cboRepetirVeces.ListIndex + 1
        .Activo = (chkActivo.Value = vbChecked)
        
        If Not .Update Then
            Exit Sub
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mMensaje = Nothing
    Set frmMensajePropiedad = Nothing
End Sub

Public Sub FillComboBoxUsuarioGrupo()
    Dim KeySave As Long
    Dim recData As ADODB.Recordset
    
    KeySave = Val(datcboGrupo.BoundText)
    Set recData = datcboGrupo.RowSource
    recData.Requery
    Set recData = Nothing
    datcboGrupo.BoundText = KeySave
End Sub
