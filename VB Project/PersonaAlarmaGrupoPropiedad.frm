VERSION 5.00
Begin VB.Form frmPersonaAlarmaGrupoPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "PersonaAlarmaGrupoPropiedad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3660
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdAuditoria 
      Caption         =   "Auditoría"
      Height          =   615
      Left            =   3600
      Picture         =   "PersonaAlarmaGrupoPropiedad.frx":054A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   60
      Width           =   975
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
      Left            =   900
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
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
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2580
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox txtNombre 
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
      Left            =   900
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1020
      Width           =   3615
   End
   Begin VB.Frame fraLine 
      Height          =   75
      Left            =   120
      TabIndex        =   8
      Top             =   780
      Width           =   4455
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
      Left            =   3300
      TabIndex        =   6
      Top             =   3120
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
      Left            =   1980
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
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
      TabIndex        =   2
      Top             =   1500
      Width           =   465
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre:"
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
      Width           =   600
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      Caption         =   "Datos del Grupo de Alarmas"
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
      TabIndex        =   7
      Top             =   300
      Width           =   2040
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "PersonaAlarmaGrupoPropiedad.frx":0B74
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmPersonaAlarmaGrupoPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPersonaAlarmaGrupo As PersonaAlarmaGrupo
Private mNew As Boolean


Private Sub cmdAuditoria_Click()
    frmAuditoriaGenerico.LoadDataAndShow mPersonaAlarmaGrupo
End Sub

Public Sub LoadDataAndShow(ByRef ParentForm As Form, ByRef PersonaAlarmaGrupo As PersonaAlarmaGrupo)
    Set mPersonaAlarmaGrupo = PersonaAlarmaGrupo
    mNew = (mPersonaAlarmaGrupo.IDPersonaAlarmaGrupo = 0)

    Load Me
    CSM_Forms.CenterToParent ParentForm, Me
    
    With mPersonaAlarmaGrupo
        txtNombre.Text = .Nombre
        txtNotas.Text = .Notas
        chkActivo.Value = IIf(.Activo, vbChecked, vbUnchecked)
    End With
    
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    Show
    SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtNombre.Text) = "" Then
        MsgBox "Debe ingresar el Nombre del Grupo.", vbInformation, App.Title
        txtNombre.SetFocus
        txtNombre_GotFocus
        Exit Sub
    End If
    
    With mPersonaAlarmaGrupo
        .Nombre = txtNombre.Text
        .Notas = txtNotas.Text
        .Activo = (chkActivo.Value = vbChecked)
        If mNew Then
            If Not .AddNew Then
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

Private Sub Form_Unload(Cancel As Integer)
    Set mPersonaAlarmaGrupo = Nothing
    Set frmPersonaAlarmaGrupoPropiedad = Nothing
End Sub

Private Sub txtNombre_Change()
    Caption = "Propiedades" & IIf(Trim(txtNombre.Text) = "", "", " de " & txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
    CSM_Control_TextBox.SelAllText txtNombre
End Sub

Private Sub txtNotas_GotFocus()
    CSM_Control_TextBox.SelAllText txtNotas
    cmdOK.Default = False
End Sub

Private Sub txtNotas_LostFocus()
    cmdOK.Default = True
End Sub
